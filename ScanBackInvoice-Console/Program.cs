//using BarcodeLib.BarcodeReader;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ScanBackInvoice_Console
{
    class Program
    {
        static ClassDB cls_db = new ClassDB();
        static DataTable dt_Success = new DataTable();
        static DataTable dt_Fail = new DataTable();
        static DataTable dt_NTUC_Details = new DataTable();
        static DataTable dt_InvoiceAndDate = new DataTable();
        static DataTable dt_toDelete = new DataTable();
        static bool CompleteWithNoError = true;
        static string DatabaseCompanyPrefix = ConfigurationManager.AppSettings["DatabaseCompanyPrefix"].ToString();

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Start");
                getInvoiceAndDateList();
                string optionSelection = ConfigurationManager.AppSettings["OptionSelection"].ToString();


                dt_toDelete = new DataTable();
                dt_toDelete.Clear();
                dt_toDelete.Columns.Add("DeletePath");

                if (optionSelection == "1")
                {
                    readBarcode("GENERAL"); RemoveDuplicate();
                    SavetoPDF("GENERAL");
                    ProcessFile_Scanback();
                }
                else if (optionSelection == "2")
                {
                    readBarcode("NTUC"); RemoveDuplicate();
                    SavetoPDF("NTUC");
                    #region NTUC
                    string[] combinefiles = new string[0];
                    DataRow[] dr_NTUC_Details;

                    for (int i = 0; i < dt_Success.Rows.Count; i++)
                    {
                        dr_NTUC_Details = null;
                        dr_NTUC_Details = dt_NTUC_Details.Select("InvoiceNo='" + dt_Success.Rows[i][1].ToString() + "'");

                        if (dr_NTUC_Details == null) continue;

                        combinefiles = new string[dr_NTUC_Details.Length];
                        for (int zz = 0; zz < dr_NTUC_Details.Length; zz++)
                        {
                            combinefiles[zz] = dr_NTUC_Details[zz][1].ToString();
                        }

                        string outputPdfPath = ConfigurationManager.AppSettings["outputpath2"].ToString() + dr_NTUC_Details[0][0].ToString() + ".pdf";
                        CombinePDFFiles(combinefiles, outputPdfPath);

                        foreach (string removepath in combinefiles)
                        {
                            File.Delete(removepath);
                        }

                        dt_Success.Rows[i]["Path"] = outputPdfPath;

                    }


                    #endregion
                    ProcessFile_Scanback();
                }
                else if (optionSelection == "3")
                {
                    readBarcode("GENERAL"); RemoveDuplicate();
                    SavetoPDF("GENERAL");
                    ProcessFile_Scanback();
                    //-------------------------------------------------------//
                    readBarcode("NTUC"); RemoveDuplicate();
                    SavetoPDF("NTUC");
                    #region NTUC
                    string[] combinefiles = new string[0];
                    DataRow[] dr_NTUC_Details;

                    for (int i = 0; i < dt_Success.Rows.Count; i++)
                    {
                        dr_NTUC_Details = null;
                        dr_NTUC_Details = dt_NTUC_Details.Select("InvoiceNo='" + dt_Success.Rows[i][1].ToString() + "'");

                        if (dr_NTUC_Details == null) continue;

                        combinefiles = new string[dr_NTUC_Details.Length];
                        for (int zz = 0; zz < dr_NTUC_Details.Length; zz++)
                        {
                            combinefiles[zz] = dr_NTUC_Details[zz][1].ToString();
                        }

                        string outputPdfPath = ConfigurationManager.AppSettings["outputpath2"].ToString() + dr_NTUC_Details[0][0].ToString() + ".pdf";
                        CombinePDFFiles(combinefiles, outputPdfPath);

                        foreach (string removepath in combinefiles)
                        {
                            File.Delete(removepath);
                        }

                        dt_Success.Rows[i]["Path"] = outputPdfPath;

                    }


                    #endregion
                    ProcessFile_Scanback();
                } 

                Console.WriteLine("Completed ");

                if (CompleteWithNoError)
                {
                    for (int ppp = 0; ppp < dt_toDelete.Rows.Count; ppp++)
                    {
                        File.Delete(dt_toDelete.Rows[ppp]["DeletePath"].ToString());
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR - INCOMPLETED "); 
            }


        }

        static void readBarcode(string type)
        {
            try
            {
                Console.WriteLine("Start ReadBarcode"); 
                dt_Success = new DataTable();
                dt_Fail = new DataTable();
                dt_NTUC_Details = new DataTable();

                dt_Success.Clear();
                dt_Success.Columns.Add("Path");
                dt_Success.Columns.Add("InvoiceNo");
                dt_Fail.Clear();
                dt_Fail.Columns.Add("Path");
                dt_Fail.Columns.Add("InvoiceNo");
                dt_NTUC_Details.Clear();
                dt_NTUC_Details.Columns.Add("InvoiceNo");
                dt_NTUC_Details.Columns.Add("FilePaths");

                string[] dirs_InvoiceFULL = null;

                string QRResult="";
                bool validbarcode = true; 
                string detail_invoiceNo = "";
                DataRow[] dr_InvoiceAndDate;


                #region General-SCAN
                if (type.ToUpper() == "GENERAL")
                {
                    validbarcode = true;

                    if (Directory.Exists(ConfigurationManager.AppSettings["sourcepath1"].ToString()))
                        dirs_InvoiceFULL = Directory.GetFiles(ConfigurationManager.AppSettings["sourcepath1"].ToString(), "*.tif");

                    Array.Sort(dirs_InvoiceFULL);

                    for (int i = 0; i < dirs_InvoiceFULL.Length; i++)
                    {
                        string[] sss = Spire.Barcode.BarcodeScanner.Scan(dirs_InvoiceFULL[i], Spire.Barcode.BarCodeType.Code39);
                        if (sss == null)
                        {
                            sss = Spire.Barcode.BarcodeScanner.Scan(dirs_InvoiceFULL[i], Spire.Barcode.BarCodeType.Code39);
                            if (sss == null)
                            {
                                sss = Spire.Barcode.BarcodeScanner.Scan(dirs_InvoiceFULL[i], Spire.Barcode.BarCodeType.Code39);
                                if (sss == null)
                                {
                                    validbarcode = false;
                                }
                            }
                        }

                        QRResult = sss == null || sss.Length == 0 ? QRResult : sss.First();

                        if (validbarcode)
                        {
                            dr_InvoiceAndDate = null;
                            dr_InvoiceAndDate = dt_InvoiceAndDate.Select("InvoiceNo='" + QRResult + "'");
                            if (dr_InvoiceAndDate.Length == 0)
                                validbarcode = false;
                        }

                        if (validbarcode)
                        {
                            DataRow rw_S = dt_Success.NewRow();
                            rw_S["Path"] = dirs_InvoiceFULL[i];
                            rw_S["InvoiceNo"] = QRResult;
                            dt_Success.Rows.Add(rw_S);
                        }
                        else
                        {
                            DataRow rw_F = dt_Fail.NewRow();
                            rw_F["Path"] = dirs_InvoiceFULL[i];
                            rw_F["InvoiceNo"] = QRResult;
                            dt_Fail.Rows.Add(rw_F);
                        }
                        Console.WriteLine("Processed " + i + " OVER " + dirs_InvoiceFULL.Length + "\r\n" + "Successful Scan:" + dt_Success.Rows.Count + "   Fail Scan:" + dt_Fail.Rows.Count);

                        validbarcode = true;
                    }
                }
                #endregion
                #region NTUC-SCAN
                else if (type.ToUpper() == "NTUC")
                {
                    validbarcode = true; 

                    if (Directory.Exists(ConfigurationManager.AppSettings["sourcepath2"].ToString()))
                        dirs_InvoiceFULL = Directory.GetFiles(ConfigurationManager.AppSettings["sourcepath2"].ToString(), "*.tif*");

                    Array.Sort(dirs_InvoiceFULL);

                    for (int i = 0; i < dirs_InvoiceFULL.Length; i++)
                    {
                        string[] sss = Spire.Barcode.BarcodeScanner.Scan(dirs_InvoiceFULL[i], Spire.Barcode.BarCodeType.Code39);
                        if (sss == null)
                        {
                            sss = Spire.Barcode.BarcodeScanner.Scan(dirs_InvoiceFULL[i], Spire.Barcode.BarCodeType.Code39);
                            if (sss == null)
                            {
                                sss = Spire.Barcode.BarcodeScanner.Scan(dirs_InvoiceFULL[i], Spire.Barcode.BarCodeType.Code39);
                                if (sss == null)
                                {
                                    validbarcode = false;
                                }
                            }
                        }

                        QRResult = sss == null || sss.Length==0 ? QRResult : sss.First();
                        detail_invoiceNo = QRResult;

                        if (validbarcode)
                        {
                            dr_InvoiceAndDate = null;
                            dr_InvoiceAndDate = dt_InvoiceAndDate.Select("InvoiceNo='" + QRResult + "'");
                            if (dr_InvoiceAndDate.Length == 0)
                                validbarcode = false;
                        }

                        if (validbarcode )
                        {
                            DataRow rw_S = dt_Success.NewRow();
                            rw_S["Path"] = dirs_InvoiceFULL[i];
                            rw_S["InvoiceNo"] = detail_invoiceNo;
                            dt_Success.Rows.Add(rw_S);
                        }
                        else  
                        {
                            DataRow rw_F = dt_Fail.NewRow();
                            rw_F["Path"] = dirs_InvoiceFULL[i];
                            rw_F["InvoiceNo"] = QRResult;
                            dt_Fail.Rows.Add(rw_F);
                        }

                        Console.WriteLine("Processed " + i + " OVER " + dirs_InvoiceFULL.Length + "\r\n" + "Successful Scan:" + dt_Success.Rows.Count + "   Fail Scan:" + dt_Fail.Rows.Count);

                        if (validbarcode)
                        {
                            DataRow rw_NTUC_Details = dt_NTUC_Details.NewRow();
                            rw_NTUC_Details["InvoiceNo"] = detail_invoiceNo;
                            rw_NTUC_Details["FilePaths"] = dirs_InvoiceFULL[i];
                            dt_NTUC_Details.Rows.Add(rw_NTUC_Details);
                        }
                        validbarcode = true;  
                    }

                }
                #endregion

                dt_NTUC_Details.DefaultView.Sort = "InvoiceNo ASC";
                dt_NTUC_Details = dt_NTUC_Details.DefaultView.ToTable();

                Console.WriteLine("Finish read Barcode."); 
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex+ "ERROR readBarcode.");
                CompleteWithNoError = false;
                System.Threading.Thread.Sleep(3000);
            }

        }

        static void ImageToPdf(string imagepath, string pdfpath)
        {
            try
            {
                iTextSharp.text.Rectangle pageSize = null;

                //A4 Size
                //int width = 1653;
                //int height = 2338;

                using (var srcImage = new Bitmap(imagepath.ToString()))
                {
                    pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);
                    //pageSize = new iTextSharp.text.Rectangle(0, 0, width, height);
                }
                using (var ms = new MemoryStream())
                {
                    var document = new iTextSharp.text.Document(pageSize, 0, 0, 0, 0);
                    iTextSharp.text.pdf.PdfWriter.GetInstance(document, ms).SetFullCompression();
                    document.Open();
                    var image = iTextSharp.text.Image.GetInstance(imagepath.ToString());
                    image.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                    document.Add(image);
                    document.PageCount = 1;
                    document.Close();

                    File.WriteAllBytes(pdfpath, ms.ToArray());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR ImageToPdf.");
                CompleteWithNoError = false;
                System.Threading.Thread.Sleep(3000);
            }

        }

        static void SavetoPDF(string type)
        {
            Console.WriteLine("SavetoPDF"); 
            try
            {
                String outputpath = "";
                if (type.ToUpper() == "GENERAL")
                    outputpath = ConfigurationManager.AppSettings["outputpath1"].ToString();
                else
                    outputpath = ConfigurationManager.AppSettings["outputpath2"].ToString();

                if (type.ToUpper() == "GENERAL")
                    for (int i = 0; i < dt_Success.Rows.Count; i++)
                    {
                        ImageToPdf(dt_Success.Rows[i][0].ToString(), outputpath + dt_Success.Rows[i][1].ToString() + "-" + i + ".pdf");
                        //File.Delete(dt_Success.Rows[i][0].ToString());
                        #region add to delete
                        DataRow rw_Delete = dt_toDelete.NewRow();
                        rw_Delete["DeletePath"] = dt_Success.Rows[i][0].ToString();
                        dt_toDelete.Rows.Add(rw_Delete);
                        #endregion
                        dt_Success.Rows[i]["Path"] = outputpath + dt_Success.Rows[i][1].ToString() + "-" + i + ".pdf";
                    }
                else
                    for (int i = 0; i < dt_NTUC_Details.Rows.Count; i++)
                    {
                        ImageToPdf(dt_NTUC_Details.Rows[i][1].ToString(), outputpath + dt_NTUC_Details.Rows[i][0].ToString() + "-" + i + ".pdf");
                        //File.Delete(dt_NTUC_Details.Rows[i][1].ToString());
                        #region add to delete
                        DataRow rw_Delete = dt_toDelete.NewRow();
                        rw_Delete["DeletePath"] = dt_NTUC_Details.Rows[i][1].ToString();
                        dt_toDelete.Rows.Add(rw_Delete);
                        #endregion
                        dt_NTUC_Details.Rows[i][1] = outputpath + dt_NTUC_Details.Rows[i][0].ToString() + "-" + i + ".pdf";
                    }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR SavetoPDF.");
                CompleteWithNoError = false;
                System.Threading.Thread.Sleep(3000);
            }
        }

        #region combine PDF
        static void CombinePDFFiles(string[] CombineFiles, string outputPdfPath)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage;
            //Currently only for NTUC
            //string outputPdfPath = ConfigurationManager.AppSettings["outputpath2"].ToString() + InvoiceNo+".pdf";

            sourceDocument = new Document();
            pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

            //Open the output file
            sourceDocument.Open();

            try
            {
                //Loop through the files list
                for (int f = 0; f <= CombineFiles.Length - 1; f++)
                {
                    if (CombineFiles[f] == "") continue;
                    int pages = get_pageCcount(CombineFiles[f]);

                    if (pages == 0)
                        pages = 1;

                    reader = new PdfReader(CombineFiles[f]);
                    //Add pages of current file
                    for (int i = 1; i <= pages; i++)
                    {
                        importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                        pdfCopyProvider.AddPage(importedPage);
                    }
                    reader.Close();
                }
                //At the end save the output file
                sourceDocument.Close();
            }
            catch (Exception ex)
            {
                CompleteWithNoError = false;
            }

        }

        static int get_pageCcount(string file)
        {
            using (StreamReader sr = new StreamReader(File.OpenRead(file)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());

                return matches.Count;
            }
        }
        #endregion

        static void getInvoiceAndDateList()
        {
            //string query = @" SELECT InvoiceNo,CONVERT(VARCHAR(10),InvoiceDate,102) AS INVOICEDATE  FROM ERPInvoiceStorageInfo";
            string query = @" SELECT No_ AS InvoiceNo,CONVERT(VARCHAR(10),[Shipment Date],102) AS INVOICEDATE,[Sell-to Customer No_] AS CustID ,[Route Code] AS RouteCode   FROM [" + DatabaseCompanyPrefix + "Sales Invoice Header] ";

            dt_InvoiceAndDate = cls_db.SelectQueryNoLock(query, ConfigurationManager.ConnectionStrings["DefaultNAVconn"].ConnectionString);

        }

        static void ProcessFile_Scanback()
        {
            try
            {
                string query = "";
                string custID = "";
                string routeCode = "";
                string invoicename = "";
                string path = "";
                DataRow[] dr_InvoiceAndDate;
                for (int i = 0; i < dt_Success.Rows.Count; i++)
                {
                    custID = "";
                    routeCode = "";
                    invoicename = dt_Success.Rows[i]["InvoiceNo"].ToString();
                    path = dt_Success.Rows[i]["Path"].ToString();
                    dr_InvoiceAndDate = null;
                    dr_InvoiceAndDate = dt_InvoiceAndDate.Select("InvoiceNo='" + dt_Success.Rows[i]["InvoiceNo"].ToString() + "'");

                    if (dr_InvoiceAndDate.Count() == 0) continue;
                    custID = dr_InvoiceAndDate[0]["CustID"].ToString();
                    routeCode = dr_InvoiceAndDate[0]["RouteCode"].ToString();

                    string DateSavePath = getSavePath(invoicename, dr_InvoiceAndDate[0]["INVOICEDATE"].ToString());
                    string DateSavePath2 = getSavePathBackup(invoicename, dr_InvoiceAndDate[0]["INVOICEDATE"].ToString());

                    string outfilepath1 = DateSavePath + routeCode + "_" + custID + "_" + invoicename + "_ScanBack_" + DateTime.Now.ToString("ddMMyyyy") + ".pdf";
                    string outfilepath2 = DateSavePath2 + routeCode + "_" + custID + "_" + invoicename + "_ScanBack_" + DateTime.Now.ToString("ddMMyyyy") + ".pdf";
                    query = "INSERT INTO ScanbackTable(GUID,Invoice#,CreatedDate,Path1,Path2) SELECT NEWID(),'" + invoicename + "',GETDATE() ,'" + outfilepath1 + "','" + outfilepath2 + "'";

                    File.Copy(path, outfilepath1, true);
                    File.Copy(path, outfilepath2, true);
                    File.Delete(path);
                    cls_db.ExecQueryNoLock(query, ConfigurationManager.ConnectionStrings["DefaultDBconn"].ConnectionString); 

                }
            }
            catch (Exception ex)
            {
                CompleteWithNoError = false;
                Console.WriteLine("ProcessFile_Scanback ", "ERROR");
            }
        }

        static string getSavePath(string invoiceName, string invoiceDate)
        {
            string path = "";
            try
            {
                string[] splitstring = invoiceDate.Split('.');

                if (splitstring.Length == 3)
                {
                    //path = ConfigurationManager.AppSettings["pathNewPath"].ToString() + @"OverallDate\" + splitstring[0] + @"\" + splitstring[1] + @"\" + splitstring[2] + @"\";
                    path = ConfigurationManager.AppSettings["outputpathNEW"].ToString() + splitstring[0] + @"\" + splitstring[1] + @"\" + splitstring[2] + @"\";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                }
                else
                {
                    Console.WriteLine("Please check with IT. \r\n invoice number : " + invoiceName, "ERROR");
                }
            }
            catch (Exception ex)
            {
                CompleteWithNoError = false;
                Console.WriteLine("Please check with IT. \r\n invoice number : " + invoiceName, "ERROR - getSavePath");
            }

            return path;
        }

        static string getSavePathBackup(string invoiceName, string invoiceDate)
        {
            string path = "";
            try
            {
                string[] splitstring = invoiceDate.Split('.');

                if (splitstring.Length == 3)
                {
                    //path = ConfigurationManager.AppSettings["pathNewPath"].ToString() + @"OverallDate\" + splitstring[0] + @"\" + splitstring[1] + @"\" + splitstring[2] + @"\";
                    path = ConfigurationManager.AppSettings["outputpathNEW2"].ToString() + splitstring[0] + @"\" + splitstring[1] + @"\" + splitstring[2] + @"\";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                }
                else
                {
                    Console.WriteLine("Please check with IT. \r\n invoice number : " + invoiceName, "ERROR");
                }
            }
            catch (Exception ex)
            {
                CompleteWithNoError = false;
                Console.WriteLine("Please check with IT. \r\n invoice number : " + invoiceName, "ERROR - getSavePath2");
            }

            return path;
        }

        static private void RemoveDuplicate()
        {
            try { 
            DataTable dt_temp = new DataTable();
            DataRow[] dr_temp;
            dt_temp = dt_Success.Clone();

            for (int i = 0; i < dt_Success.Rows.Count; i++)
            {
                dr_temp = null;
                dr_temp = dt_temp.Select("InvoiceNo='" + dt_Success.Rows[i]["InvoiceNo"] + "'");
                if (dr_temp.Length > 0)
                {
                    continue;
                }
                else
                {
                    DataRow rw_t = dt_temp.NewRow();
                    rw_t["Path"] = dt_Success.Rows[i]["Path"];
                    rw_t["InvoiceNo"] = dt_Success.Rows[i]["InvoiceNo"];
                    dt_temp.Rows.Add(rw_t);
                }

            }

            dt_Success = dt_temp;
            }
            catch (Exception ex)
            {
                CompleteWithNoError = false;
                Console.WriteLine("RemoveDuplicate ", "ERROR");
            }

        }


    }
}
