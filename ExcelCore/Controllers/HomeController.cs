using ClosedXML.Excel;
using ExcelCore.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Net;
using System.Text;
using System.Web;
using Microsoft.Office.Core;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;

namespace ExcelCore.Controllers
{
    public class HomeController : Controller
    {

        private DBCtx Context { get; }
        private readonly IHostingEnvironment _hostingEnvironment;

        public HomeController(IHostingEnvironment hostingEnvironment, DBCtx _context)
        {
            _hostingEnvironment = hostingEnvironment;
            Context = _context;
            
        }

        public static string ConvertTO_XLSX (FileInfo file)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var xlsFile = file.FullName;
            var wb = app.Workbooks.Open(xlsFile);
            var xlsxFile = xlsFile + "x";
            wb.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            return xlsxFile;
        }

        //Moblies
        public ActionResult File()
        {

            FileUpload model = new FileUpload();
            return View(model);
        }
        [HttpPost]
        public ActionResult File(FileUpload model)
        {
            string rootFolder = _hostingEnvironment.WebRootPath;
            string fileName = model.XlsFiles[0].FileName;
            DateTime dt = model.rbtDate;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            
            using (var stream = new MemoryStream())
            {
                model.XlsFiles[0].CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    package.SaveAs(file);
                    //save excel file in your wwwroot folder and get this excel file from wwwroot
                }
            }
            
            //After save excel file in wwwroot and then
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    //return or alert message here
                }
                else
                {
                    //read excel file data and add data in  model.StaffInfoViewModel.StaffList
                    var rowCount = worksheet.Dimension.Rows;
                    for (int row = 4; row <= rowCount; row++)
                    {
                        model.InfoModel.InfoList.Add(new Info
                        {
                            RankID = (worksheet.Cells[row, 1].Value ?? string.Empty).ToString().Trim(),
                            ResourceName = (worksheet.Cells[row, 2].Value ?? string.Empty).ToString().Trim(),
                            ResourceCode = (worksheet.Cells[row, 3].Value ?? string.Empty).ToString().Trim(),
                            ISRC = (worksheet.Cells[row, 4].Value ?? string.Empty).ToString().Trim(),
                            SPName = (worksheet.Cells[row, 5].Value ?? string.Empty).ToString().Trim(),
                            SPCode = (worksheet.Cells[row, 6].Value ?? string.Empty).ToString().Trim(),
                            Artist = (worksheet.Cells[row, 7].Value ?? string.Empty).ToString().Trim(),
                            ChargedPrice = (worksheet.Cells[row, 8].Value ?? string.Empty).ToString().Trim(),
                            DownloadTimes = (worksheet.Cells[row, 9].Value ?? string.Empty).ToString().Trim(),
                            CopyTimes = (worksheet.Cells[row, 10].Value ?? string.Empty).ToString().Trim(),
                            PresentTimes = (worksheet.Cells[row, 11].Value ?? string.Empty).ToString().Trim(),
                            RenewalTimes = (worksheet.Cells[row, 12].Value ?? string.Empty).ToString().Trim(),
                            TotalTimes = (worksheet.Cells[row, 13].Value ?? string.Empty).ToString().Trim(),
                        });


                    }

                    ////insert data into a new workbook


                    using (var workbook = new XLWorkbook())
                    {

                        var worksheet2 = workbook.Worksheets.Add("Sheet1");
                        var currentRow = 1;
                        worksheet2.Cell(currentRow, 1).Value = "Date";
                        worksheet2.Cell(currentRow, 2).Value = "RankID";
                        worksheet2.Cell(currentRow, 3).Value = "Resource Name";
                        worksheet2.Cell(currentRow, 4).Value = "Resource Code";
                        worksheet2.Cell(currentRow, 5).Value = "ISRC";
                        worksheet2.Cell(currentRow, 6).Value = "SP Name";
                        worksheet2.Cell(currentRow, 7).Value = "SP Code";
                        worksheet2.Cell(currentRow, 8).Value = "Artist";
                        worksheet2.Cell(currentRow, 9).Value = "Charged Price";
                        worksheet2.Cell(currentRow, 10).Value = "Download Times";
                        worksheet2.Cell(currentRow, 11).Value = "Copy Times";
                        worksheet2.Cell(currentRow, 12).Value = "Present Times";
                        worksheet2.Cell(currentRow, 13).Value = "Renewal Times";
                        worksheet2.Cell(currentRow, 14).Value = "Total Times";
                        if (currentRow == 1)
                        {
                            var col1 = worksheet2.Column("A");
                            var col2 = worksheet2.Column("B");
                            var col3 = worksheet2.Column("C");
                            var col4 = worksheet2.Column("D");
                            var col5 = worksheet2.Column("E");
                            var col6 = worksheet2.Column("F");
                            var col7 = worksheet2.Column("G");
                            var col8 = worksheet2.Column("H");
                            var col9 = worksheet2.Column("I");
                            var col10 = worksheet2.Column("J");
                            var col11 = worksheet2.Column("K");
                            var col12 = worksheet2.Column("L");
                            var col13 = worksheet2.Column("M");
                            var col14 = worksheet2.Column("N");


                            col1.Width = 20;
                            col2.Width = 20;
                            col3.Width = 20;
                            col4.Width = 20;
                            col6.Width = 20;
                            col7.Width = 20;
                            col8.Width = 20;
                            col9.Width = 20;
                            col10.Width = 20;
                            col11.Width = 20;
                            col12.Width = 20;
                            col13.Width = 20;
                            col14.Width = 20;
                        }
                        foreach (var item in model.InfoModel.InfoList)
                        {
                            currentRow++;
                            if (item.RankID == "" && currentRow > 3)
                            {
                                break;
                            }
                            worksheet2.Cell(currentRow, 1).Value = dt.ToString("MM/dd/yyyy");
                            worksheet2.Cell(currentRow, 2).Value = item.RankID;
                            worksheet2.Cell(currentRow, 3).Value = item.ResourceName;
                            worksheet2.Cell(currentRow, 4).Value = "'" + item.ResourceCode.ToString();
                            worksheet2.Cell(currentRow, 5).Value = item.ISRC;
                            worksheet2.Cell(currentRow, 6).Value = item.SPName;
                            worksheet2.Cell(currentRow, 7).Value = item.SPCode;
                            worksheet2.Cell(currentRow, 8).Value = item.Artist;
                            worksheet2.Cell(currentRow, 9).Value = item.ChargedPrice;
                            worksheet2.Cell(currentRow, 10).Value = item.DownloadTimes;
                            worksheet2.Cell(currentRow, 11).Value = item.CopyTimes;
                            worksheet2.Cell(currentRow, 12).Value = item.PresentTimes;
                            worksheet2.Cell(currentRow, 13).Value = item.RenewalTimes;
                            worksheet2.Cell(currentRow, 14).Value = item.TotalTimes;

                        }

                        /***************** UPLOAD TO FTP*************************/



                        var rbt = "Mobilis " + dt.ToString("MM/dd/yyyy");
                        string un = "FTPreports";
                        string pass = "FR@Arpu+19";
                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            FileInfo fileroot = new FileInfo(Path.Combine(rootFolder, rbt));
                            stream.CopyToAsync(stream);

                            byte[] fileBytes;
                            using (StreamReader fileStream = new StreamReader(stream))

                            {

                                fileBytes = Encoding.UTF8.GetBytes(fileStream.ReadToEnd());

                                fileStream.Close();

                            }
                            try

                            {

                                //Create FTP Request.

                                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://10.2.10.219/Mobilis-ALG_RBT/" + fileName);
                                request.Method = WebRequestMethods.Ftp.UploadFile;
                                request.Credentials = new NetworkCredential(un, pass);
                                request.UsePassive = true;

                                request.UseBinary = true;

                                request.EnableSsl = false;
                                using (Stream requestStream = request.GetRequestStream())

                                {

                                    requestStream.Write(content, 0, content.Length);

                                    requestStream.Close();

                                }
                                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                                response.Close();
                                ViewBag.Message = "Uploaded successfully for the day" + dt.ToString("MM/dd/yyyy");
                            }
                            catch (WebException ex)
                            {

                            }
                            ///DownloadReportForUser on local pc
                            return File(
                            content,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                rbt + ".xlsx");


                        }


                    }


                }
            }
            //return same view and  pass view model 
            return View(model);
        }


        public ActionResult OoredooTunis()
        {

            FileUpload model = new FileUpload();
            return View(model);
        }
        [HttpPost]
        public ActionResult OoredooTunis(FileUpload model)
        {
            string rootFolder = _hostingEnvironment.WebRootPath;
            string fileName = model.XlsFiles[0].FileName;
            DateTime dt = model.rbtDate;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            using (var stream = new MemoryStream())
            {
                model.XlsFiles[0].CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    package.SaveAs(file);
                    
                }
            }
            //After save excel file in wwwroot and then
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    //return or alert message here
                }
                else
                {
                    //read excel file data and add data in  model.StaffInfoViewModel.StaffList
                    var rowCount = worksheet.Dimension.Rows;
                    for (int row = 5; row <= rowCount; row++)
                    {
                        model.OoredooModel.InfoList.Add(new OInfo
                        {
                            ContentName = (worksheet.Cells[row, 2].Value ?? string.Empty).ToString().Trim(),
                            Artist = (worksheet.Cells[row, 3].Value ?? string.Empty).ToString().Trim(),
                            MediaId = (worksheet.Cells[row, 4].Value ?? string.Empty).ToString().Trim(),
                            Contentprovider = (worksheet.Cells[row, 5].Value ?? string.Empty).ToString().Trim(),
                            Trackpkey = (worksheet.Cells[row, 6].Value ?? string.Empty).ToString().Trim(),
                            CatName = (worksheet.Cells[row, 7].Value ?? string.Empty).ToString().Trim(),
                            ProviderName = (worksheet.Cells[row, 8].Value ?? string.Empty).ToString().Trim(),
                            PricePointValue = (worksheet.Cells[row, 9].Value ?? string.Empty).ToString().Trim(),

                        });


                    }

                    ////insert data into a new workbook


                    using (var workbook = new XLWorkbook())
                    {

                        var worksheet2 = workbook.Worksheets.Add("Sheet1");
                        var currentRow = 1;
                        worksheet2.Cell(currentRow, 1).Value = "Date";
                        worksheet2.Cell(currentRow, 2).Value = "Content Name";
                        worksheet2.Cell(currentRow, 3).Value = "Artist";
                        worksheet2.Cell(currentRow, 4).Value = "Media Id";
                        worksheet2.Cell(currentRow, 5).Value = "Content provider";
                        worksheet2.Cell(currentRow, 6).Value = "Trackpkey";
                        worksheet2.Cell(currentRow, 7).Value = "Cat Name";
                        worksheet2.Cell(currentRow, 8).Value = "Provider Name";
                        worksheet2.Cell(currentRow, 9).Value = "PricePoint Value";

                        if (currentRow == 1)
                        {
                            var col1 = worksheet2.Column("A");
                            var col2 = worksheet2.Column("B");
                            var col3 = worksheet2.Column("C");
                            var col4 = worksheet2.Column("D");
                            var col5 = worksheet2.Column("E");
                            var col6 = worksheet2.Column("F");
                            var col7 = worksheet2.Column("G");
                            var col8 = worksheet2.Column("H");
                            var col9 = worksheet2.Column("I");
                            


                            col1.Width = 20;
                            col2.Width = 20;
                            col3.Width = 20;
                            col4.Width = 20;
                            col5.Width = 20;
                            col6.Width = 20;
                            col7.Width = 20;
                            col8.Width = 20;
                            col9.Width = 20;
                          

                        }
                        foreach (var item in model.OoredooModel.InfoList)
                        {
                            currentRow++;
                            if (item.ContentName == "" && currentRow > 3)
                            {
                                break;
                            }
                            worksheet2.Cell(currentRow, 1).Value = dt.ToString("MM/dd/yyyy");
                            worksheet2.Cell(currentRow, 2).Value = item.ContentName;
                            worksheet2.Cell(currentRow, 3).Value = item.Artist;
                            worksheet2.Cell(currentRow, 4).Value = "'" + item.MediaId.ToString();
                            worksheet2.Cell(currentRow, 5).Value = item.Contentprovider;
                            worksheet2.Cell(currentRow, 6).Value = item.Trackpkey;
                            worksheet2.Cell(currentRow, 7).Value = item.CatName;
                            worksheet2.Cell(currentRow, 8).Value = item.ProviderName;
                            worksheet2.Cell(currentRow, 9).Value = item.PricePointValue;


                        }

                        /***************** UPLOAD TO FTP*************************/



                        var rbt = "Ooredoo TUN " + dt.ToString("MM/dd/yyyy");
                        string un = "FTPreports";
                        string pass = "FR@Arpu+19";
                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            FileInfo fileroot = new FileInfo(Path.Combine(rootFolder, rbt));
                            stream.CopyToAsync(stream);

                            byte[] fileBytes;
                            using (StreamReader fileStream = new StreamReader(stream))

                            {

                                fileBytes = Encoding.UTF8.GetBytes(fileStream.ReadToEnd());

                                fileStream.Close();

                            }
                            try

                            {
                                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://10.2.10.219/Ooredoo_TUN_RBT/" + fileName);
                                request.Method = WebRequestMethods.Ftp.UploadFile;
                                request.Credentials = new NetworkCredential(un, pass);
                                request.UsePassive = true;

                                request.UseBinary = true;

                                request.EnableSsl = false;
                                using (Stream requestStream = request.GetRequestStream())

                                {

                                    requestStream.Write(content, 0, content.Length);

                                    requestStream.Close();

                                }
                                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                                response.Close();
                                ViewBag.Message = "Uploaded successfully for the day" + dt.ToString("MM/dd/yyyy");
                            }
                            catch (WebException ex)
                            {

                            }
                            ///DownloadReportForUser on local pc
                            return File(
                            content,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                rbt + ".xlsx");


                        }



                    }
                }
            }
            
            return View(model);
        }

        //Orange
        public ActionResult OrangeEg()
        {
            FileUpload model = new FileUpload();
            return View(model);
        }
        [HttpPost]
        public ActionResult OrangeEg(FileUpload model)
        { string rootFolder = _hostingEnvironment.WebRootPath;
            string fileName = model.XlsFiles[0].FileName;
            DateTime dt = model.rbtDate;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            using (var stream = new MemoryStream())
            {
                model.XlsFiles[0].CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    package.SaveAs(file);
                    
                }
            }
            
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    //return or alert message here
                }
                else
                {
                    
                    var rowCount = worksheet.Dimension.Rows;
                    for (int row = 4; row <= rowCount - 1; row++)
                    {
                        if (row == 5004)
                        {
                            row += 6;
                        }
                        model.InfoModel.InfoList.Add(new Info
                        {
                            RankID = (worksheet.Cells[row, 1].Value ?? string.Empty).ToString().Trim(),
                            ResourceName = (worksheet.Cells[row, 2].Value ?? string.Empty).ToString().Trim(),
                            ResourceCode = (worksheet.Cells[row, 3].Value ?? string.Empty).ToString().Trim(),
                            ISRC = (worksheet.Cells[row, 4].Value ?? string.Empty).ToString().Trim(),
                            SPName = (worksheet.Cells[row, 5].Value ?? string.Empty).ToString().Trim(),
                            SPCode = (worksheet.Cells[row, 6].Value ?? string.Empty).ToString().Trim(),
                            Artist = (worksheet.Cells[row, 7].Value ?? string.Empty).ToString().Trim(),
                            ChargedPrice = (worksheet.Cells[row, 8].Value ?? string.Empty).ToString().Trim(),
                            DownloadTimes = (worksheet.Cells[row, 9].Value ?? string.Empty).ToString().Trim(),
                            CopyTimes = (worksheet.Cells[row, 10].Value ?? string.Empty).ToString().Trim(),
                            PresentTimes = (worksheet.Cells[row, 11].Value ?? string.Empty).ToString().Trim(),
                            RenewalTimes = (worksheet.Cells[row, 12].Value ?? string.Empty).ToString().Trim(),
                            TotalTimes = (worksheet.Cells[row, 13].Value ?? string.Empty).ToString().Trim(),
                        });


                    }


                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet2 = workbook.Worksheets.Add("Sheet1");
                        var currentRow = 1;
                        if (currentRow == 1)
                        {
                            var col1 = worksheet2.Column("A");
                            var col2 = worksheet2.Column("B");
                            var col3 = worksheet2.Column("C");
                            var col4 = worksheet2.Column("D");
                            var col5 = worksheet2.Column("E");
                            var col6 = worksheet2.Column("F");
                            var col7 = worksheet2.Column("G");
                            var col8 = worksheet2.Column("H");
                            var col9 = worksheet2.Column("I");
                            var col10 = worksheet2.Column("J");
                            var col11 = worksheet2.Column("K");
                            var col12 = worksheet2.Column("L");
                            var col13 = worksheet2.Column("M");
                            var col14 = worksheet2.Column("N");


                            col1.Width = 20;
                            col2.Width = 20;
                            col3.Width = 20;
                            col4.Width = 20;
                            col6.Width = 20;
                            col7.Width = 20;
                            col8.Width = 20;
                            col9.Width = 20;
                            col10.Width = 20;
                            col11.Width = 20;
                            col12.Width = 20;
                            col13.Width = 20;
                            col14.Width = 20;
                        }
                        worksheet2.Cell(currentRow, 1).Value = "Date";
                        worksheet2.Cell(currentRow, 2).Value = "RankID";
                        worksheet2.Cell(currentRow, 3).Value = "Resource Name";
                        worksheet2.Cell(currentRow, 4).Value = "Resource Code";
                        worksheet2.Cell(currentRow, 5).Value = "ISRC";
                        worksheet2.Cell(currentRow, 6).Value = "SP Name";
                        worksheet2.Cell(currentRow, 7).Value = "SP Code";
                        worksheet2.Cell(currentRow, 8).Value = "Artist";
                        worksheet2.Cell(currentRow, 9).Value = "Charged Price";
                        worksheet2.Cell(currentRow, 10).Value = "Download Times";
                        worksheet2.Cell(currentRow, 11).Value = "Copy Times";
                        worksheet2.Cell(currentRow, 12).Value = "Present Times";
                        worksheet2.Cell(currentRow, 13).Value = "Renewal Times";
                        worksheet2.Cell(currentRow, 14).Value = "Total Times";

                        foreach (var item in model.InfoModel.InfoList)
                        {
                            currentRow++;
                            if (item.RankID == "" && currentRow > 3)
                            {
                                break;
                            }
                            worksheet2.Cell(currentRow, 1).Value = dt.ToString("MM/dd/yyyy");
                            worksheet2.Cell(currentRow, 2).Value = item.RankID;
                            worksheet2.Cell(currentRow, 3).SetDataType(XLDataType.Text);
                            worksheet2.Cell(currentRow, 3).Style.NumberFormat.Format = "@";
                            worksheet2.Cell(currentRow, 3).SetValue<string>(Convert.ToString(item.ResourceName));
                            worksheet2.Cell(currentRow, 4).Value = "'" + item.ResourceCode;
                            worksheet2.Cell(currentRow, 5).Value = item.ISRC;
                            worksheet2.Cell(currentRow, 6).Value = item.SPName;
                            worksheet2.Cell(currentRow, 7).Value = item.SPCode;
                            worksheet2.Cell(currentRow, 8).Value = item.Artist;
                            worksheet2.Cell(currentRow, 9).Value = item.ChargedPrice;
                            worksheet2.Cell(currentRow, 10).Value = item.DownloadTimes;
                            worksheet2.Cell(currentRow, 11).Value = item.CopyTimes;
                            worksheet2.Cell(currentRow, 12).Value = item.PresentTimes;
                            worksheet2.Cell(currentRow, 13).Value = item.RenewalTimes;
                            worksheet2.Cell(currentRow, 14).Value = item.TotalTimes;


                        }

                        string un = "FTPreports";
                        string pass = "FR@Arpu+19";
                        var rbt = "Orange-Eg " + dt.ToString("MM/dd/yyyy");

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            FileInfo fileroot = new FileInfo(Path.Combine(rootFolder, fileName));//shoiuld chang to rbt 
                            stream.CopyToAsync(stream);

                            byte[] fileBytes;
                            using (StreamReader fileStream = new StreamReader(stream))

                            {

                                fileBytes = Encoding.UTF8.GetBytes(fileStream.ReadToEnd());

                                fileStream.Close();

                            }
                            try

                            {

                                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://10.2.10.219/Orange_Egy_Rbt/" + fileName);
                                FtpWebRequest request1 = (FtpWebRequest)WebRequest.Create("ftp://10.2.10.219/Orange-Egy-Renewal_Rbt/" + fileName);
                                request.Method = WebRequestMethods.Ftp.UploadFile;
                                request.Credentials = new NetworkCredential(un, pass);
                                request.UsePassive = true;

                                request.UseBinary = true;

                                request.EnableSsl = false;
                                using (Stream requestStream = request.GetRequestStream())

                                {

                                    requestStream.Write(content, 0, content.Length);

                                    requestStream.Close();

                                }
                                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                                response.Close();
                                ViewBag.Message = "Uploaded successfully for the day" + dt.ToString("MM/dd/yyyy");

                                    /////////////////////////////////////////////////////////////////////////
                                    
                                request1.Method = WebRequestMethods.Ftp.UploadFile;
                                request1.Credentials = new NetworkCredential(un, pass);
                                request1.UsePassive = true;

                                request1.UseBinary = true;

                                request1.EnableSsl = false;
                                using (Stream requestStream = request1.GetRequestStream())

                                {

                                    requestStream.Write(content, 0, content.Length);

                                    requestStream.Close();

                                }
                                FtpWebResponse response1 = (FtpWebResponse)request1.GetResponse();

                                response1.Close();
                                ViewBag.Message = "Uploaded successfully for the day" + dt.ToString("MM/dd/yyyy");

                            }
                            catch (WebException ex)
                            {

                            }
                            
                            return File(
                            content,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                rbt + ".xlsx");

                        }

                    }
                }
            }
          
            return View(model);
        }

        //Zain IQ
        public ActionResult ZainIQ()
        {
            FileUpload model = new FileUpload();
            return View(model);
        }
        [HttpPost]
        public ActionResult ZainIQ(FileUpload model)
        {
            string rootFolder = _hostingEnvironment.WebRootPath;
            string fileName = model.XlsFiles[0].FileName;
            DateTime dt = model.rbtDate;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            using (var stream = new MemoryStream())
            {
                model.XlsFiles[0].CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    package.SaveAs(file);

                }
            }

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    //return or alert message here
                }
                else
                {

                    var rowCount = worksheet.Dimension.Rows;
                    for (int row = 14; row <= rowCount+13; row++)
                    {
                       
                        model.ZainIQModel.InfoList.Add(new ZInfo
                        {
                            RBTID = (worksheet.Cells[row, 4].Value ?? string.Empty).ToString().Trim(),
                            RBTName = (worksheet.Cells[row, 5].Value ?? string.Empty).ToString().Trim(),
                            Artist = (worksheet.Cells[row, 8].Value ?? string.Empty).ToString().Trim(),
                            NumberofDownloads = (worksheet.Cells[row, 10].Value ?? string.Empty).ToString().Trim(),
                            NumberofPDownloads = (worksheet.Cells[row, 12].Value ?? string.Empty).ToString().Trim(),
                            TotalRevenue = (worksheet.Cells[row, 14].Value ?? string.Empty).ToString().Trim(),
                            Category = (worksheet.Cells[row, 15].Value ?? string.Empty).ToString().Trim(),
                            
                        });


                    }


                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet2 = workbook.Worksheets.Add("Sheet1");
                        var currentRow = 1;
                        if (currentRow == 1)
                        {
                            var col1 = worksheet2.Column("A");
                            var col2 = worksheet2.Column("B");
                            var col3 = worksheet2.Column("C");
                            var col4 = worksheet2.Column("D");
                            var col5 = worksheet2.Column("E");
                            var col6 = worksheet2.Column("F");
                            var col7 = worksheet2.Column("G");
                            var col8 = worksheet2.Column("H");
                            


                            col1.Width = 20;
                            col2.Width = 20;
                            col3.Width = 20;
                            col4.Width = 20;
                            col5.Width = 20;
                            col6.Width = 20;
                            col7.Width = 20;
                            col8.Width = 20;
                          

                        }
                        worksheet2.Cell(currentRow, 1).Value = "Date";
                        worksheet2.Cell(currentRow, 2).Value = "RBT ID";
                        worksheet2.Cell(currentRow, 3).Value = "RBT Name";
                        worksheet2.Cell(currentRow, 4).Value = "Artist";
                        worksheet2.Cell(currentRow, 5).Value = "Number of Downloads";
                        worksheet2.Cell(currentRow, 6).Value = "Number of Promotional Downloads";
                        worksheet2.Cell(currentRow, 7).Value = "Total Revenue Earned";
                        worksheet2.Cell(currentRow, 8).Value = "Category";
                        

                        foreach (var item in model.ZainIQModel.InfoList)
                        {
                            currentRow++;
                            if (item.RBTID == "Total")
                            {
                                break;
                            }
                            worksheet2.Cell(currentRow, 1).Value = dt.ToString("MM/dd/yyyy");
                            worksheet2.Cell(currentRow, 2).Value = item.RBTID;
                            worksheet2.Cell(currentRow, 3).Value = item.RBTName.ToString();
                            worksheet2.Cell(currentRow, 4).Value = item.Artist;
                            worksheet2.Cell(currentRow, 5).Value = item.NumberofDownloads;
                            worksheet2.Cell(currentRow, 6).Value = item.NumberofPDownloads;
                            worksheet2.Cell(currentRow, 7).Value = item.TotalRevenue;
                            worksheet2.Cell(currentRow, 8).Value = item.Category;

                        }

                        string un = "FTPreports";
                        string pass = "FR@Arpu+19";
                        var rbt = "Zain IQ " + dt.ToString("MM/dd/yyyy");

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            FileInfo fileroot = new FileInfo(Path.Combine(rootFolder, fileName));//shoiuld chang to rbt 
                            stream.CopyToAsync(stream);

                            byte[] fileBytes;
                            using (StreamReader fileStream = new StreamReader(stream))

                            {

                                fileBytes = Encoding.UTF8.GetBytes(fileStream.ReadToEnd());

                                fileStream.Close();

                            }
                            try

                            {
                              
                                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://10.2.10.219/Al_Jareed_ZainIQ_RBT/" + fileName);
                                
                                request.Method = WebRequestMethods.Ftp.UploadFile;
                                request.Credentials = new NetworkCredential(un, pass);
                                request.UsePassive = true;

                                request.UseBinary = true;

                                request.EnableSsl = false;
                                using (Stream requestStream = request.GetRequestStream())

                                {

                                    requestStream.Write(content, 0, content.Length);

                                    requestStream.Close();

                                }
                                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                                response.Close();
                                ViewBag.Message = "Uploaded successfully for the day";

                                /////////////////////////////////////////////////////////////////////////

                            }
                            catch (WebException ex)
                            {

                            }

                            return File(content,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",rbt + ".xlsx");

                        }

                    }
                }
            }

            return View(model);
        }

        //Youtube
        public ActionResult Youtube()
        {
            FileUpload model = new FileUpload();
            return View(model);
        }
        [HttpPost]
        public ActionResult Youtube(FileUpload model)
        {
            string rootFolder = _hostingEnvironment.WebRootPath;
            if (model.XlsFiles != null && model.XlsFiles.Count > 0)
            {
                foreach (IFormFile Xlsfile in model.XlsFiles)
                {
                    string fileName = model.XlsFiles.ElementAt(0).FileName;
                    string fileName2 = model.XlsFiles.ElementAt(1).FileName;

                    DateTime dt = model.rbtDate;
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
                    FileInfo file2 = new FileInfo(Path.Combine(rootFolder, fileName2));
                    using (var stream = new MemoryStream())
                    {
                        model.XlsFiles.ElementAt(0).CopyToAsync(stream);

                        using (var package = new ExcelPackage(stream))
                        {
                            package.SaveAs(file);
                            
                        }
                        using (var stream2 = new MemoryStream())
                        {
                            model.XlsFiles.ElementAt(1).CopyToAsync(stream2);
                            using (var package2 = new ExcelPackage(stream2))
                            {
                                package2.SaveAs(file2);
                            }
                        }

                    }


                    using (ExcelPackage package = new ExcelPackage(file))
                    {

                        ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();


                        if (worksheet == null)
                        {

                        }
                        else
                        {

                            var rowCount = worksheet.Dimension.Rows;

                            for (int row = 3; row <= rowCount - 1; row++)
                            {

                                model.YoutubeModel.InfoLists.Add(new YInfo
                                {
                                    Asset = (worksheet.Cells[row, 1].Value ?? string.Empty).ToString().Trim(),
                                    Yourer = (worksheet.Cells[row, 3].Value ?? string.Empty).ToString().Trim(),
                                    YourYPr = (worksheet.Cells[row, 4].Value ?? string.Empty).ToString().Trim(),
                                    Yourtr = (worksheet.Cells[row, 5].Value ?? string.Empty).ToString().Trim(),

                                });


                            }
                        }
                    }

                    using (ExcelPackage package = new ExcelPackage(file2))
                    {

                        ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();


                        if (worksheet == null)
                        {

                        }
                        else
                        {
                            
                            var rowCount = worksheet.Dimension.Rows;

                            for (int row = 3; row <= rowCount - 1; row++)
                            {

                                model.YoutubeModel.InfoLists.Add(new YInfo
                                {
                                    Asset = (worksheet.Cells[row, 1].Value ?? string.Empty).ToString().Trim(),
                                    Yourer = (worksheet.Cells[row, 3].Value ?? string.Empty).ToString().Trim(),
                                    YourYPr = (worksheet.Cells[row, 4].Value ?? string.Empty).ToString().Trim(),
                                    Yourtr = (worksheet.Cells[row, 5].Value ?? string.Empty).ToString().Trim(),

                                });


                            }
                        }
                    }


                    using (var workbook = new XLWorkbook())
                    {

                        var worksheet2 = workbook.Worksheets.Add("Sheet1");

                        var currentRow = 1;
                        if (currentRow == 1)
                        {
                            var col1 = worksheet2.Column("A");
                            var col2 = worksheet2.Column("B");
                            var col3 = worksheet2.Column("C");
                            var col4 = worksheet2.Column("D");
                            var col5 = worksheet2.Column("E");
                        


                            col1.Width = 20;
                            col2.Width = 20;
                            col3.Width = 20;
                            col4.Width = 20;
                            col5.Width = 20;
                        

                        }
                        worksheet2.Cell(currentRow, 1).Value = "Date";
                        worksheet2.Cell(currentRow, 2).Value = "Asset";
                        worksheet2.Cell(currentRow, 3).Value = "Your estimated revenue(USD)";
                        worksheet2.Cell(currentRow, 4).Value = "Your YouTube Permium revenue(USD)";
                        worksheet2.Cell(currentRow, 5).Value = "Your transaction revenue(USD)";


                        foreach (var item in model.YoutubeModel.InfoLists)
                        {
                            currentRow++;
                            if (item.Asset == "" && currentRow > 3)
                            {
                                break;
                            }
                            worksheet2.Cell(currentRow, 1).Value = dt.ToString("MM/dd/yyyy");
                            worksheet2.Cell(currentRow, 2).Value = item.Asset;
                            worksheet2.Cell(currentRow, 3).Value = item.Yourer;
                            worksheet2.Cell(currentRow, 4).Value = item.YourYPr;
                            worksheet2.Cell(currentRow, 5).Value = item.Yourtr;


                        }


                        var rbt = "YouTube " + dt.ToString("MM/dd/yyyy");

                        string un = "FTPreports";
                        string pass = "FR@Arpu+19";
                      
                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            FileInfo fileroot = new FileInfo(Path.Combine(rootFolder, Xlsfile.FileName));//shoiuld chang to rbt 
                            stream.CopyToAsync(stream);

                            byte[] fileBytes;
                            using (StreamReader fileStream = new StreamReader(stream))

                            {

                                fileBytes = Encoding.UTF8.GetBytes(fileStream.ReadToEnd());

                                fileStream.Close();

                            }
                            try

                            {

                                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://10.2.10.219/YouTube_WorldWide/" + fileName);

                                request.Method = WebRequestMethods.Ftp.UploadFile;
                                request.Credentials = new NetworkCredential(un, pass);
                                request.UsePassive = true;

                                request.UseBinary = true;

                                request.EnableSsl = false;
                                using (Stream requestStream = request.GetRequestStream())

                                {

                                    requestStream.Write(content, 0, content.Length);

                                    requestStream.Close();

                                }
                                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                                response.Close();
                                ViewBag.Message = "Uploaded successfully for the day" + dt.ToString("MM/dd/yyyy");


                            }
                            catch (WebException ex)
                            {

                            }
                            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", rbt + ".xlsx");
                        }
                    }

                }
            }

            return View(model);
        }

        //Djezzy
        public ActionResult Djezzy()
        {
            FileUpload model = new FileUpload();
            return View(model);
        }
        [HttpPost]
        public IActionResult Djezzy(FileUpload model)

        {
            string rootFolder = _hostingEnvironment.WebRootPath;
            if (model.XlsFiles != null && model.XlsFiles.Count > 0)
            {
                foreach (IFormFile Xlsfile in model.XlsFiles)
                {
                    var fileName = model.XlsFiles.ElementAt(1).FileName;
                    string fileName2 = model.XlsFiles.ElementAt(0).FileName;

                    DateTime dt = model.rbtDate;
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
                    FileInfo file2 = new FileInfo(Path.Combine(rootFolder, fileName2));


                    using (var stream = new MemoryStream())
                    {
                        model.XlsFiles.ElementAt(1).CopyToAsync(stream);

                        using (var package = new ExcelPackage(stream))
                        {
                            package.SaveAs(file);

                        }

                        using (var stream2 = new MemoryStream())
                        {
                            model.XlsFiles.ElementAt(0).CopyToAsync(stream2);

                            using (var package2 = new ExcelPackage(stream2))
                            {
                                package2.SaveAs(file2);
                            }
                        }
                    }


                    using (ExcelPackage package = new ExcelPackage(file))
                    {

                        ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();


                        if (worksheet == null)
                        {

                        }
                        else
                        {

                            var rowCount = worksheet.Dimension.Rows;

                            for (int row =4 ; row <= rowCount-2 ; row++)
                            {

                                model.DjezzyModel.InfoLists.Add(new DInfo
                                {
                                    RankID = (worksheet.Cells[row, 1].Value ?? string.Empty).ToString().Trim(),
                                    ResourceName = (worksheet.Cells[row, 2].Value ?? string.Empty).ToString().Trim(),
                                    ResourceCode = (worksheet.Cells[row, 3].Value ?? string.Empty).ToString().Trim(),
                                    ISRC = (worksheet.Cells[row, 4].Value ?? string.Empty).ToString().Trim(),
                                    SPName = (worksheet.Cells[row, 5].Value ?? string.Empty).ToString().Trim(),
                                    SPCode = (worksheet.Cells[row, 6].Value ?? string.Empty).ToString().Trim(),
                                    Artist = (worksheet.Cells[row, 7].Value ?? string.Empty).ToString().Trim(),
                                    ChargedPrice = (worksheet.Cells[row, 8].Value ?? string.Empty).ToString().Trim(),
                                    DownloadTimes = (worksheet.Cells[row, 9].Value ?? string.Empty).ToString().Trim(),
                                    CopyTimes = (worksheet.Cells[row, 10].Value ?? string.Empty).ToString().Trim(),
                                    PresentTimes = (worksheet.Cells[row, 11].Value ?? string.Empty).ToString().Trim(),
                                    RenewalTimes = (worksheet.Cells[row, 12].Value ?? string.Empty).ToString().Trim(),
                                    TotalTimes = (worksheet.Cells[row, 13].Value ?? string.Empty).ToString().Trim(),
                                });
                                

                            }
                            
                        }
                    }

                    using (ExcelPackage package = new ExcelPackage(file2))
                    {

                        ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();


                        if (worksheet == null)
                        {

                        }
                        else
                        {
                            var rowCount = worksheet.Dimension.Rows;

                            for (int row = 4; row <= rowCount - 2; row++)
                            {

                                model.DjezzyModel.InfoLists.Add(new DInfo
                                {
                                    RankID = (worksheet.Cells[row, 1].Value ?? string.Empty).ToString().Trim(),
                                    ResourceName = (worksheet.Cells[row, 2].Value ?? string.Empty).ToString().Trim(),
                                    ResourceCode = (worksheet.Cells[row, 3].Value ?? string.Empty).ToString().Trim(),
                                    ISRC = (worksheet.Cells[row, 4].Value ?? string.Empty).ToString().Trim(),
                                    SPName = (worksheet.Cells[row, 5].Value ?? string.Empty).ToString().Trim(),
                                    SPCode = (worksheet.Cells[row, 6].Value ?? string.Empty).ToString().Trim(),
                                    Artist = (worksheet.Cells[row, 7].Value ?? string.Empty).ToString().Trim(),
                                    ChargedPrice = (worksheet.Cells[row, 8].Value ?? string.Empty).ToString().Trim(),
                                    DownloadTimes = (worksheet.Cells[row, 9].Value ?? string.Empty).ToString().Trim(),
                                    CopyTimes = (worksheet.Cells[row, 10].Value ?? string.Empty).ToString().Trim(),
                                    PresentTimes = (worksheet.Cells[row, 11].Value ?? string.Empty).ToString().Trim(),
                                    RenewalTimes = (worksheet.Cells[row, 12].Value ?? string.Empty).ToString().Trim(),
                                    TotalTimes = (worksheet.Cells[row, 13].Value ?? string.Empty).ToString().Trim(),

                                });

                                
                            }
                        }
                    }

                    using (var workbook = new XLWorkbook())
                    {

                        var worksheet2 = workbook.Worksheets.Add("Sheet1");

                        var currentRow = 1;
                        if (currentRow == 1)
                        {
                            var col1 = worksheet2.Column("A");
                            var col2 = worksheet2.Column("B");
                            var col3 = worksheet2.Column("C");
                            var col4 = worksheet2.Column("D");
                            var col5 = worksheet2.Column("E");
                            var col6 = worksheet2.Column("F");
                            var col7 = worksheet2.Column("G");
                            var col8 = worksheet2.Column("H");
                            var col9 = worksheet2.Column("I");
                            var col10 = worksheet2.Column("J");
                            var col11 = worksheet2.Column("K");
                            var col12 = worksheet2.Column("L");
                            var col13 = worksheet2.Column("M");
                            var col14 = worksheet2.Column("N");


                            col1.Width = 20;
                            col2.Width = 20;
                            col3.Width = 20;
                            col4.Width = 20;
                            col6.Width = 20;
                            col7.Width = 20;
                            col8.Width = 20;
                            col9.Width = 20;
                            col10.Width = 20;
                            col11.Width = 20;
                            col12.Width = 20;
                            col13.Width = 20;
                            col14.Width = 20;

                        }
                        worksheet2.Cell(currentRow, 1).Value = "Date";
                        worksheet2.Cell(currentRow, 2).Value = "RankID";
                        worksheet2.Cell(currentRow, 3).Value = "Resource Name";
                        worksheet2.Cell(currentRow, 4).Value = "Resource Code";
                        worksheet2.Cell(currentRow, 5).Value = "ISRC";
                        worksheet2.Cell(currentRow, 6).Value = "SP Name";
                        worksheet2.Cell(currentRow, 7).Value = "SP Code";
                        worksheet2.Cell(currentRow, 8).Value = "Artist";
                        worksheet2.Cell(currentRow, 9).Value = "Charged Price";
                        worksheet2.Cell(currentRow, 10).Value = "Download Times";
                        worksheet2.Cell(currentRow, 11).Value = "Copy Times";
                        worksheet2.Cell(currentRow, 12).Value = "Present Times";
                        worksheet2.Cell(currentRow, 13).Value = "Renewal Times";
                        worksheet2.Cell(currentRow, 14).Value = "Total Times";


                        foreach (var item in model.DjezzyModel.InfoLists)
                        {
                            currentRow++;
                            if (item.RankID == "" && currentRow > 3)
                            {
                                break;
                            }
                            worksheet2.Cell(currentRow, 1).Value = dt.ToString("MM/dd/yyyy");
                            worksheet2.Cell(currentRow, 2).Value = item.RankID;
                            worksheet2.Cell(currentRow, 3).Value = item.ResourceName;
                            worksheet2.Cell(currentRow, 4).Value = "'" + item.ResourceCode;
                            worksheet2.Cell(currentRow, 5).Value = item.ISRC;
                            worksheet2.Cell(currentRow, 6).Value = item.SPName;
                            worksheet2.Cell(currentRow, 7).Value = item.SPCode;
                            worksheet2.Cell(currentRow, 8).Value = item.Artist;
                            worksheet2.Cell(currentRow, 9).Value = item.ChargedPrice;
                            worksheet2.Cell(currentRow, 10).Value = item.DownloadTimes;
                            worksheet2.Cell(currentRow, 11).Value = item.CopyTimes;
                            worksheet2.Cell(currentRow, 12).Value = item.PresentTimes;
                            worksheet2.Cell(currentRow, 13).Value = item.RenewalTimes;
                            worksheet2.Cell(currentRow, 14).Value = item.TotalTimes;


                        }

                        string un = "FTPreports";
                        string pass = "FR@Arpu+19";
                        var rbt = "Djezzy " + dt.ToString("MM/dd/yyyy");

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            FileInfo fileroot = new FileInfo(Path.Combine(rootFolder, Xlsfile.FileName));//shoiuld chang to rbt 
                            stream.CopyToAsync(stream);

                            byte[] fileBytes;
                            using (StreamReader fileStream = new StreamReader(stream))

                            {

                                fileBytes = Encoding.UTF8.GetBytes(fileStream.ReadToEnd());

                                fileStream.Close();

                            }
                            try

                            {

                                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://10.2.10.219/Djezzy_ALG_Rbt/" + fileName);

                                request.Method = WebRequestMethods.Ftp.UploadFile;
                                request.Credentials = new NetworkCredential(un, pass);
                                request.UsePassive = true;

                                request.UseBinary = true;

                                request.EnableSsl = false;
                                using (Stream requestStream = request.GetRequestStream())

                                {

                                    requestStream.Write(content, 0, content.Length);

                                    requestStream.Close();

                                }
                                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                                //Console.WriteLine("Upload File Complete, status {0}", response.StatusDescription);
                                Response.WriteAsJsonAsync("<script>alert('Data inserted successfully')</script>");

                                response.Close();
                                ViewBag.Message = "Uploaded successfully for the day" + dt.ToString("MM/dd/yyyy");
                                return View();

                            }
                            catch (WebException ex)
                            {
                                //var resp = new StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                                //Console.WriteLine(resp);
                                //Console.ReadKey();
                             
                                    Console.WriteLine(ex.Message);
                                
                            }
                            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", rbt + ".xlsx");
                            
                        }
                    }

                }
            }
            return View(model);
        }

        public string FtpTest(string path)
        {
            string un = "FTPreports";
            string pass = "FR@Arpu+19";

            try
            {

                return "hi";
            }
            catch
            {
                return "Error!";
            }

        }
    }
}
