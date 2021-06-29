using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using TesExportImport.Model;

namespace TesExportImport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExportController : Controller
    {

        [HttpGet("ExportToExcel")]
        public async Task<DemoResponse<string>> ExportToExcel(CancellationToken cancellationToken)
        {
            string folder = "C:\\Users\\MKI\\Downloads";
            string excelName = $"UserList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
            string downloadUrl = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, excelName);
            FileInfo file = new FileInfo(Path.Combine(folder, excelName));
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(folder, excelName));
            }

            // query data from database  
            await Task.Yield();

            var list = new List<UserInfo>()
            {
                new UserInfo { UserName = "catcher", Age = 18 },
                new UserInfo { UserName = "james", Age = 20 },
            };

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(file))
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells.LoadFromCollection(list, true);
                package.Save();
            }

            return DemoResponse<string>.GetResult(0, "OK", downloadUrl);
        }

        [HttpGet("Export_2")]
        public async Task<DemoResponse<string>> Export_2([FromBody] List<KertasKerjaModel> data)
        {
            string folder = "C:\\Users\\MKI\\Downloads";
            string excelName = $"KertasKerja-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
            string downloadUrl = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, excelName);
            FileInfo file = new FileInfo(Path.Combine(folder, excelName));
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(folder, excelName));
            }

            var kerangkaKerja = data.Select(k => new KerangkaKerjaModel
            {
                aktivitas = k.aktivitas,
                bobot = k.bobot,
                bobot_fv = k.bobot_fv,
                checklist = k.checklist,
                faktor_verifikatif = k.faktor_verifikatif,
                indikator = k.indikator,
                level = k.level,
                nomor = k.nomor,
                parameter = k.parameter
            }).ToList();

            var sdm = data.Select(k => new SDMModel {
                aktivitas = k.aktivitas,
                bobot = k.bobot,
                bobot_fv = k.bobot_fv,
                checklist = k.checklist,
                faktor_verifikatif = k.faktor_verifikatif,
                indikator = k.indikator,
                level = k.level,
                nomor = k.nomor,
                parameter = k.parameter
            }).ToList();

            var proses = data.Select(k => new ProsesModel
            {
                aktivitas = k.aktivitas,
                bobot = k.bobot,
                bobot_fv = k.bobot_fv,
                checklist = k.checklist,
                faktor_verifikatif = k.faktor_verifikatif,
                indikator = k.indikator,
                level = k.level,
                nomor = k.nomor,
                parameter = k.parameter
            }).ToList();

            // query data from database  
            await Task.Yield();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(file))
            {
                var kepemimpinanSheet = package.Workbook.Worksheets.Add("Kepemimpinan");
                kepemimpinanSheet.Cells.LoadFromCollection(data, true);

                var kerangkaKerjaSheet = package.Workbook.Worksheets.Add("Kerangka Kerja");
                kerangkaKerjaSheet.Cells.LoadFromCollection(kerangkaKerja, true);

                var sdmSheet = package.Workbook.Worksheets.Add("SDM");
                sdmSheet.Cells.LoadFromCollection(sdm, true);

                var prosesSheet = package.Workbook.Worksheets.Add("Proses");
                prosesSheet.Cells.LoadFromCollection(proses, true);

                package.Save();
            }

            return DemoResponse<string>.GetResult(0, "OK", downloadUrl);
        }

        [HttpGet]
        public ActionResult<string> tesGetData()
        {
            return Ok("it works");
        }

        //    [HttpPost("import")]
        //    public ActionResult OnPostImportFromExcel()
        //    {
        //        IFormFile file = Request.Form.Files[0];
        //        //string folderName = "Upload";
        //        //string webRootPath = _hostingEnvironment.WebRootPath;
        //        //string newPath = Path.Combine(webRootPath, folderName);
        //        string newPath = "C:\\Users\\MKI\\Downloads";
        //        StringBuilder sb = new StringBuilder();
        //        if (!Directory.Exists(newPath))
        //            Directory.CreateDirectory(newPath);
        //        if (file.Length > 0)
        //        {
        //            string sFileExtension = Path.GetExtension(file.FileName).ToLower();
        //            ISheet sheet;
        //            string fullPath = Path.Combine(newPath, file.FileName);
        //            using (var stream = new FileStream(fullPath, FileMode.Create))
        //            {
        //                file.CopyTo(stream);
        //                stream.Position = 0;
        //                if (sFileExtension == ".xls")//This will read the Excel 97-2000 formats    
        //                {
        //                    HSSFWorkbook hssfwb = new HSSFWorkbook(stream);
        //                    sheet = hssfwb.GetSheetAt(0);
        //                }
        //                else //This will read 2007 Excel format    
        //                {
        //                    XSSFWorkbook hssfwb = new XSSFWorkbook(stream);
        //                    sheet = hssfwb.GetSheetAt(0);
        //                }
        //                IRow headerRow = sheet.GetRow(0);
        //                int cellCount = headerRow.LastCellNum;
        //                // Start creating the html which would be displayed in tabular format on the screen  
        //                sb.Append("<table class='table'><tr>");
        //                for (int j = 0; j < cellCount; j++)
        //                {
        //                    NPOI.SS.UserModel.ICell cell = headerRow.GetCell(j);
        //                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
        //                    sb.Append("<th>" + cell.ToString() + "</th>");
        //                }
        //                sb.Append("</tr>");
        //                sb.AppendLine("<tr>");
        //                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
        //                {
        //                    IRow row = sheet.GetRow(i);
        //                    if (row == null) continue;
        //                    if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
        //                    for (int j = row.FirstCellNum; j < cellCount; j++)
        //                    {
        //                        if (row.GetCell(j) != null)
        //                            sb.Append("<td>" + row.GetCell(j).ToString() + "</td>");
        //                    }
        //                    sb.AppendLine("</tr>");
        //                }
        //                sb.Append("</table>");
        //            }
        //        }
        //        return this.Content(sb.ToString());
        //    }
        //}

        //[HttpPost("ImportDatabase")]
        //public ActionResult OnPostImportFromExcel()
        //{
        //    IFormFile file = Request.Form.Files[0];
        //    string newPath = "C:\\Users\\MKI\\Downloads";
        //    StringBuilder sb = new StringBuilder();
        //    if (!Directory.Exists(newPath))
        //        Directory.CreateDirectory(newPath);
        //    if (file.Length > 0)
        //    {
        //        string sFileExtension = Path.GetExtension(file.FileName).ToLower();
        //        ISheet sheet;
        //        string fullPath = Path.Combine(newPath, file.FileName);
        //        using (var stream = new FileStream(fullPath, FileMode.Create))
        //        {
        //            file.CopyTo(stream);
        //            stream.Position = 0;
        //            if (sFileExtension == ".xls")//This will read the Excel 97-2000 formats    
        //            {
        //                HSSFWorkbook hssfwb = new HSSFWorkbook(stream);
        //                sheet = hssfwb.GetSheetAt(0);
        //            }
        //            else //This will read 2007 Excel format    
        //            {
        //                XSSFWorkbook hssfwb = new XSSFWorkbook(stream);
        //                sheet = hssfwb.GetSheetAt(0);
        //            }
        //            IRow headerRow = sheet.GetRow(0);
        //            int cellCount = headerRow.LastCellNum;
        //        }
        //    }
        //}
    }
}
