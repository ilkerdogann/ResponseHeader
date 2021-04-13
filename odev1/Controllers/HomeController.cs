using Abis.Web.Core.ActionResults;
using odev1.Models;
using OfficeOpenXml;
using Rotativa;
using System.Web.Mvc;

namespace odev1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Download(StudentModel model)
        {
            if (model.DownloadType == 1)
            {
                return new WordResult("Download") { FileName = $"{model.No.ToString()}_{model.Name}_{model.Surname}", Model = model };
            }
            if (model.DownloadType == 2)
            {
                return new ViewAsPdf("Download", model);
            }
            if (model.DownloadType == 3)
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excel = new ExcelPackage())
                {
                    var workSheet = excel.Workbook.Worksheets.Add($"{model.No.ToString()}_{model.Name}_{model.Surname}");
                    workSheet.Row(1).Style.Font.Bold = true;
                    workSheet.Cells[1, 1].Value = "Öğrenci No";
                    workSheet.Cells[1, 2].Value = "Öğrenci Adı";
                    workSheet.Cells[1, 3].Value = "Öğrenci Soyadı";
                    int recordIndex = 2;
                    workSheet.Cells[recordIndex, 1].Value = model.No;
                    workSheet.Cells[recordIndex, 2].Value = model.Name;
                    workSheet.Cells[recordIndex, 3].Value = model.Surname;
                    recordIndex++;
                    workSheet.Cells.AutoFitColumns();
                    var result = excel.GetAsByteArray();
                    var fileName = $"{model.No.ToString()}_{model.Name}_{model.Surname}.xlsx";
                    return File(result, "application/vnd.ms-excel", fileName);
                }
            }
            return View(model);
        }
    }
}