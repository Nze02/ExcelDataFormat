using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelDataFormat.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportController : ControllerBase
    {
        private IWebHostEnvironment _hostEnvironment;

        public ReportController(IWebHostEnvironment hostEnvironment)
        {
            _hostEnvironment = hostEnvironment;
        }


        [HttpGet]
        public async Task<IActionResult> GetStaffDetailsInExcel()
        {
            //Declare and initialize an empty list
            List<StaffDetails> staffDetails = new List<StaffDetails>{
                new StaffDetails{Id = 1, FirstName = "Jim", LastName = "John", Gender = "Male", Phone = "123456", Address = "Cape Town", Department = "Software Engineering"},
                new StaffDetails{Id = 2, FirstName = "Jane", LastName = "Doe", Gender = "Female", Phone = "9876543", Address = "New York", Department = "Q & A"},
                new StaffDetails{Id = 3, FirstName = "Jeff", LastName = "Peters", Gender = "Male", Phone = "4567234", Address = "Lagos", Department = "HR" },
                new StaffDetails{Id = 4, FirstName = "Jenny", LastName = "Paul", Gender = "Female", Phone = "87222134", Address = "Abuja", Department = "Software Engineering"},
                new StaffDetails{Id = 5, FirstName = "Sophia", LastName = "Shalom", Gender = "Female", Phone = "65775432", Address = "Port-Harcourt", Department = "Training"} };


            string fileName = "\\StaffDetails.xlsx";
            string reportFullName = _hostEnvironment.ContentRootPath + fileName;


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage xlPackage = new ExcelPackage())
            {
                xlPackage.Workbook.Worksheets.Add("Sheet 1").Cells[1,1].LoadFromCollection(staffDetails, true);
                xlPackage.SaveAs(new FileInfo(reportFullName));
            }
            return Ok(reportFullName);
            
        }


        [HttpGet("DownloadReport/{FullFilePath}")]
        public FileContentResult DownloadReport(string FullFilePath)
        {

            try
            {
                var data = System.IO.File.ReadAllBytes(FullFilePath);
                var result = new FileContentResult(data, "application/octet-stream")
                {
                    FileDownloadName = FullFilePath
                };
                return result;

            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}
