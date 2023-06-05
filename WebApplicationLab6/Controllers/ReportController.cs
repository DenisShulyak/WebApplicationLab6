using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System;
using WebApplicationLab6.Data;
using WebApplicationLab6.Objects;
using OfficeOpenXml;

namespace WebApplicationLab6.Controllers
{
    public class ReportController : Controller
    {
     

        private readonly ApplicationDbContext _context;

        public ReportController(ApplicationDbContext context)
        {
            _context = context;
        }
        public IActionResult Download()
        {
            // Get data from the database or any other source
            IEnumerable<Contract> contracts = _context.Contracts;

            // Create a new Excel package
            using (ExcelPackage package = new ExcelPackage())
            {
                // Create the worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Contracts");

                // Add column headers
                worksheet.Cells[1, 1].Value = "Begin Date";
                worksheet.Cells[1, 2].Value = "End Date";
                worksheet.Cells[1, 3].Value = "Customer";
                worksheet.Cells[1, 4].Value = "Executor";

                // Add data to the worksheet
                int row = 2;
                foreach (var contract in contracts)
                {
                    worksheet.Cells[row, 1].Value = contract.BeginDate.ToString("MM/dd/yyyy");
                    worksheet.Cells[row, 2].Value = contract.EndDate.ToString("MM/dd/yyyy");
                    worksheet.Cells[row, 3].Value = contract.CustomerId;
                    worksheet.Cells[row, 4].Value = contract.ExecutorId;
                    row++;
                }

                // Auto-fit columns for better visibility
                worksheet.Cells.AutoFitColumns();

                // Convert the Excel package to a byte array
                byte[] excelBytes = package.GetAsByteArray();

                // Set the content type and file name for the response
                string fileName = "contracts.xlsx";
                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                return File(excelBytes, contentType, fileName);
            }
        }

        private IEnumerable<Contract> GetContractsFromDatabase()
        {
            // Implement your logic to retrieve contracts from the database
            // and return them as a collection
            throw new NotImplementedException();
        }
    }
}
