using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System;
using WebApplicationLab6.Data;
using WebApplicationLab6.Objects;
using OfficeOpenXml;
using Microsoft.AspNetCore.Authorization;
using Microsoft.EntityFrameworkCore;
using System.Linq;
using System.Security.Policy;

namespace WebApplicationLab6.Controllers
{
    public class ReportController : Controller
    {
     

        private readonly ApplicationDbContext _context;

        public ReportController(ApplicationDbContext context)
        {
            _context = context;
        }

        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ")]
        public IActionResult Index()
        {
            var cities = _context.Cities.ToList();

            var report = new Report();

            foreach (var c in cities)
            {
                var contractCities = _context.ContractCities.Where(x => x.CityId == c.Id).ToList();
                foreach (var contractCity in contractCities)
                {
                    var contract = _context.Contracts.Find(contractCity.ContractId);
                    var countDone = _context.CaptureActs.Include(x => x.Claim).Where(x => x.ContractId == contract.Id).Select(x => x.Claim).Where(x => x.IsDone).Count();
                    int sum = 0;
                    var acts = _context.CaptureActs.Include(x => x.Claim).Where(x => x.ContractId == contract.Id).ToList();
                    foreach (var act in acts)
                    {
                        sum = act.CountCats + act.CountDogs + sum;
                    }
                    report.Rows.Add(new ReportRow
                    {
                        ContractId = contract.Id,
                        CountClaimsDone = countDone,
                        CountAnimals = sum,
                        City = c.Name,
                        Decimal = contractCity.Price * sum
                    });
                }
            }

            return View(report.Rows);
        }

        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ")]
        public IActionResult Download()
        {
            var cities = _context.Cities.ToList();

            var report = new Report();
            
            foreach(var c in cities)
            {
                var contractCities = _context.ContractCities.Where(x => x.CityId == c.Id).ToList();
                foreach(var contractCity in contractCities)
                {
                    var contract = _context.Contracts.Find(contractCity.ContractId);
                    var countDone = _context.CaptureActs.Include(x => x.Claim).Where(x => x.ContractId == contract.Id).Select(x => x.Claim).Where(x => x.IsDone).Count();
                    int sum = 0;
                    var acts = _context.CaptureActs.Include(x => x.Claim).Where(x => x.ContractId == contract.Id).ToList();
                    foreach(var act in acts)
                    {
                        sum = act.CountCats + act.CountDogs + sum;
                    }
                    report.Rows.Add(new ReportRow
                    {
                        ContractId = contract.Id,
                        CountClaimsDone = countDone,
                        CountAnimals = sum,
                        City = c.Name,
                        Decimal = contractCity.Price * sum
                    });
                }
            }

            // город - сколько закрыто заявок - сколько отловленно - общая стоимость 

            // Create a new Excel package
            using (ExcelPackage package = new ExcelPackage())
            {
                // Create the worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Contracts");

                // Add column headers
                worksheet.Cells[1, 1].Value = "Контракт";
                worksheet.Cells[1, 2].Value = "Город";
                worksheet.Cells[1, 3].Value = "Сколько Закрыто заявок";
                worksheet.Cells[1, 4].Value = "Сколько отловленно";
                worksheet.Cells[1, 5].Value = "Общая стоимость";

                // Add data to the worksheet
                int row = 2;
                foreach (var item in report.Rows)
                {
                    worksheet.Cells[row, 1].Value = item.ContractId;
                    worksheet.Cells[row, 2].Value = item.City;
                    worksheet.Cells[row, 3].Value = item.CountClaimsDone;
                    worksheet.Cells[row, 4].Value = item.CountAnimals;
                    worksheet.Cells[row, 5].Value = item.Decimal;
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
