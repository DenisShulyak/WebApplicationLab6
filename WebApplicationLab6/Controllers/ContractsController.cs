using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using WebApplicationLab6.Data;
using WebApplicationLab6.Objects;
using Contract = WebApplicationLab6.Objects.Contract;

namespace WebApplicationLab6.Controllers
{
    public class ContractsController : Controller
    {
        private readonly ApplicationDbContext _context;

        private static List<Contract> _lastList = new List<Contract>();

        public ContractsController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: Contracts
        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ,Оператор ВетСлужбы,Куратор ВетСлужбы,Подписант ВетСлужбы,Куратор по отлову,Куратор приюта,Подписант приюта")]
        public async Task<IActionResult> Index(string searchField, string searchTerm, string sortField, string sortOrder, string downloadData)
        {
            if (!string.IsNullOrEmpty(downloadData) && downloadData.ToLower() == "true")
            {
                return Download(_lastList);
            }
            var applicationDbContext = _context.Contracts.Include(c => c.Customer).Include(c => c.Executor);
            var contracts = applicationDbContext.ToList();
            if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                contracts = contracts.Where(x => x.ExecutorId == user.OrganizationId).ToList();
            }
            else if(User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                contracts = contracts.Where(x => x.CustomerId == user.OrganizationId).ToList();
            }

            if (!string.IsNullOrEmpty(searchField) && !string.IsNullOrEmpty(searchTerm))
            {
                switch (searchField)
                {
                    case "ContractId":
                        contracts = contracts.Where(x => x.Id == Guid.Parse(searchTerm)).ToList();
                        break;
                    case "ExecutorName":
                        contracts = contracts.Where(x => x.Executor.Name.Contains(searchTerm)).ToList();
                        break;
                    case "CustomerName":
                        contracts = contracts.Where(x => x.Customer.Name.Contains(searchTerm)).ToList();
                        break;
                    default:
                        // Поле для поиска не выбрано, игнорируем поиск
                        break;
                }
            }

            if (!string.IsNullOrEmpty(sortField))
            {
                if (sortField == "BeginDate")
                {
                    contracts = sortOrder == "Ascending" ? contracts.OrderBy(c => c.BeginDate).ToList() : contracts.OrderByDescending(c => c.BeginDate).ToList();
                }
                else if (sortField == "EndDate")
                {
                    contracts = sortOrder == "Ascending" ? contracts.OrderBy(c => c.EndDate).ToList() : contracts.OrderByDescending(c => c.EndDate).ToList();
                }
                else if (sortField == "Customer")
                {
                    contracts = sortOrder == "Ascending" ? contracts.OrderBy(c => c.CustomerId).ToList() : contracts.OrderByDescending(c => c.CustomerId).ToList();
                }
                else if (sortField == "Executor")
                {
                    contracts = sortOrder == "Ascending" ? contracts.OrderBy(c => c.ExecutorId).ToList() : contracts.OrderByDescending(c => c.ExecutorId).ToList();
                }
            }

            ViewBag.SearchField = searchField;
            ViewBag.SearchTerm = searchTerm;

            _lastList = contracts;

            return View(contracts);
        }

        private IActionResult Download(List<Contract> models)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                // Create the worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Contracts");

                // Add column headers
                worksheet.Cells[1, 1].Value = "BeginDate";
                worksheet.Cells[1, 2].Value = "EndDate";
                worksheet.Cells[1, 3].Value = "Customer";
                worksheet.Cells[1, 4].Value = "Executor";

                // Add data to the worksheet
                int row = 2;
                foreach (var item in models)
                {
                    worksheet.Cells[row, 1].Value = item.BeginDate.ToString();
                    worksheet.Cells[row, 2].Value = item.EndDate.ToString();
                    worksheet.Cells[row, 3].Value = item.Customer.Name;
                    worksheet.Cells[row, 4].Value = item.Executor.Name;
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

        // GET: Contracts/Details/5
        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ,Оператор ВетСлужбы,Куратор ВетСлужбы,Подписант ВетСлужбы,Куратор по отлову,Куратор приюта,Подписант приюта")]
        public async Task<IActionResult> Details(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var contracts = _context.Contracts
                .Include(c => c.Customer)
                .Include(c => c.Executor)
                .Where(m => m.Id == id).ToList();

            if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                contracts = contracts.Where(x => x.ExecutorId == user.OrganizationId).ToList();
            }
            else if (User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                contracts = contracts.Where(x => x.CustomerId == user.OrganizationId).ToList();
            }

            var contract = contracts.FirstOrDefault();

            if (contract == null)
            {
                return NotFound();
            }

            return View(contract);
        }

        // GET: Contracts/Create
        [Authorize(Roles = "Оператор ОМСУ")]
        public IActionResult Create()
        {
            var customers = _context.Organizations.ToList();
            var excutors = _context.Organizations.ToList();
            if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                excutors = excutors.Where(x => x.Id == user.OrganizationId).ToList();
            }
            else if (User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                customers = customers.Where(x => x.Id == user.OrganizationId).ToList();
            }
            ViewBag.Customers = new SelectList(customers, "Id", "Name");
            ViewBag.Executors = new SelectList(excutors, "Id", "Name");
            return View();
        }

        // POST: Contracts/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [Authorize(Roles = "Оператор ОМСУ")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("BeginDate,EndDate,CustomerId,ExecutorId,Id")] Contract contract)
        {
            var customers = _context.Organizations.ToList();
            var excutors = _context.Organizations.ToList();
            if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                excutors = excutors.Where(x => x.Id == user.OrganizationId).ToList();
            }
            else if (User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                customers = customers.Where(x => x.Id == user.OrganizationId).ToList();
            }
            if (ModelState.IsValid)
            {
                contract.Id = Guid.NewGuid();
                _context.Add(contract);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            ViewBag.Customers = new SelectList(customers, "Id", "Name", contract.CustomerId);
            ViewBag.Executors = new SelectList(excutors, "Id", "Name", contract.ExecutorId);
            return View(contract);
        }

        // GET: Contracts/Edit/5
        [Authorize(Roles = "Оператор ОМСУ")]
        public async Task<IActionResult> Edit(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var contracts = _context.Contracts.Where(x=>x.Id == id).ToList();
            var customers = _context.Organizations.ToList();
            var excutors = _context.Organizations.ToList();
            if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                excutors = excutors.Where(x => x.Id == user.OrganizationId).ToList();
                contracts = contracts.Where(x => x.ExecutorId == user.OrganizationId).ToList();
            }
            else if (User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                customers = customers.Where(x => x.Id == user.OrganizationId).ToList();
                contracts = contracts.Where(x => x.CustomerId == user.OrganizationId).ToList();
            }
            var contract = contracts.FirstOrDefault();
            if (contract == null)
            {
                return NotFound();
            }
            ViewBag.Customers = new SelectList(customers, "Id", "Name", contract.CustomerId);
            ViewBag.Executors = new SelectList(excutors, "Id", "Name", contract.ExecutorId);
            return View(contract);
        }

        // POST: Contracts/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [Authorize(Roles = "Оператор ОМСУ")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(Guid id, [Bind("BeginDate,EndDate,CustomerId,ExecutorId,Id")] Contract contract)
        {
            if (id != contract.Id)
            {
                return NotFound();
            }
            var customers = _context.Organizations.ToList();
            var excutors = _context.Organizations.ToList();
            if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                excutors = excutors.Where(x => x.Id == user.OrganizationId).ToList();
            }
            else if (User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                customers = customers.Where(x => x.Id == user.OrganizationId).ToList();
            }
            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(contract);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!ContractExists(contract.Id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            ViewBag.Customers = new SelectList(customers, "Id", "Name", contract.CustomerId);
            ViewBag.Executors = new SelectList(excutors, "Id", "Name", contract.ExecutorId);
            return View(contract);
        }

        // GET: Contracts/Delete/5
        [Authorize(Roles = "Оператор ОМСУ")]
        public async Task<IActionResult> Delete(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var contracts = _context.Contracts
                .Include(c => c.Customer)
                .Include(c => c.Executor)
                .Where(m => m.Id == id).ToList();

            if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта") || User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                contracts = contracts.Where(x => x.ExecutorId == user.OrganizationId).ToList();
            }
            else if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта") || User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                contracts = contracts.Where(x => x.CustomerId == user.OrganizationId).ToList();
            }
            var contract = contracts.FirstOrDefault();
            if (contract == null)
            {
                return NotFound();
            }

            return View(contract);
        }

        // POST: Contracts/Delete/5
        [Authorize(Roles = "Оператор ОМСУ")]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(Guid id)
        {

            var contracts = _context.Contracts
                .Include(c => c.Customer)
                .Include(c => c.Executor)
                .Where(m => m.Id == id).ToList();

            if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта") || User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                contracts = contracts.Where(x => x.ExecutorId == user.OrganizationId).ToList();
            }
            else if (User.IsInRole("Куратор приюта") || User.IsInRole("Подписан приюта") || User.IsInRole("Оператор приюта") || User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                contracts = contracts.Where(x => x.CustomerId == user.OrganizationId).ToList();
            }
            var contract = contracts.FirstOrDefault();
            if (contract == null)
            {
                return NotFound();
            }
            _context.Contracts.Remove(contract);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool ContractExists(Guid id)
        {
            return _context.Contracts.Any(e => e.Id == id);
        }
    }
}
