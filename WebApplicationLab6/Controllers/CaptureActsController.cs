using System;
using System.Collections.Generic;
using System.Composition;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using WebApplicationLab6.Data;
using WebApplicationLab6.Objects;

namespace WebApplicationLab6.Controllers
{
    public class CaptureActsController : Controller
    {
        private readonly ApplicationDbContext _context;

        private static List<CaptureAct> _lastList = new List<CaptureAct>();

        public CaptureActsController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: CaptureActs
        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ,Оператор ВетСлужбы,Куратор ВетСлужбы,Подписант ВетСлужбы,Оператор по отлову,Куратор по отлову,Подписант по отлову,Оператор приюта,Куратор приюта,Подписант приюта")]
        public async Task<IActionResult> Index(string searchField, string searchTerm, string sortField, string sortOrder, string downloadData)
        {
            if (!string.IsNullOrEmpty(downloadData) && downloadData.ToLower() == "true")
            {
                return Download(_lastList);
            }
            var applicationDbContext = _context.CaptureActs.Include(c => c.Claim).Include(c => c.Contract).Include(c => c.Organization);
            var captureActs = applicationDbContext.ToList();
            if (User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ") || User.IsInRole("Оператор по отлову")
               || User.IsInRole("Куратор по отлову") || User.IsInRole("Подписант  по отлову"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                captureActs = captureActs.Where(x => x.OrganizationId == user.OrganizationId).ToList();
            }
            if (!string.IsNullOrEmpty(searchField) && !string.IsNullOrEmpty(searchTerm))
            {
                switch (searchField)
                {
                    case "CapturePurpose":
                        captureActs = captureActs.Where(act => act.CapturePurpose.Contains(searchTerm)).ToList();
                        break;
                    case "OrganizationName":
                        captureActs = captureActs.Where(act => act.Organization.Name.Contains(searchTerm)).ToList();
                        break;
                    case "ClaimDistrict":
                        captureActs = captureActs.Where(act => act.Claim.District.Contains(searchTerm)).ToList();
                        break;
                    default:
                        // Поле для поиска не выбрано, игнорируем поиск
                        break;
                }
            }
            if (!string.IsNullOrEmpty(sortField))
            {
                if (sortField == "CountDogs")
                {
                    captureActs = sortOrder == "Ascending" ? captureActs.OrderBy(c => c.CountDogs).ToList() : captureActs.OrderByDescending(c => c.CountDogs).ToList();
                }
                else if (sortField == "CountCats")
                {
                    captureActs = sortOrder == "Ascending" ? captureActs.OrderBy(c => c.CountCats).ToList() : captureActs.OrderByDescending(c => c.CountCats).ToList();
                }
                else if (sortField == "CountAnimals")
                {
                    captureActs = sortOrder == "Ascending" ? captureActs.OrderBy(c => c.CountAnimals).ToList() : captureActs.OrderByDescending(c => c.CountAnimals).ToList();
                }
                else if (sortField == "CaptureDate")
                {
                    captureActs = sortOrder == "Ascending" ? captureActs.OrderBy(c => c.CaptureDate).ToList() : captureActs.OrderByDescending(c => c.CaptureDate).ToList();
                }
                else if (sortField == "CapturePurpose")
                {
                    captureActs = sortOrder == "Ascending" ? captureActs.OrderBy(c => c.CapturePurpose).ToList() : captureActs.OrderByDescending(c => c.CapturePurpose).ToList();
                }
                else if (sortField == "Organization")
                {
                    captureActs = sortOrder == "Ascending" ? captureActs.OrderBy(c => c.OrganizationId).ToList() : captureActs.OrderByDescending(c => c.OrganizationId).ToList();
                }
                else if (sortField == "Contract")
                {
                    captureActs = sortOrder == "Ascending" ? captureActs.OrderBy(c => c.ContractId).ToList() : captureActs.OrderByDescending(c => c.ContractId).ToList();
                }
                else if (sortField == "Claim")
                {
                    captureActs = sortOrder == "Ascending" ? captureActs.OrderBy(c => c.ClaimId).ToList() : captureActs.OrderByDescending(c => c.ClaimId).ToList();
                }
            }

            ViewBag.SearchField = searchField;
            ViewBag.SearchTerm = searchTerm;

            _lastList = captureActs;
            return View(captureActs);
        }

        // GET: CaptureActs/Details/5
        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ,Оператор ВетСлужбы,Куратор ВетСлужбы,Подписант ВетСлужбы,Оператор по отлову,Куратор по отлову,Подписант по отлову,Оператор приюта,Куратор приюта,Подписант приюта")]
        public async Task<IActionResult> Details(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var captureActs = _context.CaptureActs
                .Include(c => c.Claim)
                .Include(c => c.Contract)
                .Include(c => c.Organization)
                .Where(m => m.Id == id).ToList();
            if (User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ") || User.IsInRole("Оператор по отлову")
               || User.IsInRole("Куратор по отлову") || User.IsInRole("Подписант  по отлову"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                captureActs = captureActs.Where(x => x.OrganizationId == user.OrganizationId).ToList();
            }
            var captureAct = captureActs.FirstOrDefault();
            if (captureAct == null)
            {
                return NotFound();
            }

            return View(captureAct);
        }

        // GET: CaptureActs/Create
        [Authorize(Roles = "Оператор по отлову")]
        public IActionResult Create()
        {
            var organizations = _context.Organizations.ToList();
            if (User.IsInRole("Оператор по отлову"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                organizations = organizations.Where(x => x.Id == user.OrganizationId).ToList();
            }
            ViewData["ClaimId"] = new SelectList(_context.Claims, "Id", "Id");
            ViewData["ContractId"] = new SelectList(_context.Contracts, "Id", "Id");
            ViewData["OrganizationId"] = new SelectList(organizations, "Id", "Id");
            return View();
        }

        // POST: CaptureActs/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [Authorize(Roles = "Оператор по отлову")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("CountDogs,CountCats,CountAnimals,CaptureDate,CapturePurpose,OrganizationId,ContractId,ClaimId,Id")] CaptureAct captureAct)
        {
            var organizations = _context.Organizations.ToList();
            if (User.IsInRole("Оператор по отлову"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                organizations = organizations.Where(x => x.Id == user.OrganizationId).ToList();
            }
            if (ModelState.IsValid)
            {
                captureAct.Id = Guid.NewGuid();
                _context.Add(captureAct);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            ViewData["ClaimId"] = new SelectList(_context.Claims, "Id", "Id", captureAct.ClaimId);
            ViewData["ContractId"] = new SelectList(_context.Contracts, "Id", "Id", captureAct.ContractId);
            ViewData["OrganizationId"] = new SelectList(organizations, "Id", "Id", captureAct.OrganizationId);
            return View(captureAct);
        }

        // GET: CaptureActs/Edit/5
        [Authorize(Roles = "Оператор по отлову")]
        public async Task<IActionResult> Edit(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var captureActs = _context.CaptureActs.Where(x => x.Id == id).ToList();
            var organizations = _context.Organizations.ToList();
            if (User.IsInRole("Оператор по отлову"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                organizations = organizations.Where(x => x.Id == user.OrganizationId).ToList();
                captureActs = captureActs.Where(x => x.OrganizationId == user.OrganizationId).ToList();
            }
            var captureAct = captureActs.FirstOrDefault();
            if (captureAct == null)
            {
                return NotFound();
            }
            ViewData["ClaimId"] = new SelectList(_context.Claims, "Id", "Id", captureAct.ClaimId);
            ViewData["ContractId"] = new SelectList(_context.Contracts, "Id", "Id", captureAct.ContractId);
            ViewData["OrganizationId"] = new SelectList(_context.Organizations, "Id", "Id", captureAct.OrganizationId);
            return View(captureAct);
        }

        // POST: CaptureActs/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [Authorize(Roles = "Оператор по отлову")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(Guid id, [Bind("CountDogs,CountCats,CountAnimals,CaptureDate,CapturePurpose,OrganizationId,ContractId,ClaimId,Id")] CaptureAct captureAct)
        {
            if (id != captureAct.Id)
            {
                return NotFound();
            }
            var organizations = _context.Organizations.ToList();
            if (User.IsInRole("Оператор по отлову"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                organizations = organizations.Where(x => x.Id == user.OrganizationId).ToList();
            }
            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(captureAct);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!CaptureActExists(captureAct.Id))
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
            ViewData["ClaimId"] = new SelectList(_context.Claims, "Id", "Id", captureAct.ClaimId);
            ViewData["ContractId"] = new SelectList(_context.Contracts, "Id", "Id", captureAct.ContractId);
            ViewData["OrganizationId"] = new SelectList(_context.Organizations, "Id", "Id", captureAct.OrganizationId);
            return View(captureAct);
        }

        // GET: CaptureActs/Delete/5
        [Authorize(Roles = "Оператор по отлову")]
        public async Task<IActionResult> Delete(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var captureActs = _context.CaptureActs
                .Include(c => c.Claim)
                .Include(c => c.Contract)
                .Include(c => c.Organization)
                .Where(m => m.Id == id).ToList();

            if (User.IsInRole("Оператор по отлову"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                captureActs = captureActs.Where(x => x.OrganizationId == user.OrganizationId).ToList();
            }
            var captureAct = captureActs.FirstOrDefault();
            if (captureAct == null)
            {
                return NotFound();
            }

            return View(captureAct);
        }

        // POST: CaptureActs/Delete/5
        [Authorize(Roles = "Оператор по отлову")]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(Guid id)
        {
            var captureActs = _context.CaptureActs
                .Include(c => c.Claim)
                .Include(c => c.Contract)
                .Include(c => c.Organization)
                .Where(m => m.Id == id).ToList();

            if (User.IsInRole("Оператор по отлову"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                captureActs = captureActs.Where(x => x.OrganizationId == user.OrganizationId).ToList();
            }
            var captureAct = captureActs.FirstOrDefault();
            _context.CaptureActs.Remove(captureAct);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }
        private IActionResult Download(List<CaptureAct> models)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                // Create the worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Contracts");

                // Add column headers
                worksheet.Cells[1, 1].Value = "CountDogs";
                worksheet.Cells[1, 2].Value = "CountCats";
                worksheet.Cells[1, 3].Value = "CountAnimals";
                worksheet.Cells[1, 4].Value = "CaptureDate";
                worksheet.Cells[1, 5].Value = "CapturePurpose";
                worksheet.Cells[1, 6].Value = "Organization";
                worksheet.Cells[1, 7].Value = "Contract";
                worksheet.Cells[1, 8].Value = "Claim";

                // Add data to the worksheet
                int row = 2;
                foreach (var item in models)
                {
                    worksheet.Cells[row, 1].Value = item.CountDogs;
                    worksheet.Cells[row, 2].Value = item.CountCats;
                    worksheet.Cells[row, 3].Value = item.CountAnimals;
                    worksheet.Cells[row, 4].Value = item.CaptureDate.ToString();
                    worksheet.Cells[row, 5].Value = item.CapturePurpose;
                    worksheet.Cells[row, 6].Value = item.Organization.Name;
                    worksheet.Cells[row, 7].Value = item.Contract.Id;
                    worksheet.Cells[row, 8].Value = item.Claim.District;
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

        private bool CaptureActExists(Guid id)
        {
            return _context.CaptureActs.Any(e => e.Id == id);
        }
    }
}
