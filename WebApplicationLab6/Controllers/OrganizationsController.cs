using System;
using System.Collections.Generic;
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
    public class OrganizationsController : Controller
    {
        private readonly ApplicationDbContext _context;

        private static List<Organization> _lastList = new List<Organization>();

        public OrganizationsController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: Organizations
        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ,Оператор ВетСлужбы,Куратор ВетСлужбы,Подписант ВетСлужбы,Куратор по отлову,Подписант по отлову,Куратор приюта,Подписант приюта")]
        public async Task<IActionResult> Index(string searchField, string searchTerm, string sortField, string sortOrder, string downloadData)
        {
            if (!string.IsNullOrEmpty(downloadData) && downloadData.ToLower() == "true")
            {
                return Download(_lastList);
            }

            var applicationDbContext = _context.Organizations.Include(o => o.City).Include(o => o.OrganizationType);
            var organizations = await applicationDbContext.ToListAsync();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Приют" ||
                        x.OrganizationType.Name == "Организация по отлову" ||
                        x.OrganizationType.Name == "Организация по отлову и приюту" ||
                        x.OrganizationType.Name == "Организация по транспортировке" ||
                        x.OrganizationType.Name == "ГосВетКлиника" ||
                        x.OrganizationType.Name == "ЧастВетКлиника" ||
                        x.OrganizationType.Name == "Благотворительный фонд" ||
                        x.OrganizationType.Name == "Организации по продаже товаров и предоставлению услуг для животных")
                    .ToList();
            }
            else if (User.IsInRole("Оператор ВетСлужбы"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Исполнительный орган государственной власти" ||
                        x.OrganizationType.Name == "ОМСУ" ||
                        x.OrganizationType.Name == "ГосВетКлиника")
                    .ToList();
            }
            else if (User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                organizations = organizations.Where(x => x.Id == user.OrganizationId).ToList();
            }
            if (!string.IsNullOrEmpty(searchField) && !string.IsNullOrEmpty(searchTerm))
            {
                switch (searchField)
                {
                    case "OrganizationName":
                        organizations = organizations.Where(x => x.Name.Contains(searchTerm)).ToList();
                        break;
                    case "OrhanizationType":
                        organizations = organizations.Where(x => x.OrganizationType.Name.Contains(searchTerm)).ToList();
                        break;
                    case "City":
                        organizations = organizations.Where(x => x.City.Name.Contains(searchTerm)).ToList();
                        break;
                    default:
                        // Поле для поиска не выбрано, игнорируем поиск
                        break;
                }
            }

            if (!string.IsNullOrEmpty(sortField))
            {
                if (sortField == "Name")
                {
                    organizations = sortOrder == "Ascending" ? organizations.OrderBy(c => c.Name).ToList() : organizations.OrderByDescending(c => c.Name).ToList();
                }
                else if (sortField == "Inn")
                {
                    organizations = sortOrder == "Ascending" ? organizations.OrderBy(c => c.Inn).ToList() : organizations.OrderByDescending(c => c.Inn).ToList();
                }
                else if (sortField == "Kpp")
                {
                    organizations = sortOrder == "Ascending" ? organizations.OrderBy(c => c.Kpp).ToList() : organizations.OrderByDescending(c => c.Kpp).ToList();
                }
                else if (sortField == "Address")
                {
                    organizations = sortOrder == "Ascending" ? organizations.OrderBy(c => c.Address).ToList() : organizations.OrderByDescending(c => c.Address).ToList();
                }
                else if (sortField == "City")
                {
                    organizations = sortOrder == "Ascending" ? organizations.OrderBy(c => c.CityId).ToList() : organizations.OrderByDescending(c => c.CityId).ToList();
                }
                else if (sortField == "OrganizationType")
                {
                    organizations = sortOrder == "Ascending" ? organizations.OrderBy(c => c.OrganizationTypeId).ToList() : organizations.OrderByDescending(c => c.OrganizationTypeId).ToList();
                }
            }

            ViewBag.SearchField = searchField;
            ViewBag.SearchTerm = searchTerm;
            _lastList = organizations;
            return View(organizations);
        }

        private IActionResult Download(List<Organization> models)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                // Create the worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Contracts");

                // Add column headers
                worksheet.Cells[1, 1].Value = "Name";
                worksheet.Cells[1, 2].Value = "Inn";
                worksheet.Cells[1, 3].Value = "Kpp";
                worksheet.Cells[1, 4].Value = "Address";
                worksheet.Cells[1, 5].Value = "City";
                worksheet.Cells[1, 6].Value = "OrganizationType";

                // Add data to the worksheet
                int row = 2;
                foreach (var item in models)
                {
                    worksheet.Cells[row, 1].Value = item.Name;
                    worksheet.Cells[row, 2].Value = item.Inn;
                    worksheet.Cells[row, 3].Value = item.Kpp;
                    worksheet.Cells[row, 4].Value = item.Address;
                    worksheet.Cells[row, 5].Value = item.City.Name;
                    worksheet.Cells[row, 6].Value = item.OrganizationType.Name;
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

        // GET: Organizations/Details/5
        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ,Оператор ВетСлужбы,Куратор ВетСлужбы,Подписант ВетСлужбы,Куратор по отлову,Подписант по отлову,Куратор приюта,Подписант приюта")]
        public async Task<IActionResult> Details(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var organizations = _context.Organizations
                .Include(o => o.City)
                .Include(o => o.OrganizationType)
                .Where(m => m.Id == id).ToList();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Приют" ||
                        x.OrganizationType.Name == "Организация по отлову" ||
                        x.OrganizationType.Name == "Организация по отлову и приюту" ||
                        x.OrganizationType.Name == "Организация по транспортировке" ||
                        x.OrganizationType.Name == "ГосВетКлиника" ||
                        x.OrganizationType.Name == "ЧастВетКлиника" ||
                        x.OrganizationType.Name == "Благотворительный фонд" ||
                        x.OrganizationType.Name == "Организации по продаже товаров и предоставлению услуг для животных")
                    .ToList();
            }
            else if (User.IsInRole("Оператор ВетСлужбы"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Исполнительный орган государственной власти" ||
                        x.OrganizationType.Name == "ОМСУ" ||
                        x.OrganizationType.Name == "ГосВетКлиника")
                    .ToList();
            }
            else if (User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                organizations = organizations.Where(x => x.Id == user.OrganizationId).ToList();
            }
            var organization = organizations.FirstOrDefault();
            if (organization == null)
            {
                return NotFound();
            }

            return View(organization);
        }

        // GET: Organizations/Create
        [Authorize(Roles = "Оператор ОМСУ,Оператор ВетСлужбы")]
        public IActionResult Create()
        {
            ViewData["CityId"] = new SelectList(_context.Cities, "Id", "Id");
            List<OrganizationType> organizationTypes = null;
            if (User.IsInRole("Оператор ОМСУ"))
            {
                organizationTypes = _context.OrganizationTypes.Where(x => x.Name == "Приют" ||
                        x.Name == "Организация по отлову" ||
                        x.Name == "Организация по отлову и приюту" ||
                        x.Name == "Организация по транспортировке" ||
                        x.Name == "ГосВетКлиника" ||
                        x.Name == "ЧастВетКлиника" ||
                        x.Name == "Благотворительный фонд" ||
                        x.Name == "Организации по продаже товаров и предоставлению услуг для животных")
                    .ToList();
            }
            else if (User.IsInRole("Оператор ВетСлужбы"))
            {
                organizationTypes = _context.OrganizationTypes.Where(x => x.Name == "Исполнительный орган государственной власти" ||
                        x.Name == "ОМСУ" ||
                        x.Name == "ГосВетКлиника")
                    .ToList();
            }
            else
            {
                organizationTypes = _context.OrganizationTypes.ToList();
            }
            ViewData["OrganizationTypeId"] = new SelectList(organizationTypes, "Id", "Id");
            return View();
        }

        // POST: Organizations/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [Authorize(Roles = "Оператор ОМСУ,Оператор ВетСлужбы")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Name,Inn,Kpp,Address,CityId,OrganizationTypeId,Id")] Organization organization)
        {
            if (ModelState.IsValid)
            {
                organization.Id = Guid.NewGuid();
                _context.Add(organization);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            ViewData["CityId"] = new SelectList(_context.Cities, "Id", "Id", organization.CityId);
            ViewData["OrganizationTypeId"] = new SelectList(_context.OrganizationTypes, "Id", "Id", organization.OrganizationTypeId);
            return View(organization);
        }

        // GET: Organizations/Edit/5
        [Authorize(Roles = "Оператор ОМСУ,Оператор ВетСлужбы")]
        public async Task<IActionResult> Edit(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var organizations = _context.Organizations.Where(x=>x.Id == id).ToList();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Приют" ||
                        x.OrganizationType.Name == "Организация по отлову" ||
                        x.OrganizationType.Name == "Организация по отлову и приюту" ||
                        x.OrganizationType.Name == "Организация по транспортировке" ||
                        x.OrganizationType.Name == "ГосВетКлиника" ||
                        x.OrganizationType.Name == "ЧастВетКлиника" ||
                        x.OrganizationType.Name == "Благотворительный фонд" ||
                        x.OrganizationType.Name == "Организации по продаже товаров и предоставлению услуг для животных")
                    .ToList();
            }
            else if (User.IsInRole("Оператор ВетСлужбы"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Исполнительный орган государственной власти" ||
                        x.OrganizationType.Name == "ОМСУ" ||
                        x.OrganizationType.Name == "ГосВетКлиника")
                    .ToList();
            }
            var organization = organizations.FirstOrDefault();
            if (organization == null)
            {
                return NotFound();
            }
            ViewData["CityId"] = new SelectList(_context.Cities, "Id", "Id", organization.CityId);
            ViewData["OrganizationTypeId"] = new SelectList(_context.OrganizationTypes, "Id", "Id", organization.OrganizationTypeId);
            return View(organization);
        }

        // POST: Organizations/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [Authorize(Roles = "Оператор ОМСУ,Оператор ВетСлужбы")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(Guid id, [Bind("Name,Inn,Kpp,Address,CityId,OrganizationTypeId,Id")] Organization organization)
        {
            if (id != organization.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(organization);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!OrganizationExists(organization.Id))
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
            ViewData["CityId"] = new SelectList(_context.Cities, "Id", "Id", organization.CityId);
            ViewData["OrganizationTypeId"] = new SelectList(_context.OrganizationTypes, "Id", "Id", organization.OrganizationTypeId);
            return View(organization);
        }

        // GET: Organizations/Delete/5
        [Authorize(Roles = "Оператор ОМСУ,Оператор ВетСлужбы")]
        public async Task<IActionResult> Delete(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var organizations =  _context.Organizations
                .Include(o => o.City)
                .Include(o => o.OrganizationType)
                .Where(m => m.Id == id).ToList();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Приют" ||
                        x.OrganizationType.Name == "Организация по отлову" ||
                        x.OrganizationType.Name == "Организация по отлову и приюту" ||
                        x.OrganizationType.Name == "Организация по транспортировке" ||
                        x.OrganizationType.Name == "ГосВетКлиника" ||
                        x.OrganizationType.Name == "ЧастВетКлиника" ||
                        x.OrganizationType.Name == "Благотворительный фонд" ||
                        x.OrganizationType.Name == "Организации по продаже товаров и предоставлению услуг для животных")
                    .ToList();
            }
            else if (User.IsInRole("Оператор ВетСлужбы"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Исполнительный орган государственной власти" ||
                        x.OrganizationType.Name == "ОМСУ" ||
                        x.OrganizationType.Name == "ГосВетКлиника")
                    .ToList();
            }
            var organization = organizations.FirstOrDefault();
            if (organization == null)
            {
                return NotFound();
            }

            return View(organization);
        }

        // POST: Organizations/Delete/5
        [Authorize(Roles = "Оператор ОМСУ,Оператор ВетСлужбы")]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(Guid id)
        {
            var organizations = _context.Organizations.Where(x=>x.Id == id).ToList();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Приют" ||
                        x.OrganizationType.Name == "Организация по отлову" ||
                        x.OrganizationType.Name == "Организация по отлову и приюту" ||
                        x.OrganizationType.Name == "Организация по транспортировке" ||
                        x.OrganizationType.Name == "ГосВетКлиника" ||
                        x.OrganizationType.Name == "ЧастВетКлиника" ||
                        x.OrganizationType.Name == "Благотворительный фонд" ||
                        x.OrganizationType.Name == "Организации по продаже товаров и предоставлению услуг для животных")
                    .ToList();
            }
            else if (User.IsInRole("Оператор ВетСлужбы"))
            {
                organizations = organizations.Where(x => x.OrganizationType.Name == "Исполнительный орган государственной власти" ||
                        x.OrganizationType.Name == "ОМСУ" ||
                        x.OrganizationType.Name == "ГосВетКлиника")
                    .ToList();
            }
            var organization = organizations.FirstOrDefault();
            if (organization == null)
            {
                return NotFound();
            }

            _context.Organizations.Remove(organization);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool OrganizationExists(Guid id)
        {
            return _context.Organizations.Any(e => e.Id == id);
        }
    }
}
