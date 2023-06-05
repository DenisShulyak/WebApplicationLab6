using System;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using WebApplicationLab6.Data;
using WebApplicationLab6.Objects;

namespace WebApplicationLab6.Controllers
{
    public class CaptureActsController : Controller
    {
        private readonly ApplicationDbContext _context;

        public CaptureActsController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: CaptureActs
        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ,Оператор ВетСлужбы,Куратор ВетСлужбы,Подписант ВетСлужбы,Оператор по отлову,Куратор по отлову,Подписант по отлову,Оператор приюта,Куратор приюта,Подписант приюта")]
        public async Task<IActionResult> Index(string searchField, string searchTerm, string sortField, string sortOrder)
        {
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
                    case "ContractId":
                        captureActs = captureActs.Where(act => act.ContractId == Guid.Parse(searchTerm)).ToList();
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

            var captureActs = _context.CaptureActs.Where(x=>x.Id == id).ToList();
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

        private bool CaptureActExists(Guid id)
        {
            return _context.CaptureActs.Any(e => e.Id == id);
        }
    }
}
