using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using WebApplicationLab6.Data;
using Claim = WebApplicationLab6.Objects.Claim;

namespace WebApplicationLab6.Controllers
{
    public class ClaimsController : Controller
    {
        private readonly ApplicationDbContext _context;

        public ClaimsController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: Claims
        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ,Оператор ВетСлужбы,Куратор ВетСлужбы,Подписант ВетСлужбы,Оператор по отлову,Куратор по отлову,Подписант по отлову")]
        public async Task<IActionResult> Index(string searchField, string searchTerm, string sortField, string sortOrder)
        {
            var applicationDbContext = _context.Claims.Include(c => c.City);
            var claims = await applicationDbContext.ToListAsync();
            if (User.IsInRole("Оператор ОМСУ") || User.IsInRole("Куратор ОМСУ") || User.IsInRole("Подписант ОМСУ") || User.IsInRole("Оператор по отлову")
                || User.IsInRole("Куратор по отлову") || User.IsInRole("Подписант  по отлову"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x=>x.Organization).FirstOrDefault(x => x.Login == userName);

                claims = claims.Where(x=>x.CityId == user.Organization.CityId).ToList();
            }

            if (!string.IsNullOrEmpty(searchField) && !string.IsNullOrEmpty(searchTerm))
            {
                switch (searchField)
                {
                    case "Claim":
                        claims = claims.Where(x => x.Id == Guid.Parse(searchTerm)).ToList();
                        break;
                    case "City":
                        claims = claims.Where(x => x.City.Name.Contains(searchTerm)).ToList();
                        break;
                    case "Description":
                        claims = claims.Where(x => x.Description.Contains(searchTerm)).ToList();
                        break;
                    default:
                        // Поле для поиска не выбрано, игнорируем поиск
                        break;
                }
            }
            if (!string.IsNullOrEmpty(sortField))
            {
                if (sortField == "FilingDate")
                {
                    claims = sortOrder == "Ascending" ? claims.OrderBy(c => c.FilingDate).ToList() : claims.OrderByDescending(c => c.FilingDate).ToList();
                }
                else if (sortField == "CategoryCustomer")
                {
                    claims = sortOrder == "Ascending" ? claims.OrderBy(c => c.CategoryCustomer).ToList() : claims.OrderByDescending(c => c.CategoryCustomer).ToList();
                }
                else if (sortField == "District")
                {
                    claims = sortOrder == "Ascending" ? claims.OrderBy(c => c.District).ToList() : claims.OrderByDescending(c => c.District).ToList();
                }
                else if (sortField == "Description")
                {
                    claims = sortOrder == "Ascending" ? claims.OrderBy(c => c.Description).ToList() : claims.OrderByDescending(c => c.Description).ToList();
                }
                else if (sortField == "IsDone")
                {
                    claims = sortOrder == "Ascending" ? claims.OrderBy(c => c.IsDone).ToList() : claims.OrderByDescending(c => c.IsDone).ToList();
                }
                else if (sortField == "City")
                {
                    claims = sortOrder == "Ascending" ? claims.OrderBy(c => c.CityId).ToList() : claims.OrderByDescending(c => c.CityId).ToList();
                }

            }
                ViewBag.SearchField = searchField;
            ViewBag.SearchTerm = searchTerm;

            return View(claims);
        }

        // GET: Claims/Details/5
        [Authorize(Roles = "Оператор ОМСУ,Куратор ОМСУ,Подписант ОМСУ,Оператор ВетСлужбы,Куратор ВетСлужбы,Подписант ВетСлужбы,Оператор по отлову,Куратор по отлову,Подписант по отлову")]
        public async Task<IActionResult> Details(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var claim = await _context.Claims
                .Include(c => c.City)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (claim == null)
            {
                return NotFound();
            }

            return View(claim);
        }

        // GET: Claims/Create
        [Authorize(Roles = "Оператор ОМСУ")]
        public IActionResult Create()
        {
            var cities = _context.Cities.ToList();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                cities = cities.Where(x => x.Id == user.Organization.CityId).ToList();
            }
            ViewData["CityId"] = new SelectList(cities, "Id", "Id");
            return View();
        }

        // POST: Claims/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [Authorize(Roles = "Оператор ОМСУ")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("FilingDate,CategoryCustomer,District,Description,IsDone,CityId,Id")] Claim claim)
        {
            var cities = _context.Cities.ToList();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                cities = cities.Where(x => x.Id == user.Organization.CityId).ToList();
            }
            if (ModelState.IsValid)
            {
                claim.Id = Guid.NewGuid();
                _context.Add(claim);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            ViewData["CityId"] = new SelectList(cities, "Id", "Id", claim.CityId);
            return View(claim);
        }

        // GET: Claims/Edit/5
        [Authorize(Roles = "Оператор ОМСУ")]
        public async Task<IActionResult> Edit(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var claims = _context.Claims.Include(x=>x.City).Where(x=>x.Id == id).ToList();
            var cities = _context.Cities.ToList();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                claims = claims.Where(x => x.City.Id == user.Organization.CityId).ToList();
                cities = cities.Where(x => x.Id == user.Organization.CityId).ToList();
            }
            var claim = claims.FirstOrDefault();
            if (claim == null)
            {
                return NotFound();
            }
            ViewData["CityId"] = new SelectList(cities, "Id", "Id", claim.CityId);
            return View(claim);
        }

        [Authorize(Roles = "Оператор ОМСУ")]
        // POST: Claims/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(Guid id, [Bind("FilingDate,CategoryCustomer,District,Description,IsDone,CityId,Id")] Claim claim)
        {
            if (id != claim.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(claim);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!ClaimExists(claim.Id))
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
            ViewData["CityId"] = new SelectList(_context.Cities, "Id", "Id", claim.CityId);
            return View(claim);
        }

        // GET: Claims/Delete/5
        [Authorize(Roles = "Оператор ОМСУ,Оператор по отлову")]
        public async Task<IActionResult> Delete(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var claims = _context.Claims
                .Include(c => c.City)
                .Where(m => m.Id == id).ToList();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                claims = claims.Where(x => x.City.Id == user.Organization.CityId).ToList();
            }
            var claim = claims.FirstOrDefault();
            if (claim == null)
            {
                return NotFound();
            }

            return View(claim);
        }

        [Authorize(Roles = "Оператор ОМСУ,Оператор по отлову")]
        // POST: Claims/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(Guid id)
        {
            var claims = _context.Claims
                 .Include(c => c.City)
                 .Where(m => m.Id == id).ToList();
            if (User.IsInRole("Оператор ОМСУ"))
            {
                string userName = User.Identity.Name;

                var user = _context.Users.Include(x => x.Organization).FirstOrDefault(x => x.Login == userName);

                claims = claims.Where(x => x.City.Id == user.Organization.CityId).ToList();
            }
            var claim = claims.FirstOrDefault();
            if (claim == null)
            {
                return NotFound();
            }
            _context.Claims.Remove(claim);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool ClaimExists(Guid id)
        {
            return _context.Claims.Any(e => e.Id == id);
        }
    }
}
