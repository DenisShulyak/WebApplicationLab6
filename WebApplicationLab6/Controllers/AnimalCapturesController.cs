using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using WebApplicationLab6.Data;
using WebApplicationLab6.Objects;

namespace WebApplicationLab6.Controllers
{
    public class AnimalCapturesController : Controller
    {
        private readonly ApplicationDbContext _context;

        public AnimalCapturesController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: AnimalCaptures
        public async Task<IActionResult> Index()
        {
            var applicationDbContext = _context.AnimalCaptures.Include(a => a.CaptureAct).Include(a => a.City);
            return View(await applicationDbContext.ToListAsync());
        }

        // GET: AnimalCaptures/Details/5
        public async Task<IActionResult> Details(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var animalCapture = await _context.AnimalCaptures
                .Include(a => a.CaptureAct)
                .Include(a => a.City)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (animalCapture == null)
            {
                return NotFound();
            }

            return View(animalCapture);
        }

        // GET: AnimalCaptures/Create
        public IActionResult Create()
        {
            ViewData["CaptureActId"] = new SelectList(_context.CaptureActs, "Id", "Id");
            ViewData["CityId"] = new SelectList(_context.Cities, "Id", "Id");
            return View();
        }

        // POST: AnimalCaptures/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Property,Animal,Ears,Tail,Features,IdentityMark,EChipNum,CaptureActId,CityId,Id")] AnimalCapture animalCapture)
        {
            if (ModelState.IsValid)
            {
                animalCapture.Id = Guid.NewGuid();
                _context.Add(animalCapture);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            ViewData["CaptureActId"] = new SelectList(_context.CaptureActs, "Id", "Id", animalCapture.CaptureActId);
            ViewData["CityId"] = new SelectList(_context.Cities, "Id", "Id", animalCapture.CityId);
            return View(animalCapture);
        }

        // GET: AnimalCaptures/Edit/5
        public async Task<IActionResult> Edit(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var animalCapture = await _context.AnimalCaptures.FindAsync(id);
            if (animalCapture == null)
            {
                return NotFound();
            }
            ViewData["CaptureActId"] = new SelectList(_context.CaptureActs, "Id", "Id", animalCapture.CaptureActId);
            ViewData["CityId"] = new SelectList(_context.Cities, "Id", "Id", animalCapture.CityId);
            return View(animalCapture);
        }

        // POST: AnimalCaptures/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(Guid id, [Bind("Property,Animal,Ears,Tail,Features,IdentityMark,EChipNum,CaptureActId,CityId,Id")] AnimalCapture animalCapture)
        {
            if (id != animalCapture.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(animalCapture);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!AnimalCaptureExists(animalCapture.Id))
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
            ViewData["CaptureActId"] = new SelectList(_context.CaptureActs, "Id", "Id", animalCapture.CaptureActId);
            ViewData["CityId"] = new SelectList(_context.Cities, "Id", "Id", animalCapture.CityId);
            return View(animalCapture);
        }

        // GET: AnimalCaptures/Delete/5
        public async Task<IActionResult> Delete(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var animalCapture = await _context.AnimalCaptures
                .Include(a => a.CaptureAct)
                .Include(a => a.City)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (animalCapture == null)
            {
                return NotFound();
            }

            return View(animalCapture);
        }

        // POST: AnimalCaptures/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(Guid id)
        {
            var animalCapture = await _context.AnimalCaptures.FindAsync(id);
            _context.AnimalCaptures.Remove(animalCapture);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool AnimalCaptureExists(Guid id)
        {
            return _context.AnimalCaptures.Any(e => e.Id == id);
        }
    }
}
