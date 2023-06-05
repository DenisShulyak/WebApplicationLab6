using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Security.Claims;
using System.Threading.Tasks;
using WebApplicationLab6.ViewModels;
using Microsoft.EntityFrameworkCore;
using WebApplicationLab6.Objects;
using Microsoft.Extensions.Configuration;
using WebApplicationLab6.Data;
using WebApplicationLab6.Services;
using System.Linq;
using System;
using WebApplicationLab6.Configurations;
using Claim = System.Security.Claims.Claim;

namespace WebApplicationLab6.Controllers
{
    public class AccountController : Controller
    {

        private readonly JwtService _jwtService;

        private readonly ApplicationDbContext _context;

        public AccountController(JwtService jwtService, ApplicationDbContext context)
        {
            _jwtService = jwtService;
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> Login()
        {
            if (User.Identity.IsAuthenticated)
            {
                await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
                return RedirectToAction("Index", "Home");
            }
            var model = new UserViewModel(); // Создание экземпляра модели данных
            return View(model);
        }

        [AllowAnonymous]
        [HttpPost]
        public async Task<IActionResult> Login(UserViewModel model)
        {
            var user = Authenticate(model.Login, model.Password);

            if (user == null)
            {
                ModelState.AddModelError(string.Empty, "Неверный логин или пароль");
                return View(model);
            }

            var token = _jwtService.GenerateToken(user.Login, user.Role.Name);

            if (string.IsNullOrEmpty(token))
            {
                ModelState.AddModelError(string.Empty, "Ошибка генерации токена");
                return View(model);
            }

            var claims = new List<Claim>
    {
        new Claim(ClaimTypes.Name, user.Login),
        new Claim(ClaimTypes.Role, user.Role.Name)
    };

            var claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);

            var authProperties = new AuthenticationProperties
            {
                AllowRefresh = true,
                ExpiresUtc = DateTimeOffset.UtcNow.AddMinutes(AuthOptions.LIFETIME),
                IsPersistent = true
            };

            await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity), authProperties);

            return RedirectToAction("Index", "Home");
        }

        private User Authenticate(string login, string password)
        {
            var user = _context.Users.Include(x => x.Role).FirstOrDefault(x => x.Login == login && x.Password == password);

            return user;
        }

        [Authorize]
        [HttpGet]
        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            return RedirectToAction("Index", "Home"); ;
        }

        [AllowAnonymous]
        public IActionResult AccessDenied()
        {
            TempData["Message"] = "У вас нет прав доступа.";
            return RedirectToAction("Index", "Home");
        }
    }
}
