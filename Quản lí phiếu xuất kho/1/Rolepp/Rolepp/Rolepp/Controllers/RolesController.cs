using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;

namespace Rolepp.Controllers
{
    public class RolesController : Controller
    {
        private readonly RoleManager<IdentityRole> _manager;
        public RolesController(RoleManager<IdentityRole> roleManager) { 
            _manager = roleManager; 
        }
        public IActionResult Index()
        {
            var roles = _manager.Roles;
            return View(roles);
        }


        [HttpGet]
        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Create(IdentityRole role)
        {
            //check if the role exist
            if(!_manager.RoleExistsAsync(role.Name).GetAwaiter().GetResult())
            {
                _manager.CreateAsync(new IdentityRole(role.Name)).GetAwaiter().GetResult();
            }

            return RedirectToAction("Index");
        }
    }
}
