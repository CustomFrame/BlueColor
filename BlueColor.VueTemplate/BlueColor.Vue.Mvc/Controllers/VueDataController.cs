using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace BlueColor.Vue.Mvc.Controllers
{
    public class VueDataController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}