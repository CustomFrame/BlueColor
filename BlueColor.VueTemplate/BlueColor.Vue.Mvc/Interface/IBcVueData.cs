using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BlueColor.Vue.Mvc.Interface
{
    /// <summary>
    /// IBcVueData
    /// </summary>
    interface IBcVueData
    {
        IActionResult Index();

        IActionResult HandleSearch();

        IActionResult HandleSave();

        IActionResult HandleDelete();
    }
}
