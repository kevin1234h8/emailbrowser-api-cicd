using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Net;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http.Results;
using System.Web.Http;
using System.Web.Mvc;
using AGOServer.Components;
using System.Web.Http;
using System.Web.Http.Description;
using System.Web.Http.Results;

namespace AGOServer.Controllers.Version1
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            //ViewBag.Title = "Home Page";
            //return View();
            return new EmptyResult();
        }

       
    }
}
