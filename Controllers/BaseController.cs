using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExportDocApi.Controllers
{
    public class BaseController : Controller
    {
        
        public ActionResult Download(string fileGuid, string fileName)
        {
            var dataCache = System.Web.HttpContext.Current.Cache.Get(fileGuid);
            if (dataCache != null)
            {
                byte[] data = dataCache as byte[];
                System.Web.HttpContext.Current.Cache.Remove(fileGuid);

                return File(data, "application/octet-stream", fileName);
            }
            else
            {
                return new EmptyResult();
            }
        }
    }
}