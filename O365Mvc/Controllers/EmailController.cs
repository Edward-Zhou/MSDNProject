using O365Mvc.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace O365Mvc.Controllers
{
    [Authorize]
    public class EmailController : Controller
    {
        private MailOperations mailOperations = new MailOperations();
        // Create Email
        public async Task<ActionResult> CreateEmail()
        {
            await mailOperations.CreateEmail();
            return View();
        }
        //get attachment
        public async Task<ActionResult> GetAttachment()
        {
            await mailOperations.GetAttachment();
            return View();
        }
    }
}