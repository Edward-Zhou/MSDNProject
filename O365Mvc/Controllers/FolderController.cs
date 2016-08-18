using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using O365Mvc.Helpers;
using O365Mvc.Models;
using O365Mvc.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace O365Mvc.Controllers
{
    [Authorize]
    public class FolderController : Controller
    {
        private MailOperations mailOperations = new MailOperations();
        // GET: Folder
        public ActionResult getFolder() {
            List<string> Folders = new List<string>();
            Folders.Add("tao");
            Folders.Add("sven");
            return View("FolderCollections", Folders);
        }
        //query Folders
        public async Task<ActionResult> FolderCollections()
        {
            List<string> Folders = new List<string>();
            try
            {
                Folders = await mailOperations.GetEmailFolders();
            }
            catch (Exception e)
            { }
            return View("FolderCollections", Folders);
        }
        //create Folder
        public async Task<ActionResult> FolderCreate()
        {
            List<string> Folders = new List<string>();
            List<string> newFolders = new List<string>();
            newFolders.Add("New1");
            newFolders.Add("New2");
            try
            {
                Folders = await mailOperations.CreateEmailFolder(newFolders);
            }
            catch (Exception e)
            { }
            return View(Folders);
        }
        //delete Folder
        public async Task<ActionResult> FolderDelete()
        {
            List<string> Folders = new List<string>();
            string newFolders = "New1";
            try
            {
                Folders = await mailOperations.DeleteEmailFolder(newFolders);
            }
            catch (Exception e)
            { }
            return View(Folders);
        }
    }
}