
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Web.Mvc;
using System.Text.RegularExpressions;
using System.IO;
using NewsProj.DB.Interfaces;
using NewsProj.DB.Model;
using NewsProj.DB;

namespace NewsProj.Controllers
{
    public class HomeController : Controller
    {

        INewsletterRepository _Repo;
        ModelMapping _ModelMap;

        // Added empty constructor
        public HomeController() { }

        public HomeController(INewsletterRepository Repo, ModelMapping ModelMap)
        {
            _Repo = Repo;
            _ModelMap = ModelMap;
        }


        public ActionResult Index()
        {
            SelectListModel dropdowns = new SelectListModel();

            int latestNUM = 1;
            //NL19QC4
            var year = DateTime.Now.Year.ToString().Substring(2);
            string TicketId = "NL" + year + "QC" + (latestNUM + 1);
            var today = DateTime.Now;
            dropdowns.editionLink = "http://newsletters/" + TicketId + ".aspx";
            return View(dropdowns);
        }


        [HttpPost]
        public ActionResult AddTemplate(string Region, string Trends, int? trNumb, string LinkEdition, int lang, string subj, string editionNumb)
        {
            GlobalResult ResultObj = new GlobalResult();
        
            var UserStamp = "123456";
            var cc = new List<string>() { "" };
            string subject = subj;
            string Reg = "";
            switch (Region)
            {
                case "Quebec": Reg = "QC"; break;
                case "Ontario": Reg = "ON"; break;
                case "West": Reg = "MB"; break;
                case "Atlantic": Reg = "ATL"; break;
                default: Reg = "QC"; break;
            }
            try
            {
                // using NewsLetterContext
                using (var context = new NewsLetterContext())
                {
                    string lastTicket = "NL123456";

                    int latestNUM = int.Parse(Regex.Match(lastTicket, @"\d+").Value);
                    //NL19QC4
                    var year = DateTime.Now.Year.ToString().Substring(2);
                    string TicketId = "NL" + year + Reg + (latestNUM + 1);
                    var today = DateTime.Now;



                    if (!ResultObj.IsValid)
                    {
                        ResultObj.Message = "Unexpected error occurred";
                    }
                    NewsLetter nl = new NewsLetter();
                    string Emailbody = "";


                    var pein = "12345678";
                    nl.editionLink = LinkEdition;
                    nl.editionNumber = editionNumb;//TicketId;
                    nl.trends = Trends;
                    nl.createPein = UserStamp;
                    nl.lang = lang;


                    Emailbody = RenderPartialViewToString("EmailTemplate", nl);
                    Send(Emailbody, "", subject, cc);

                    ResultObj.IsValid = _Repo.SendRec(pein, "", DateTime.Now, Reg, UserStamp);

                    System.Threading.Thread.Sleep(200);

                }
            }
            catch (Exception ex)
            {
                ResultObj.IsValid = false;

                ResultObj.Message = "Unexpected error occurred";
            }
            return Json(ResultObj, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public ActionResult GetManagerEmployees(string Pein)
        {
            return Json(null, JsonRequestBehavior.AllowGet);
        }

        protected string RenderPartialViewToString(string viewName, object model)
        {
            if (string.IsNullOrEmpty(viewName))
            {
                viewName = ControllerContext.RouteData.GetRequiredString("action");
            }

            ViewData.Model = model;

            using (StringWriter sw = new StringWriter())
            {
                ViewEngineResult viewResult = ViewEngines.Engines.FindPartialView(ControllerContext, viewName);


                try
                {
                    ViewContext viewContext = new ViewContext(ControllerContext, viewResult.View, ViewData, TempData, sw);
                    viewResult.View.Render(viewContext, sw);

                }
                catch (Exception Ex)
                {

                }
                return sw.GetStringBuilder().ToString();
            }
        }

        public bool Send(string emailBody, string selectedEmail, string subject, List<string> cc = null)
        {
            try
            {
                //Create the email and send it
                using (var email = new MailMessage())
                {
                    email.From = new MailAddress("s");
                    email.Subject = subject;
                    email.Body = emailBody;
                    email.IsBodyHtml = true;

                    email.To.Add(new MailAddress(selectedEmail));

                    if (cc != null)
                    {
                        foreach (var c in cc)
                        {
                            email.CC.Add(new MailAddress(c));
                        }
                    }
                    var myMessage = new SendGrid.SendGridMessage();

                    // email.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(text, null, MediaTypeNames.Text.Plain));
                    // email.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(html, null, MediaTypeNames.Text.Html));

                    // Init SmtpClient and send
                    SmtpClient smtpClient = new SmtpClient("smtp.sendgrid.net", Convert.ToInt32(587));
                    System.Net.NetworkCredential credentials = new System.Net.NetworkCredential("", "mrusnaG@D2018");
                    
                    smtpClient.Credentials = credentials;

                    /*
                    var smtpHost = ConfigurationManager.AppSettings["SmtpHost"];
                    //You will have to modify the code below:
                    SmtpClient smtp = new SmtpClient
                    {
                        Host = smtpHost,
                        Port = 25,
                        EnableSsl = false,
                        DeliveryMethod = SmtpDeliveryMethod.Network,
                        UseDefaultCredentials = false
                    };
                    */
                    smtpClient.Send(email);
                }
                return true;
            }
            catch (Exception e)
            {

                return false;
            }
        }
    }
}