﻿using System.Web.Mvc;


namespace EmailSender.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            return View();
        }

        public ActionResult SendEmail()
        {
            ViewBag.Message = "Your contact page.";

            EmailSenderNetMail.SendCalendarEvent(EmailSenderNetMail.CreateCalendar());

            return View();
        }
    }
}