﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace teams_export.Controllers
{
    public class HomeController : BaseController
    {
        public ActionResult Index()
        {
            return RedirectToAction("Index", "Chat");
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult Error(string message, string debug)
        {
            Flash(message, debug);
            return RedirectToAction("Index");
        }
    }
}