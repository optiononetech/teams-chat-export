﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using teams_export.Models;
using System.Collections.Generic;
using System.Web.Mvc;
using teams_export.TokenStorage;
using System.Security.Claims;
using System.Web;
using Microsoft.Owin.Security.Cookies;

namespace teams_export.Controllers
{
    public abstract class BaseController : Controller
    {
        protected void Flash(string message, string debug = null)
        {
            var alerts = TempData.ContainsKey(Alert.AlertKey) ?
                (List<Alert>)TempData[Alert.AlertKey] :
                new List<Alert>();

            alerts.Add(new Alert
            {
                Message = message,
                Debug = debug
            });

            TempData[Alert.AlertKey] = alerts;
        }

        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (Request.IsAuthenticated)
            {
                // Get the user's token cache
                var tokenStore = new SessionTokenStore(null,
                    System.Web.HttpContext.Current, ClaimsPrincipal.Current);

                if (tokenStore.HasData())
                {
                    // Add the user to the view bag
                    ViewBag.User = tokenStore.GetUserDetails();
                }
                else
                {
                    // The session has lost data. This happens often
                    // when debugging. Log out so the user can log back in
                    Request.GetOwinContext().Authentication.SignOut(CookieAuthenticationDefaults.AuthenticationType);
                    filterContext.Result = RedirectToAction("Index", "Home");
                }
            }

            base.OnActionExecuting(filterContext);
        }
    }
}