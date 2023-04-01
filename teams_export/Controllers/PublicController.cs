// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System.IO;
using System.Web;
using System.Web.Mvc;

namespace teams_export.Controllers
{
    public class PublicController : BaseController
    {
        [Route("/public/{chatId}/{fileName}/{downloadName}")]
        public ActionResult Index(string chatId, string fileName, string downloadName)
        {
            string mimeType = MimeMapping.GetMimeMapping(fileName);
            var filePath = Server.MapPath($"~/Content/public/{chatId}/{fileName}");
            return File(new FileStream(filePath, FileMode.Open, FileAccess.Read), mimeType, downloadName);
        }
    }
}