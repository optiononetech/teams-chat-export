// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using teams_export.Helpers;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Linq;
using static System.Windows.Forms.AxHost;

namespace teams_export.Controllers
{
    [Authorize]
    [SessionState(System.Web.SessionState.SessionStateBehavior.ReadOnly)]
    public class ChatController : BaseController
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Status(string actionId)
        {
            return Json(HttpContext.GetCurrentAction(actionId), JsonRequestBehavior.AllowGet);
        }

        public ActionResult Delta(string actionId)
        {
            return Json(HttpContext.GetDelta(actionId), JsonRequestBehavior.AllowGet);
        }

        public async Task<ActionResult> ChatList(string actionId)
        {
            var chats = await GraphHelper.GetChatsAsync(actionId, HttpContext);
            return Json(chats, JsonRequestBehavior.AllowGet);
        }

        public async Task<ActionResult> MemberList(string id)
        {
            var members = await GraphHelper.GetChatMembersAsync(id);
            return Json(members, JsonRequestBehavior.AllowGet);
        }

        public async Task<ActionResult> Open(string id)
        {
            var chat = await GraphHelper.GetChatAsync(id);
            return View(chat);
        }

        public async Task<ActionResult> MessageList(string actionId, string chatId, DateTime since, DateTime until)
        {
            var hash = Hash(chatId);
            var dir = Server.MapPath("~/Content/public/" + hash);
            if (!System.IO.Directory.Exists(dir)) System.IO.Directory.CreateDirectory(dir);
            var messages = await GraphHelper.GetChatMessagesAsync(chatId, since, until, dir, false, actionId, HttpContext);
            return Json(messages, JsonRequestBehavior.AllowGet);
        }

        public async Task<ActionResult> Export(string actionId, string chatId, DateTime since, DateTime until)
        {
            var chat = await GraphHelper.GetChatAsync(chatId);

            var safeChatId = Hash(chatId + since.ToString("o") + until.ToString("o"));
            var tempDir = Path.Combine(AppContext.BaseDirectory, "export", "raw", safeChatId);
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            var assetsDir = Path.Combine(tempDir, "assets");
            if (!Directory.Exists(assetsDir)) Directory.CreateDirectory(assetsDir);
            var outputZip = Path.Combine(AppContext.BaseDirectory, "export", $"{safeChatId}.zip");
            if (System.IO.File.Exists(outputZip)) System.IO.File.Delete(outputZip);
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            var htmlDoc = Path.Combine(tempDir, "index.html");
            using (var stream = new FileStream(htmlDoc, FileMode.Create))
            using (var writer = new StreamWriter(stream, Encoding.UTF8, 8192, true))
            {
                writer.WriteLine("<!doctype HTML>");
                writer.WriteLine("<html>");
                {
                    writer.WriteLine("<head>");
                    {
                        writer.WriteLine($"<title>{(string.IsNullOrWhiteSpace(chat.Topic) ? "Private Chat" : chat.Topic)}</title>");
                        writer.WriteLine($"<style>");
                        {
                            writer.WriteLine(System.IO.File.ReadAllText(Server.MapPath("~/Content/bootstrap.min.css")));
                        }
                        writer.WriteLine($"</style>");
                    }
                    writer.WriteLine("</head>");
                    writer.WriteLine("<body>");
                    {
                        writer.WriteLine($"<h1>{(string.IsNullOrWhiteSpace(chat.Topic) ? "Private Chat" : chat.Topic)}</h1>");
                        foreach (var member in chat.Members)
                            writer.WriteLine($"<span class='btn btn-secondary m-2'>{member?.DisplayName}</span>");

                        writer.WriteLine("<div class='messages'>");
                        {
                            var messages = await GraphHelper.GetChatMessagesAsync(chatId, since, until, assetsDir, true, actionId, HttpContext);

                            writer.WriteLine("<div class='container'>");
                            {
                                writer.WriteLine("<div class='row'>");
                                {
                                    writer.WriteLine("<div class='col-12'>");
                                    {
                                        writer.WriteLine("<div class='card m-1' data-id='message-container'>");
                                        {
                                            writer.WriteLine("<div class='card-body'>");
                                            {
                                                writer.WriteLine("<div class='container' style='overflow: auto; max-height: calc(100vh - 300px)'>");
                                                {
                                                    foreach (var message in messages.AsEnumerable().Reverse())
                                                    {
                                                        var fromPrincipal = message.From.User.DisplayName == chat.Members.FirstOrDefault()?.DisplayName;
                                                        var style = ""; if (fromPrincipal) style = "background-color: rgb(232, 235, 250)";
                                                        writer.WriteLine("<div class='row m-2'>");
                                                        {
                                                            if (fromPrincipal) writer.WriteLine("<div class='col-7'></div>");
                                                            writer.WriteLine($"<div class='col-5 p-2 shadow'  style='{style}'>");
                                                            {
                                                                writer.WriteLine($"<div>");
                                                                {
                                                                    writer.WriteLine($"<b>"); writer.WriteLine(message.From.User.DisplayName); writer.WriteLine("</b>");
                                                                    writer.WriteLine($"<span style='float: right; font-weight: light'>"); writer.WriteLine(message.CreatedDateTime.Value.DateTime.ToLongDateString() + " " + message.CreatedDateTime.Value.DateTime.ToLongTimeString()); writer.WriteLine("</span>");
                                                                }
                                                                writer.WriteLine("</div>");
                                                                writer.WriteLine($"<hr style='margin-top: 0px; margin-bottom: 2px;'/>");
                                                                writer.WriteLine($"<div>");
                                                                {
                                                                    writer.WriteLine(message.Body.Content);
                                                                }
                                                                writer.WriteLine("</div>");
                                                            }
                                                            writer.WriteLine("</div>");
                                                        }
                                                        writer.WriteLine("</div>");
                                                    }
                                                }
                                                writer.WriteLine("</div>");
                                            }
                                            writer.WriteLine("</div>");
                                        }
                                        writer.WriteLine("</div>");
                                    }
                                    writer.WriteLine("</div>");
                                }
                                writer.WriteLine("</div>");
                            }
                            writer.WriteLine("</div>");
                        }
                        writer.WriteLine("</div>");
                    }
                    writer.WriteLine("</body>");
                }
                writer.WriteLine("</html>");
            }
            ZipFile.CreateFromDirectory(tempDir, outputZip);
            return File(new FileStream(outputZip, FileMode.Open), "application/zip", Path.GetFileName(outputZip));
        }

        private async Task<string> BuildHtmlMessage(string chatId, string assetsDir, Microsoft.Graph.ChatMessage message, string messageClass = "message")
        {
            Debug.WriteLine("{0} [{1}] {2}", message.Id, message.CreatedDateTime, message.From?.User?.DisplayName);
            StringBuilder messageBuilder = new StringBuilder(8192);
            messageBuilder.AppendLine($"<div class='{messageClass}' data-id={message.Id}>");
            {
                messageBuilder.AppendLine($"<span class='from'>");
                {
                    messageBuilder.Append(message.From?.User.DisplayName);
                }
                messageBuilder.AppendLine("</span>");

                messageBuilder.AppendLine($"<span class='date'>");
                {
                    messageBuilder.Append(message.CreatedDateTime?.DateTime.ToLongDateString());
                    messageBuilder.Append(" ");
                    messageBuilder.Append(message.CreatedDateTime?.DateTime.ToLongTimeString());
                }
                messageBuilder.AppendLine("</span>");

                if (message.LastEditedDateTime != null && message.LastEditedDateTime != message.CreatedDateTime)
                {
                    messageBuilder.AppendLine($"<span class='date'>");
                    {
                        messageBuilder.Append(" Edit at: ");
                        messageBuilder.Append(message.LastEditedDateTime?.DateTime.ToLongDateString());
                        messageBuilder.Append(" ");
                        messageBuilder.Append(message.LastEditedDateTime?.DateTime.ToLongTimeString());
                    }
                    messageBuilder.AppendLine("</span>");
                }

                messageBuilder.AppendLine($"<div class='content'>");
                {
                    var messageContent = message.Body.Content;
                    if (messageContent.Contains($"https://graph.microsoft.com/v1.0/chats/{chatId}/messages/{message.Id}/hostedContents"))
                    {
                        foreach (var item in await GraphHelper.GetChatMessageHostedContentsAsync(chatId, message.Id))
                        {
                            var content = await GraphHelper.GetChatMessageHostedContentsAsync(chatId, message.Id, item.Id);

                            var currentContentTag = messageContent.Substring(0, messageContent.IndexOf($"https://graph.microsoft.com/v1.0/chats/{chatId}/messages/{message.Id}/hostedContents/{item.Id}"));
                            currentContentTag = currentContentTag.Substring(currentContentTag.LastIndexOf("<") + 1);
                            currentContentTag = currentContentTag.Substring(0, currentContentTag.IndexOf(" "));

                            var ext = "";
                            switch (currentContentTag)
                            {
                                case "img": ext = ".jpg"; break;
                                default: throw new Exception("Unknown tag");
                            }

                            var safeContentId = Hash(item.Id) + ext;

                            using (var contentStream = new FileStream(Path.Combine(assetsDir, safeContentId), FileMode.Create))
                                await content.CopyToAsync(contentStream);
                            messageContent = messageContent.Replace($"https://graph.microsoft.com/v1.0/chats/{chatId}/messages/{message.Id}/hostedContents/{item.Id}/$value", $"assets/{safeContentId}");
                        }
                    }

                    foreach (var attachment in message.Attachments)
                    {
                        if (attachment.ContentType == "messageReference")
                        {
                            var reply = JsonConvert.DeserializeObject<ChatReply>(attachment.Content);
                            var replyMessage = await GraphHelper.GetChatMessageAsync(chatId, reply.messageId);
                            messageContent = messageContent.Replace($"<attachment id=\"{attachment.Id}\"></attachment>", await BuildHtmlMessage(chatId, assetsDir, replyMessage, "reply"));
                        }
                        else if (!string.IsNullOrEmpty(attachment.ContentUrl))
                        {
                            var content = await GraphHelper.DownloadAttachmentAsync(message, attachment);
                            if (content != null)
                            {
                                var safeContentId = Hash(attachment.Id) + "_" + attachment.Name;

                                using (var contentStream = new FileStream(Path.Combine(assetsDir, safeContentId), FileMode.Create))
                                    await content.CopyToAsync(contentStream);

                                messageContent = messageContent.Replace($"<attachment id=\"{attachment.Id}\"></attachment>", $"<a href=\"assets/{safeContentId}\">{attachment.Name}</a>");
                            }
                            else
                            {
                                messageContent = messageContent.Replace($"<attachment id=\"{attachment.Id}\"></attachment>", $"<a href=\"#404\">*MISSING* {attachment.Name}</a>");
                            }
                        }
                    }
                    messageBuilder.Append(messageContent);
                }
                messageBuilder.AppendLine("</div>");

                if (message.MessageType != Microsoft.Graph.ChatMessageType.Message)
                {
                    messageBuilder.AppendLine($"<span class='type'>");
                    {
                        messageBuilder.Append(message.MessageType);
                    }
                    messageBuilder.AppendLine("</span>");

                    messageBuilder.AppendLine($"<span class='summary'>");
                    {
                        messageBuilder.Append(JsonConvert.SerializeObject(message));
                    }
                    messageBuilder.AppendLine("</span>");
                }
            }
            messageBuilder.AppendLine("</div>");
            messageBuilder.AppendLine("</br>");
            return messageBuilder.ToString();
        }

        public string Hash(string input)
        {
            using (var hash = SHA1.Create())
                return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(input))).Replace("-", "");
        }

        public class ChatReply
        {
            public string messageId { get; set; }
            public string messagePreview { get; set; }
            public Microsoft.Graph.IdentitySet messageSender { get; set; }
        }
    }
}