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

namespace teams_export.Controllers
{
    public class ChatController : BaseController
    {
        // GET: Chat
        [Authorize]
        public async Task<ActionResult> Index()
        {
            var chats = await GraphHelper.GetChatsAsync();
            return View(chats);
        }


        // GET: Chat
        [Authorize]
        public async Task<ActionResult> Export(string id)
        {
            var chat = await GraphHelper.GetChatAsync(id);
            var currentPage = await GraphHelper.GetChatMessagesAsync(id);

            var safeChatId = Hash(id);
            var tempDir = Path.Combine(AppContext.BaseDirectory, "export", "raw", safeChatId);
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            var assetsDir = Path.Combine(tempDir, "assets");
            if (!Directory.Exists(assetsDir)) Directory.CreateDirectory(assetsDir);
            var outputZip = Path.Combine(AppContext.BaseDirectory, "export", $"{safeChatId}.zip");
            if (!System.IO.File.Exists(outputZip))
            {
                if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
                var htmlDoc = Path.Combine(tempDir, "index.html");
                using (var stream = new FileStream(htmlDoc, FileMode.Create))
                using (var writer = new StreamWriter(stream, Encoding.UTF8, 81920, true))
                {
                    writer.WriteLine("<!doctype HTML>");
                    writer.WriteLine("<html>");
                    {
                        writer.WriteLine("<head>");
                        {
                            writer.WriteLine($"<title>{id} - {chat.Topic}</title>");
                            writer.WriteLine($"<style>");
                            {
                                writer.WriteLine(
@"
    .messages {
        background-color: rgb(245, 245, 245);
    }

    .message{
        background-color: white;
    }
    .message, .reply {
        display: inline-block;
        padding: 10px;
        margin: 10px;
    }
    .from {
        font-weight: bold;
    }

    .date {
        color: gray;
    }
    .reply {
        background-color: rgb(250, 249, 248);
        border: 1px solid silver;
    }
");
                            }
                            writer.WriteLine($"</style>");
                        }
                        writer.WriteLine("</head>");
                        writer.WriteLine("<body>");
                        {
                            writer.WriteLine("<h2>Members</h2>");
                            foreach (var member in chat.Members)
                                writer.WriteLine($"<h4>{member?.DisplayName}</h4>");

                            writer.WriteLine("<h2>Messages</h2>");
                            writer.WriteLine("<div class='messages'>");
                            {
                                while (currentPage.Count > 0)
                                {
                                    foreach (var message in currentPage)
                                    {
                                        writer.Write(await BuildHtmlMessage(id, assetsDir, message));
                                    }
                                    if (currentPage.NextPageRequest == null)
                                        break;
                                    currentPage = await currentPage.NextPageRequest.GetAsync();
                                }
                            }
                            writer.WriteLine("</div>");
                        }
                        writer.WriteLine("</body>");
                    }
                    writer.WriteLine("</html>");
                }
                ZipFile.CreateFromDirectory(tempDir, outputZip);
            }
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
                                throw new Exception("Unknown file");
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
            using (var hash = SHA256.Create())
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