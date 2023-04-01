// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using teams_export.Models;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using teams_export.TokenStorage;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.IO;
using System.Net.Http;
using System;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using System.Web.Helpers;
using System.Net.Mail;

namespace teams_export.Helpers
{
    public static class GraphHelper
    {
        public static async Task<CachedUser> GetUserDetailsAsync(string accessToken)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                    }));

            var user = await graphClient.Me.Request()
                .Select(u => new
                {
                    u.DisplayName,
                    u.Mail,
                    u.UserPrincipalName
                })
                .GetAsync();

            return new CachedUser
            {
                Avatar = string.Empty,
                DisplayName = user.DisplayName,
                Email = string.IsNullOrEmpty(user.Mail) ?
                    user.UserPrincipalName : user.Mail
            };
        }

        // Load configuration settings from PrivateSettings.config
        private static string appId = ConfigurationManager.AppSettings["ida:AppId"];
        private static string appSecret = ConfigurationManager.AppSettings["ida:AppSecret"];
        private static string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
        private static List<string> graphScopes =
            new List<string>(ConfigurationManager.AppSettings["ida:AppScopes"].Split(' '));

        public static async Task<IEnumerable<Event>> GetEventsAsync()
        {
            var graphClient = GetAuthenticatedClient();

            var events = await graphClient.Me.Events.Request()
                .Select("subject,organizer,start,end")
                .OrderBy("createdDateTime DESC")
                .GetAsync();

            return events.CurrentPage;
        }

        public static async Task<IEnumerable<Chat>> GetChatsAsync(string actionId, HttpContextBase session)
        {
            var graphClient = GetAuthenticatedClient();

            List<Chat> chats = new List<Chat>();

            session.SetCurrentAction(actionId, "Getting initial chat list");
            var graphChats = await graphClient.Me.Chats.Request()
                .GetAsync();

            while (graphChats.Count > 0)
            {
                chats.AddRange(graphChats);
                session.SetResult(actionId, chats.ToList());
                if (graphChats.NextPageRequest == null) break;
                session.SetCurrentAction(actionId, $"Getting next chat list. Last chat: [{chats.LastOrDefault()?.Topic ?? "Private Chat"}] {chats.LastOrDefault()?.Id}");
                graphChats = await graphChats.NextPageRequest.GetAsync();
            }
            session.SetCurrentAction(actionId, "Done");
            return chats;
        }

        public static async Task<IEnumerable<ConversationMember>> GetChatMembersAsync(string id)
        {
            var graphClient = GetAuthenticatedClient();
            return await graphClient.Me.Chats[id].Members.Request().GetAsync();
        }

        public static async Task<Chat> GetChatAsync(string id)
        {
            var graphClient = GetAuthenticatedClient();

            var chat = await graphClient.Me.Chats[id].Request()
                .GetAsync();
            chat.Members = await graphClient.Me.Chats[chat.Id].Members.Request().GetAsync();

            return chat;
        }

        public static async Task<IChatMessagesCollectionPage> GetChatMessagesAsync(string chatId)
        {
            var graphClient = GetAuthenticatedClient();
            var request = graphClient.Me.Chats[chatId].Messages.Request();
            return await request.GetAsync();
        }

        public static string Hash(string input)
        {
            using (var hash = SHA1.Create())
                return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(input))).Replace("-", "");
        }

        public class ChatReply
        {
            public string messageId { get; set; }
            public string messagePreview { get; set; }
            public IdentitySet messageSender { get; set; }
        }

        public static async Task<List<ChatMessage>> GetChatMessagesAsync(string chatId, DateTime since, DateTime until, string publicDir, bool staticPath, string actionId, HttpContextBase session)
        {
            var graphClient = GetAuthenticatedClient();
            var request = graphClient.Me.Chats[chatId].Messages.Request();
            request = request.Filter($"lastModifiedDateTime gt {since:o}Z and lastModifiedDateTime lt {until.Date.AddDays(1).AddSeconds(-1):o}Z").Top(50);
            session.SetCurrentAction(actionId, "Get initial message list");
            var currentPage = await request.GetAsync();
            var messageList = new List<ChatMessage>();
            while (currentPage.Count > 0)
            {
                foreach (var message in currentPage)
                {
                    if (message.LastModifiedDateTime.Value.Date >= since)
                    {
                        session.SetCurrentAction(actionId, $"Getting next message list. Last message: {message.From?.User?.DisplayName}/{message.LastModifiedDateTime}");
                        messageList.Add(await DownloadChatMessage(chatId, message, publicDir, staticPath));
                    }
                    else
                        break;
                }
                session.SetResult(actionId, messageList);
                if (currentPage.NextPageRequest == null) break;
                currentPage = await currentPage.NextPageRequest.GetAsync();
            }
            session.SetCurrentAction(actionId, "Done");
            return messageList;
        }

        private static async Task<ChatMessage> DownloadChatMessage(string chatId, ChatMessage message, string publicDir, bool staticPath)
        {
            var messageContent = message.Body.Content;
            if (messageContent.Contains($"https://graph.microsoft.com/v1.0/chats/{chatId}/messages/{message.Id}/hostedContents"))
            {
                foreach (var item in await GetChatMessageHostedContentsAsync(chatId, message.Id))
                {
                    var content = await GetChatMessageHostedContentsAsync(chatId, message.Id, item.Id);

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
                    if (!System.IO.File.Exists(Path.Combine(publicDir, safeContentId)))
                    {
                        using (var contentStream = new FileStream(Path.Combine(publicDir, safeContentId), FileMode.Create))
                            await content.CopyToAsync(contentStream);
                    }
                    if (staticPath)
                    {
                        messageContent = messageContent.Replace($"https://graph.microsoft.com/v1.0/chats/{chatId}/messages/{message.Id}/hostedContents/{item.Id}/$value", $"assets/{safeContentId}");
                    }
                    else
                    {
                        var chatIdHash = Hash(chatId);
                        messageContent = messageContent.Replace($"https://graph.microsoft.com/v1.0/chats/{chatId}/messages/{message.Id}/hostedContents/{item.Id}/$value", $"/public?chatId={chatIdHash}&fileName={safeContentId}&downloadName={safeContentId}");
                    }
                }
            }

            foreach (var attachment in message.Attachments)
            {
                if (attachment.ContentType == "messageReference")
                {
                    var reply = JsonConvert.DeserializeObject<ChatReply>(attachment.Content);
                    var replyMessage = await GetChatMessageAsync(chatId, reply.messageId);
                    var processedReply = await DownloadChatMessage(chatId, replyMessage, publicDir, staticPath);
                    messageContent = messageContent.Replace($"<attachment id=\"{attachment.Id}\"></attachment>", processedReply.Body.Content);
                }
                else if (!string.IsNullOrEmpty(attachment.ContentUrl))
                {
                    var content = await DownloadAttachmentAsync(message, attachment);
                    if (content != null)
                    {
                        var safeContentId = Hash(attachment.Id) + "_" + attachment.Name;
                        if (!System.IO.File.Exists(Path.Combine(publicDir, safeContentId)))
                        {
                            using (var contentStream = new FileStream(Path.Combine(publicDir, safeContentId), FileMode.Create))
                                await content.CopyToAsync(contentStream);
                        }

                        string mimeType = MimeMapping.GetMimeMapping(attachment.Name);
                        if (mimeType.StartsWith("image"))
                        {
                            if (staticPath)
                            {
                                messageContent = messageContent.Replace($"<attachment id=\"{attachment.Id}\"></attachment>", $"<a href=\"assets/{safeContentId}\"><img style='width: 100%' src=\"assets/{safeContentId}\" />{attachment.Name}</a>");
                            }
                            else
                            {
                                var chatIdHash = Hash(chatId);
                                messageContent = messageContent.Replace($"<attachment id=\"{attachment.Id}\"></attachment>", $"<a href=\"/public?chatId={chatIdHash}&fileName={safeContentId}&downloadName={attachment.Name}\"><img style='width: 100%' src=\"/public?chatId={chatIdHash}&fileName={safeContentId}&downloadName={attachment.Name}\" />{attachment.Name}</a>");
                            }
                        }
                        else
                        {
                            if (staticPath)
                            {
                                messageContent = messageContent.Replace($"<attachment id=\"{attachment.Id}\"></attachment>", $"<a href=\"assets/{safeContentId}\">{attachment.Name}</a>");
                            }
                            else
                            {
                                var chatIdHash = Hash(chatId);
                                messageContent = messageContent.Replace($"<attachment id=\"{attachment.Id}\"></attachment>", $"<a href=\"/public?chatId={chatIdHash}&fileName={safeContentId}&downloadName={attachment.Name}\">{attachment.Name}</a>");
                            }
                        }
                    }
                    else
                    {
                        messageContent = messageContent.Replace($"<attachment id=\"{attachment.Id}\"></attachment>", $"<a href=\"#404\">*MISSING* {attachment.Name}</a>");
                    }
                }
            }
            message.Body.Content = messageContent;
            return message;
        }

        public static async Task<ChatMessage> GetChatMessageAsync(string chatId, string messageId)
        {
            var graphClient = GetAuthenticatedClient();
            return await graphClient.Me.Chats[chatId].Messages[messageId].Request().GetAsync();
        }

        public static async Task<IChatMessageRepliesCollectionPage> GetChatMessageRepliesAsync(string chatId, string messageId)
        {
            var graphClient = GetAuthenticatedClient();
            return await graphClient.Me.Chats[chatId].Messages[messageId].Replies.Request().GetAsync();
        }

        public static async Task<IChatMessageHostedContentsCollectionPage> GetChatMessageHostedContentsAsync(string chatId, string messageId)
        {
            var graphClient = GetAuthenticatedClient();
            return await graphClient.Me.Chats[chatId].Messages[messageId].HostedContents.Request().GetAsync();
        }

        public static async Task<Stream> GetChatMessageHostedContentsAsync(string chatId, string messageId, string hostedContentId)
        {
            var graphClient = GetAuthenticatedClient();
            return await graphClient.Me.Chats[chatId].Messages[messageId].HostedContents[hostedContentId].Content.Request().GetAsync();
        }

        /*
 function getDownloadUrl(attachment, team) {
  const url = new URL(attachment.contentUrl);
  const path = url.pathname.split('/').slice(4).join('/');
  if (team) {
    return `/groups/${team.id}/drive/root:/${path}`;
  } else {
    return `/me/drive/root:/${path}`;
  }
}
         */

        public static async Task<Stream> DownloadAttachmentAsync(ChatMessage message, ChatMessageAttachment attachment)
        {
            var graphClient = GetAuthenticatedClient();
            var me = await graphClient.Me.Request().GetAsync();
            if (me.Id == message.From?.User?.Id)
            {
                var fileUrl = new Uri(attachment.ContentUrl);
                var mePath = fileUrl.Segments.Skip(4).Aggregate((a, b) => $"{a}{b}");
                return await graphClient.Me.Drive.Root.ItemWithPath(mePath).Content.Request().GetAsync();
            }
            else
            {
                var teams = await graphClient.Me.JoinedTeams.Request().GetAsync();
                var sharedWithMe = await graphClient.Me.Drive.SharedWithMe().Request().GetAsync();
                var attachmentUrl = attachment.ContentUrl.ToUpper();
                var attachmentId = attachment.Id.ToUpper();

                var driveFile = sharedWithMe.FirstOrDefault(file => HttpUtility.UrlDecode(file.WebUrl).ToUpper().Contains($"{{{attachmentId}}}"));
                if (driveFile == null) driveFile = sharedWithMe.FirstOrDefault(file => HttpUtility.UrlDecode(file.WebUrl).ToUpper().Equals($"{{{attachmentId}}}"));
                if (driveFile == null) driveFile = sharedWithMe.FirstOrDefault(file => HttpUtility.UrlDecode(file.WebUrl).ToUpper().StartsWith($"{{{attachmentId}}}"));

                if (driveFile != null)
                {
                    if (driveFile.RemoteItem != null)
                    {
                        if (!string.IsNullOrWhiteSpace(driveFile.RemoteItem.SharepointIds.SiteId))
                        {
                            return await graphClient.Sites[driveFile.RemoteItem.SharepointIds.SiteId].Drive.Items[driveFile.Id].Content.Request().GetAsync();
                        }
                    }
                }
                return null;
            }
        }

        private static GraphServiceClient GetAuthenticatedClient()
        {
            return new GraphServiceClient(GetAuthenticationProvider());
        }
        private static DelegateAuthenticationProvider GetAuthenticationProvider()
        {
            return new DelegateAuthenticationProvider(
                   async (requestMessage) =>
                   {
                       var idClient = ConfidentialClientApplicationBuilder.Create(appId)
                           .WithRedirectUri(redirectUri)
                           .WithClientSecret(appSecret)
                           .Build();

                       var tokenStore = new SessionTokenStore(idClient.UserTokenCache,
                               HttpContext.Current, ClaimsPrincipal.Current);

                       var accounts = await idClient.GetAccountsAsync();

                       // By calling this here, the token can be refreshed
                       // if it's expired right before the Graph call is made
                       var result = await idClient.AcquireTokenSilent(graphScopes, accounts.FirstOrDefault())
                                  .ExecuteAsync();

                       requestMessage.Headers.Authorization =
                           new AuthenticationHeaderValue("Bearer", result.AccessToken);
                   });
        }
    }
}