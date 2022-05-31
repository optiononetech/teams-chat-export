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

        public static async Task<IEnumerable<Chat>> GetChatsAsync()
        {
            var graphClient = GetAuthenticatedClient();

            List<Chat> chats = new List<Chat>();

            var graphChats = await graphClient.Me.Chats.Request()
                .GetAsync();

            while (graphChats.Count > 0)
            {
                chats.AddRange(graphChats);
                if (graphChats.NextPageRequest == null)
                    break;
                graphChats = await graphChats.NextPageRequest.GetAsync();
            }

            foreach (var chat in chats)
                chat.Members = await graphClient.Me.Chats[chat.Id].Members.Request().GetAsync();

            return chats;
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
            return await graphClient.Me.Chats[chatId].Messages.Request().GetAsync();
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