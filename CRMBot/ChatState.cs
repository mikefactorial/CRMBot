using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Web;
using Microsoft.WindowsAzure.Storage.Table;
using Microsoft.Bot.Connector;
using System.Configuration;

namespace CRMBot
{
    public class ChatState
    {
        public Dictionary<string, object> Data = new Dictionary<string, object>();

        private static double chatCacheDurationMinutes = 30.0000;

        public static string Attachments = "Attachments";
        public static string FilteredEntities = "FilteredEntities";
        public static string SelectedEntity = "SelectedEntity";

        public static ChatState RetrieveChatState(string conversationId)
        {
            if (!MemoryCache.Default.Contains(conversationId))
            {
                CacheItemPolicy policy = new CacheItemPolicy();
                policy.Priority = CacheItemPriority.Default;
                policy.SlidingExpiration = TimeSpan.FromMinutes(chatCacheDurationMinutes);

                ChatState state = new ChatState();
                //MODEBUG TODO Pull from cobaltlab
                state.OrganizationServiceUrl = ConfigurationManager.AppSettings["OrganizationServiceUrl"];
                state.UserName = ConfigurationManager.AppSettings["UserName"];
                state.Password = ConfigurationManager.AppSettings["Password"];
                MemoryCache.Default.Add(conversationId, state, policy);
            }
            return MemoryCache.Default[conversationId] as ChatState;
        }

        public string OrganizationServiceUrl
        {
            get; set;
        }
        public string UserName
        {
            get; set;
        }
        public string Password
        {
            get; set;
        }
        public void Set(string key, object data)
        {
            Data[key] = data;
        }
        public object Get(string key)
        {
            if (Data.ContainsKey(key))
            {
                return Data[key];
            }
            return null;
        }
    }
}