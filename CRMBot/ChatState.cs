using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Web;
using Microsoft.WindowsAzure.Storage.Table;
using Microsoft.Bot.Connector;

namespace CRMBot
{
    public class ChatState
    {
        public Dictionary<string, object> Data = new Dictionary<string, object>();

        private static double chatCacheDurationMinutes = 30.0000;

        public static string Attachments = "Attachments";
        public static string SelectedEntity = "SelectedEntity";

        public static ChatState RetrieveChatState(string chatId)
        {
            if (!MemoryCache.Default.Contains(chatId))
            {
                CacheItemPolicy policy = new CacheItemPolicy();
                policy.Priority = CacheItemPriority.Default;
                policy.SlidingExpiration = TimeSpan.FromMinutes(chatCacheDurationMinutes);

                CacheItem item = new CacheItem(chatId, new ChatState());
                MemoryCache.Default.Add(chatId, new ChatState(), policy);
            }
            return MemoryCache.Default[chatId] as ChatState;
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