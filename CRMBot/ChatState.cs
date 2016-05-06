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
        private static double chatCacheDurationMinutes = 30.0000;

        public List<byte[]> Attachments { get; set; }

        public Entity SelectedEntity { get; set; }

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
    }
}