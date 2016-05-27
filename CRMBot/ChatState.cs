using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Web;
using Microsoft.WindowsAzure.Storage.Table;
using Microsoft.Bot.Connector;
using System.Configuration;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace CRMBot
{
    public class ChatState
    {
        public Dictionary<string, object> Data = new Dictionary<string, object>();

        private static double chatCacheDurationMinutes = 30.0000;

        public static string Attachments = "Attachments";
        public static string FilteredEntities = "FilteredEntities";
        public static string SelectedEntity = "SelectedEntity";

        public static bool SetChatState(Message message)
        {
            bool returnValue = false;
            if (!MemoryCache.Default.Contains(message.ConversationId))
            {
                if (message.From != null)
                {
                    if (message.From.ChannelId.ToLower() == "facebook" || message.From.ChannelId.ToLower() == "skype")
                    {
                        QueryExpression query = new QueryExpression("cobalt_crmorganization");
                        query.ColumnSet = new ColumnSet(new string[] { "cobalt_organizationurl", "cobalt_username", "cobalt_password" });
                        if (message.From.ChannelId.ToLower() == "facebook")
                        {
                            query.Criteria.AddCondition("cobalt_facebookmessengerid", ConditionOperator.Equal, message.From.Id);
                        }
                        else if (message.From.ChannelId.ToLower() == "skype")
                        {
                            query.Criteria.AddCondition("cobalt_skypeid", ConditionOperator.Equal, message.From.Id);
                        }

                        using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(Guid.Empty.ToString()))
                        {
                            EntityCollection collection = serviceProxy.RetrieveMultiple(query);
                            if (collection.Entities != null && collection.Entities.Count == 1)
                            {
                                CacheItemPolicy policy = new CacheItemPolicy();
                                policy.Priority = CacheItemPriority.Default;
                                policy.SlidingExpiration = TimeSpan.FromMinutes(chatCacheDurationMinutes);

                                ChatState state = new ChatState();
                                state.OrganizationServiceUrl = (string)collection.Entities[0]["cobalt_organizationurl"];
                                state.OrganizationServiceUrl += (!state.OrganizationServiceUrl.EndsWith("/")) ? "/XRMServices/2011/Organization.svc" : "XRMServices/2011/Organization.svc";
                                state.UserName = (string)collection.Entities[0]["cobalt_username"];
                                state.Password = (string)collection.Entities[0]["cobalt_password"];
                                MemoryCache.Default.Add(message.ConversationId, state, policy);
                                returnValue = true;
                            }
                        }
                    }
                }
            }
            else
            {
                returnValue = true;
            }
            return returnValue;
        }
        public static ChatState RetrieveChatState(string conversationId)
        {
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