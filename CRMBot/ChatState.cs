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
using System.Security.Cryptography;
using System.IO;
using System.Text;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.WebServiceClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Crm.Sdk.Messages;

namespace CRMBot
{
    public class ChatState
    {
        public Dictionary<string, object> Data = new Dictionary<string, object>();
        private Dictionary<string, EntityMetadata> metadata = null;
        private static double chatCacheDurationMinutes = 30.0000;
        private string channelId = string.Empty;
        private string userId = string.Empty;
        private object metadataLock = new object();
        private Microsoft.Xrm.Sdk.Entity userEntity = null;
        public static string Attachments = "Attachments";
        public static string FilteredEntities = "FilteredEntities";
        public static string SelectedEntity = "SelectedEntity";
        public static string CurrentPageIndex = "CurrentPageIndex";

        protected ChatState(string channelId, string userId)
        {
            this.channelId = channelId;
            this.userId = userId;
        }
        public static bool IsChatStateSet(Activity message)
        {
            if (!MemoryCache.Default.Contains(message.Conversation.Id))
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        public static void ClearChatState(string channelId, string userId)
        {
            if (MemoryCache.Default.Contains(channelId + userId))
            {
                MemoryCache.Default.Remove(channelId + userId);
            }
        }
        public static ChatState RetrieveChatState(string channelId, string userId)
        {
            if (!MemoryCache.Default.Contains(channelId + userId))
            {
                CacheItemPolicy policy = new CacheItemPolicy();
                policy.Priority = CacheItemPriority.Default;
                policy.SlidingExpiration = TimeSpan.FromMinutes(chatCacheDurationMinutes);
                ChatState state = new ChatState(channelId, userId);
                MemoryCache.Default.Add(channelId + userId, state, policy);
            }
            return MemoryCache.Default[channelId + userId] as ChatState;
        }

        public string AccessToken
        {
            get; set;
        }

        public string OrganizationUrl
        {
            get; set;
        }
        public string UserFirstName
        {
            get
            {
                if (userEntity == null)
                {
                    using (var service = CrmHelper.CreateOrganizationService(this.channelId, this.userId))
                    {
                        // Display information about the logged on user.
                        Guid userid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
                        userEntity = service.Retrieve("systemuser", userid,
                            new ColumnSet(new string[] { "firstname" }));
                    }
                }
                if (userEntity["firstname"] != null)
                {
                    return userEntity["firstname"].ToString();
                }
                return "Friend";
            }
        }
        public EntityMetadata RetrieveEntityMetadata(string entityLogicalName)
        {
            EntityMetadata entity = this.Metadata.FirstOrDefault(e => e.LogicalName == entityLogicalName);
            if (entity != null)
            {
                if (entity.Attributes == null || entity.Attributes.Length == 0)
                {
                    RetrieveEntityRequest request = new RetrieveEntityRequest();
                    request.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.All;
                    request.LogicalName = entityLogicalName;
                    RetrieveEntityResponse response;
                    using (OrganizationWebProxyClient service = CrmHelper.CreateOrganizationService(channelId, userId))
                    {
                        response = (RetrieveEntityResponse)service.Execute(request);
                        if (response != null)
                        {
                            this.metadata[entityLogicalName] = response.EntityMetadata;
                            entity = this.Metadata.FirstOrDefault(e => e.LogicalName == entityLogicalName);
                        }
                    }
                }
            }
            return entity;
        }
        public EntityMetadata[] RetrieveMetadata()
        {
            return this.Metadata;
        }
        private EntityMetadata[] Metadata
        {
            get
            {
                if (metadata == null)
                {
                    lock (metadataLock)
                    {
                        if (metadata == null)
                        {
                            RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest();
                            request.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity;
                            RetrieveAllEntitiesResponse response;
                            using (OrganizationWebProxyClient service = CrmHelper.CreateOrganizationService(channelId, userId))
                            {
                                response = (RetrieveAllEntitiesResponse)service.Execute(request);
                            }

                            this.metadata = new Dictionary<string, EntityMetadata>();
                            foreach (EntityMetadata metadata in response.EntityMetadata)
                            {
                                this.metadata.Add(metadata.LogicalName, metadata);
                            }
                        }
                    }
                }
                return metadata.Values.ToArray();
            }
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