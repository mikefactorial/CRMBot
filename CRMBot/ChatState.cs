﻿using Microsoft.Xrm.Sdk;
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

namespace CRMBot
{
    public class ChatState
    {
        public Dictionary<string, object> Data = new Dictionary<string, object>();
        private EntityMetadata[] metadata = null;
        private static double chatCacheDurationMinutes = 30.0000;
        private string conversationId = string.Empty;
        private object metadataLock = new object();
        public static string Attachments = "Attachments";
        public static string FilteredEntities = "FilteredEntities";
        public static string SelectedEntity = "SelectedEntity";

        public ChatState(string conversationId)
        {
            this.conversationId = conversationId;
        }
        public static bool SetChatState(Message message)
        {
            bool returnValue = false;
            //MODEBUG TODO
            if (message.From.ChannelId.ToString() != "facebook" && message.From.ChannelId.ToString() != "skype")
            {
                CacheItemPolicy policy = new CacheItemPolicy();
                policy.Priority = CacheItemPriority.Default;
                policy.SlidingExpiration = TimeSpan.FromMinutes(chatCacheDurationMinutes);

                ChatState state = new ChatState(message.ConversationId);
                state.OrganizationServiceUrl = ConfigurationManager.AppSettings["OrganizationServiceUrl"];
                state.UserName = ConfigurationManager.AppSettings["UserName"];
                state.Password = ConfigurationManager.AppSettings["Password"];

                MemoryCache.Default.Add(message.ConversationId, state, policy);
                returnValue = true;
            }
            else if (!MemoryCache.Default.Contains(message.ConversationId))
            {
                if (message.From != null)
                {
                    if (message.From.ChannelId.ToLower() == "facebook" || message.From.ChannelId.ToLower() == "skype")
                    {
                        QueryExpression query = new QueryExpression("cobalt_crmorganization");
                        query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 533470000);
                        query.ColumnSet = new ColumnSet(new string[] { "cobalt_organizationurl", "cobalt_username", "cobalt_encryptedpassword" });
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
                            if (collection.Entities != null && collection.Entities.Count > 0 )
                            {
                                CacheItemPolicy policy = new CacheItemPolicy();
                                policy.Priority = CacheItemPriority.Default;
                                policy.SlidingExpiration = TimeSpan.FromMinutes(chatCacheDurationMinutes);

                                ChatState state = new ChatState(message.ConversationId);
                                state.OrganizationUrl = (string)collection.Entities[0]["cobalt_organizationurl"];
                                state.OrganizationServiceUrl = (string)collection.Entities[0]["cobalt_organizationurl"];
                                state.OrganizationServiceUrl += (!state.OrganizationServiceUrl.EndsWith("/")) ? "/XRMServices/2011/Organization.svc" : "XRMServices/2011/Organization.svc";
                                state.UserName = (string)collection.Entities[0]["cobalt_username"];
                                state.Password = Decrypt((string)collection.Entities[0]["cobalt_encryptedpassword"]);
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
        public static void ClearChatState(string conversationId)
        {
        }
        public static ChatState RetrieveChatState(string conversationId)
        {
            return MemoryCache.Default[conversationId] as ChatState;
        }

        private static string Decrypt(string cryptedString)
        {
            if (String.IsNullOrEmpty(cryptedString))
            {
                throw new ArgumentNullException("The string which needs to be decrypted can not be null.");
            }


            if (CrmHelper.DefaultSettings["cobalt_organizationdeskey"] != null && !string.IsNullOrEmpty(CrmHelper.DefaultSettings["cobalt_organizationdeskey"].ToString()))
            {
                DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
                cryptoProvider.Key = ASCIIEncoding.ASCII.GetBytes(CrmHelper.DefaultSettings["cobalt_organizationdeskey"].ToString());
                cryptoProvider.IV = ASCIIEncoding.ASCII.GetBytes(ConfigurationManager.AppSettings["IV"]);

                cryptedString = cryptedString.Replace(" ", "+");

                MemoryStream memoryStream = new MemoryStream
                        (Convert.FromBase64String(cryptedString));
                CryptoStream cryptoStream = new CryptoStream(memoryStream,
                    cryptoProvider.CreateDecryptor(cryptoProvider.Key, cryptoProvider.IV), CryptoStreamMode.Read);
                StreamReader reader = new StreamReader(cryptoStream);
                return reader.ReadToEnd();
            }
            return cryptedString;
        }

        public string OrganizationUrl
        {
            get; set;
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
        public EntityMetadata[] Metadata
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
                            request.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.All;
                            RetrieveAllEntitiesResponse response;
                            using (OrganizationServiceProxy service = CrmHelper.CreateOrganizationService(conversationId))
                            {
                                response = (RetrieveAllEntitiesResponse)service.Execute(request);
                            }

                            this.metadata = response.EntityMetadata;
                        }
                    }
                }
                return metadata;
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