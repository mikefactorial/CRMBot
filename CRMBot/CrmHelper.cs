using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.ServiceModel.Description;
using System.Text;
using System.Web;

namespace CRMBot
{
    public class CrmHelper
    {
        private const int MIN_TEXTLENGTHFORFIELDSEARCH = 4;
        private const int MIN_TEXTLENGTHFORENTITYSEARCH = 4;
        private static Entity defaultSettings = null;

        public static string FindEntity(string conversationId, string text)
        {
            EntityMetadata[] metadata = ChatState.RetrieveChatState(conversationId).Metadata;
            string subText = text.ToLower();

            //Equals
            EntityMetadata entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
            if (entity != null)
            {
                return entity.LogicalName;
            }
            entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
            if (entity != null)
            {
                return entity.LogicalName;
            }
            entity = metadata.FirstOrDefault(e => e.LogicalName == subText);
            if (entity != null)
            {
                return entity.LogicalName;
            }
            entity = metadata.FirstOrDefault(e => e.LogicalName == subText);
            if (entity != null)
            {
                return entity.LogicalName;
            }
            //Substring Equals
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                entity = metadata.FirstOrDefault(e => e.LogicalName == subText);
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                entity = metadata.FirstOrDefault(e => e.LogicalName == subText);
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }

            //Contains
            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                entity = metadata.FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                entity = metadata.FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                entity = metadata.FirstOrDefault(e => e.LogicalName.Contains(subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                entity = metadata.FirstOrDefault(e => e.LogicalName.Contains(subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }

            return string.Empty;

        }
        public static string FindAttribute(EntityMetadata entity, string text)
        {
            string subText = text.ToLower();
            //Equals
            AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
            if (att != null)
            {
                return att.LogicalName;
            }
            att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
            if (att != null)
            {
                return att.LogicalName;
            }
            att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
            if (att != null)
            {
                return att.LogicalName;
            }
            att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
            if (att != null)
            {
                return att.LogicalName;
            }

            //Substring Equals

            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }

            //Substring Contains
            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.LogicalName.Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.LogicalName.Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            return string.Empty;
        }

        public static Entity DefaultSettings
        {
            get
            {
                if (defaultSettings == null)
                {
                    QueryExpression query = new QueryExpression("cobalt_settings");
                    query.ColumnSet = new ColumnSet(new string[] { "cobalt_settingsid", "cobalt_organizationdeskey" });
                    query.PageInfo = new PagingInfo()
                    {
                        PageNumber = 1,
                        Count = 1
                    };

                    using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(Guid.Empty.ToString()))
                    {
                        EntityCollection collection = serviceProxy.RetrieveMultiple(query);
                        if (collection.Entities != null && collection.Entities.Count == 1)
                        {
                            defaultSettings = collection.Entities[0];
                        }
                    }
                }

                return defaultSettings;
            }
        }
        public static EntityMetadata RetrieveEntityMetadata(string conversationId, string entityLogicalName)
        {
            return ChatState.RetrieveChatState(conversationId).Metadata.FirstOrDefault(e => e.LogicalName == entityLogicalName);
        }
        public static OrganizationServiceProxy CreateOrganizationService(string conversationId)
        {
            Uri oUri;
            ClientCredentials clientCredentials = new ClientCredentials();
            if (conversationId != Guid.Empty.ToString())
            {
                ChatState state = ChatState.RetrieveChatState(conversationId);
                oUri = new Uri(state.OrganizationServiceUrl);
                clientCredentials.UserName.UserName = state.UserName;
                clientCredentials.UserName.Password = state.Password;
                clientCredentials.Windows.ClientCredential = new System.Net.NetworkCredential(state.UserName, state.Password);
            }
            else
            {
                oUri = new Uri(ConfigurationManager.AppSettings["OrganizationServiceUrl"]);
                clientCredentials.UserName.UserName = ConfigurationManager.AppSettings["UserName"];
                clientCredentials.UserName.Password = ConfigurationManager.AppSettings["Password"];
                clientCredentials.Windows.ClientCredential = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["UserName"], ConfigurationManager.AppSettings["Password"]);
            }
            //Create your Organization Service Proxy  
            return new OrganizationServiceProxy(
                oUri,
                null,
                clientCredentials,
                null);

        }

    }
}