using Microsoft.Bot.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.WebServiceClient;
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

        public static string FindEntityLogicalName(string channelId, string userId, string text)
        {
            EntityMetadata[] metadata = ChatState.RetrieveChatState(channelId, userId).RetrieveMetadata();
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
        public static string FindAttributeLogicalName(EntityMetadata entity, string text)
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
        public static EntityMetadata RetrieveEntityMetadata(string channelId, string userId, string entityLogicalName)
        {
            return ChatState.RetrieveChatState(channelId, userId).RetrieveEntityMetadata(entityLogicalName);
        }

        public static OrganizationWebProxyClient CreateOrganizationService(string channelId, string userId)
        {
            if (!string.IsNullOrEmpty(channelId) && !string.IsNullOrEmpty(userId))
            {
                ChatState state = ChatState.RetrieveChatState(channelId, userId);
                //Create your Organization Service Proxy  
                OrganizationWebProxyClient service =  new OrganizationWebProxyClient(
                    new Uri(state.OrganizationUrl + @"/xrmservices/2011/organization.svc/web?SdkClientVersion=8.2"),
                    false);
                service.HeaderToken = state.AccessToken;
                return service;
            }
            else
            {
                throw new Exception("Cannot connect to the organization");
            }
        }
    }
}