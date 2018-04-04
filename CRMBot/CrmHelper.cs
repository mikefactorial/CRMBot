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

        public static string FindEntityLogicalName(string conversationId, string text)
        {
            EntityMetadata[] metadata = ChatState.RetrieveChatState(conversationId).RetrieveMetadata();
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
        public static EntityMetadata RetrieveEntityMetadata(string conversationId, string entityLogicalName)
        {
            return ChatState.RetrieveChatState(conversationId).RetrieveEntityMetadata(entityLogicalName);
        }

        public static OrganizationWebProxyClient CreateOrganizationService(string conversationId)
        {
            if (conversationId != Guid.Empty.ToString())
            {
                ChatState state = ChatState.RetrieveChatState(conversationId);
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