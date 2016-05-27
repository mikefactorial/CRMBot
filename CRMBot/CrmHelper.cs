using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.ServiceModel.Description;
using System.Web;

namespace CRMBot
{
    public class CrmHelper
    {
        private const int MIN_TEXTLENGTHFORFIELDSEARCH = 4;
        private const int MIN_TEXTLENGTHFORENTITYSEARCH = 4;
        private static EntityMetadata[] metadata = null;
        public static EntityMetadata[] RetrieveMetadata(string conversationId)
        {
            if (metadata == null)
            {
                RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest();
                request.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.All;
                RetrieveAllEntitiesResponse response;
                using (OrganizationServiceProxy service = CreateOrganizationService(conversationId))
                {
                    response = (RetrieveAllEntitiesResponse)service.Execute(request);
                }
                metadata = response.EntityMetadata;
            }
            return metadata;
        }

        public static string FindEntity(string conversationId, string text)
        {
            EntityMetadata[] metadata = RetrieveMetadata(conversationId);
            //Equals
            string subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                EntityMetadata entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                EntityMetadata entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                EntityMetadata entity = metadata.FirstOrDefault(e => e.LogicalName == subText);
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                EntityMetadata entity = metadata.FirstOrDefault(e => e.LogicalName == subText);
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
                EntityMetadata entity = metadata.FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                EntityMetadata entity = metadata.FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                EntityMetadata entity = metadata.FirstOrDefault(e => e.LogicalName.Contains(subText));
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                EntityMetadata entity = metadata.FirstOrDefault(e => e.LogicalName.Contains(subText));
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

            //Equals
            string subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }

            //Contains
            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.LogicalName.Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.LogicalName.Contains(subText));
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
            return RetrieveMetadata(conversationId).FirstOrDefault(e => e.LogicalName == entityLogicalName);
        }
        public static OrganizationServiceProxy CreateOrganizationService(string conversationId)
        {
            ChatState state = ChatState.RetrieveChatState(conversationId);
            Uri oUri = new Uri(state.OrganizationServiceUrl);
            ClientCredentials clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = state.UserName;
            clientCredentials.UserName.Password = state.Password;
            clientCredentials.Windows.ClientCredential = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["UserName"], ConfigurationManager.AppSettings["Password"]);

            //Create your Organization Service Proxy  
            return new OrganizationServiceProxy(
                oUri,
                null,
                clientCredentials,
                null);

        }

    }
}