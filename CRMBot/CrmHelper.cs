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
        public static EntityMetadata[] RetrieveMetadata()
        {
            if (metadata == null)
            {
                RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest();
                request.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.All;
                RetrieveAllEntitiesResponse response;
                using (OrganizationServiceProxy service = CreateOrganizationService())
                {
                    response = (RetrieveAllEntitiesResponse)service.Execute(request);
                }
                metadata = response.EntityMetadata;
            }
            return metadata;
        }

        public static string FindEntity(string text)
        {
            EntityMetadata[] metadata = RetrieveMetadata();
            //Equals
            string subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                EntityMetadata entity = metadata.FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
                if (entity != null)
                {
                    return entity.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
            {
                EntityMetadata entity = metadata.FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
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

        public static EntityMetadata RetrieveEntityMetadata(string entityLogicalName)
        {
            return RetrieveMetadata().FirstOrDefault(e => e.LogicalName == entityLogicalName);
        }
        public static OrganizationServiceProxy CreateOrganizationService()
        {
            Uri oUri = new Uri(ConfigurationManager.AppSettings["OrganizationServiceUrl"]);
            ClientCredentials clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = ConfigurationManager.AppSettings["UserName"];
            clientCredentials.UserName.Password = ConfigurationManager.AppSettings["Password"];
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