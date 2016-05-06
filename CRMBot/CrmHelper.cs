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