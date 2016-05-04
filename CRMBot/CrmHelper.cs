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
        public static EntityMetadata RetrieveEntityMetadata(string entityLogicalName)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest();
            request.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity;
            request.LogicalName = entityLogicalName;
            RetrieveEntityResponse response;
            using (OrganizationServiceProxy service = CreateOrganizationService())
            {
                response = (RetrieveEntityResponse)service.Execute(request);
            }
            return response.EntityMetadata;
        }

        public static EntityMetadata RetrieveEntityRelationships(string entityLogicalName)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest();
            request.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Relationships;
            request.LogicalName = entityLogicalName;
            RetrieveEntityResponse response;
            using (OrganizationServiceProxy service = CreateOrganizationService())
            {
                response = (RetrieveEntityResponse)service.Execute(request);
            }
            return response.EntityMetadata;
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