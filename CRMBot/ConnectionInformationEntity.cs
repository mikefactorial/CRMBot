using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CRMBot
{
    public class ConnectionInformationEntity : TableEntity
    {
        public ConnectionInformationEntity(string userAddress, string organizationServiceUrl, string userName, string password)
        {
            this.RowKey = userAddress;
            this.PartitionKey = organizationServiceUrl;
            this.UserName = userName;
            this.Password = password;
        }

        public ConnectionInformationEntity()
        {
        }

        public string OrganizationServiceUrl
        {
            get
            {
                return this.PartitionKey;
            }
        }

        public string UserAddress
        {
            get
            {
                return this.RowKey;
            }
        }

        public string UserName { get; set; }
        public string Password { get; set; }
    }
}