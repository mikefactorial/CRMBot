using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace CRMBot
{
    public class DataAccess
    {
        public static ConnectionInformationEntity RetrieveConnectionInformation(string userAddress)
        {
            /*
            string connectionString = ConfigurationManager.ConnectionStrings["AzureStorage"].ConnectionString;
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            // Create the table client.
            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();

            // Create the table if it doesn't exist.
            CloudTable table = tableClient.GetTableReference("ConnectionInformation");
            table.CreateIfNotExists();

            TableQuery<ConnectionInformationEntity> query = new TableQuery<ConnectionInformationEntity>().Where(TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, userAddress));
            return table.ExecuteQuery(query).FirstOrDefault();
            */
            return new ConnectionInformationEntity(userAddress, "https://whatever/XRMServices/2011/Organization.svc", "test@test.test", "password");
        }
    }
}
