using Microsoft.Bot.Builder.Luis.Models;
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
using System.Text.RegularExpressions;
using System.Web;

namespace CRMBot
{
    public class CrmHelper
    {
        public static EntityMetadata RetrieveEntityMetadata(string channelId, string userId, string entityLogicalName)
        {
            return ChatState.RetrieveChatState(channelId, userId).RetrieveEntityMetadata(entityLogicalName);
        }

        public static string ParseCrmUrl(Activity message)
        {
            if (message.From.Properties.ContainsKey("crmUrl"))
            {
                return message.From.Properties["crmUrl"].ToString();
            }
            else if(!string.IsNullOrEmpty(message.Text))
            {
                var regex = new Regex("<a [^>]*href=(?:'(?<href>.*?)')|(?:\"(?<href>.*?)\")", RegexOptions.IgnoreCase);
                var urls = regex.Matches(message.Text).OfType<Match>().Select(m => m.Groups["href"].Value).ToList();
                if (urls.Count > 0)
                {
                    return urls[0];
                }
                else if (message.Text.ToLower().StartsWith("http") && message.Text.ToLower().Contains(".dynamics.com"))
                {
                    return message.Text;
                }
            }
            return string.Empty;

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