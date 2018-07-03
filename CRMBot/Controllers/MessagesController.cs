using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using Microsoft.Bot.Connector;

using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;

using Newtonsoft.Json;
using RestSharp;
using System.Configuration;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using System.Web;
using System.Text.RegularExpressions;
using CRMBot.Dialogs;

namespace CRMBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        [ResponseType(typeof(void))]
        public async Task<HttpResponseMessage> Post([FromBody]Activity message)
        {
            /* Bot out of service message
            if (message.Type == ActivityTypes.Message)
            {
                ConnectorClient connector = new ConnectorClient(new Uri(message.ServiceUrl));
                // return our reply to the user
                Activity reply = message.CreateReply("Hi there. Dynamics Bot is undergoing an upgrade of it's framework version to better serve your CRM needs. Check back soon. CRM you later...");
                await connector.Conversations.ReplyToActivityAsync(reply);
            }
            else
            {
                //HandleSystemMessage(activity);
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            return response;
            MicrosoftAppCredentials.TrustServiceUrl(message.ServiceUrl, DateTime.MaxValue);
            */
            ConnectorClient connector = new ConnectorClient(new Uri(message.ServiceUrl));

            try
            {
                if (message.Type == ActivityTypes.Message)
                {
                    await connector.Conversations.ReplyToActivityAsync(CreateTypingMessage(message));
                    ChatState state = ChatState.RetrieveChatState(message.ChannelId, message.From.Id);

                    string crmUrl = CrmDialog.ParseCrmUrl(message);
                    if (string.IsNullOrEmpty(state.OrganizationUrl) && CrmDialog.ParseCrmUrl(message) == string.Empty)
                    {
                        await connector.Conversations.ReplyToActivityAsync(message.CreateReply("Hi there, before we can work together you need to tell me your Dynamics 365 URL (e.g. https://contoso.crm.dynamics.com)"));
                    }
                    else if (string.IsNullOrEmpty(state.AccessToken) || (!string.IsNullOrEmpty(crmUrl) && crmUrl != state.OrganizationUrl))
                    {
                        string extraQueryParams = string.Empty;
                        if (crmUrl != string.Empty && state.OrganizationUrl != crmUrl)
                        {
                            if (!string.IsNullOrEmpty(state.OrganizationUrl) && state.OrganizationUrl != crmUrl)
                            {
                                //extraQueryParams = HttpUtility.UrlEncode("prompt=login");
                            }
                            state.OrganizationUrl = crmUrl;
                        }

                        string applicationUrl = HttpContext.Current.Request.Url.Scheme + "://" + HttpContext.Current.Request.Url.Authority;
                        if (message.From.Properties.ContainsKey("crmUrl"))
                        {
                            await connector.Conversations.SendToConversationAsync(message.CreateReply("Give me a second. I'm logging you in..."));

                            Activity backChannelReply = message.CreateReply();
                            backChannelReply.Text = $"{applicationUrl}/Home/Login?channelId={HttpUtility.UrlEncode(message.ChannelId)}&userId={HttpUtility.UrlEncode(message.From.Id)}&userName={HttpUtility.UrlEncode(message.From.Name)}&fromId={HttpUtility.UrlEncode(message.Recipient.Id)}&fromName={HttpUtility.UrlEncode(message.Recipient.Name)}&serviceUrl={HttpUtility.UrlEncode(message.ServiceUrl)}&conversationId={HttpUtility.UrlEncode(message.Conversation.Id)}&extraQueryParams={extraQueryParams}";
                            backChannelReply.Name = "openUrl";
                            backChannelReply.Recipient = message.From;
                            backChannelReply.Type = ActivityTypes.Event;
                            await connector.Conversations.SendToConversationAsync(backChannelReply);
                        }
                        else
                        {
                            Activity replyToConversation = message.CreateReply();
                            replyToConversation.Recipient = message.From;
                            replyToConversation.Type = "message";
                            replyToConversation.Attachments = new List<Attachment>();


                            List<CardAction> cardButtons = new List<CardAction>();
                            CardAction plButton = new CardAction()
                            {
                                // ASP.NET Web Application Hosted in Azure
                                // Pass the user id
                                Value = $"{applicationUrl}/Home/Login?channelId={HttpUtility.UrlEncode(message.ChannelId)}&userId={HttpUtility.UrlEncode(message.From.Id)}&userName={HttpUtility.UrlEncode(message.From.Name)}&fromId={HttpUtility.UrlEncode(message.Recipient.Id)}&fromName={HttpUtility.UrlEncode(message.Recipient.Name)}&serviceUrl={HttpUtility.UrlEncode(message.ServiceUrl)}&conversationId={HttpUtility.UrlEncode(message.Conversation.Id)}&extraQueryParams={extraQueryParams}",
                                Type = "signin",
                                Title = "Connect"
                            };

                            cardButtons.Add(plButton);

                            SigninCard plCard = new SigninCard("Click connect to signin to Dynamics 365 (" + state.OrganizationUrl + ").", new List<CardAction>() { plButton });
                            Attachment plAttachment = plCard.ToAttachment();
                            replyToConversation.Attachments.Add(plAttachment);
                            await connector.Conversations.SendToConversationAsync(replyToConversation);
                        }
                    }
                    else
                    {
                        if (message.Attachments != null && message.Attachments.Count > 0)
                        {
                            List<CrmAttachment> attachments = new List<CrmAttachment>();
                            foreach (Attachment attach in message.Attachments)
                            {
                                if (!string.IsNullOrEmpty(attach.ContentUrl))
                                {
                                    attachments.Add(new CrmAttachment()
                                    {
                                        Attachment = new System.Net.WebClient().DownloadData(attach.ContentUrl),
                                        FileName = attach.Name
                                    });
                                }
                            }

                            Dialogs.CrmDialog dialog = new Dialogs.CrmDialog(message);
                            dialog.Attachments = attachments;

                            if (string.IsNullOrEmpty(message.Text))
                            {
                                await connector.Conversations.ReplyToActivityAsync(message.CreateReply($"I got your file. What would you like to do with it? You can say {string.Join(" or ", Dialogs.CrmDialog.AttachmentActionPhrases)}."));
                            }
                        }
                        else
                        {
                            await connector.Conversations.ReplyToActivityAsync(CreateTypingMessage(message));
                            await Conversation.SendAsync(message, () => new CRMBot.Dialogs.CrmDialog(message));
                        }
                    }
                }
                else
                {
                    Activity systemReply = HandleSystemMessage(message, connector);
                    if (systemReply != null)
                    {
                        await connector.Conversations.ReplyToActivityAsync(systemReply);
                    }
                }
            }
            catch (Exception ex)
            {
                await connector.Conversations.ReplyToActivityAsync(message.CreateReply($"Kabloooey! Well played human you just fried my circuits. Thanks for being patient, I'm still learning to do some things while in preview. Hopefully, I'll get this worked out soon. Here's your prize: {ex.Message}"));
            }
            return new HttpResponseMessage(System.Net.HttpStatusCode.Accepted);
        }

        public static Activity CreateTypingMessage(Activity message)
        {
            Activity reply = message.CreateReply();
            reply.Type = ActivityTypes.Typing;
            reply.Text = null;
            return reply;
        }
        public static Activity HandleSystemMessage(Activity message, ConnectorClient connector)
        {
            if (message.Type == ActivityTypes.Ping)
            {
                /*
                QueryExpression query = new QueryExpression("systemuser");
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 1;
                query.PageInfo.PageNumber = 1;
                query.ColumnSet = new ColumnSet(new string[] { "systemuserid" });
                using (OrganizationWebProxyClient serviceProxy = CrmHelper.CreateOrganizationService(Guid.Empty.ToString()))
                {
                    EntityCollection collection = serviceProxy.RetrieveMultiple(query);
                    Activity ping = message.CreateReply();
                    ping.Type = ActivityTypes.Ping;
                    ping.Text = null;
                    return ping;
                }
                */
            }
            return null;
        }
    }
}