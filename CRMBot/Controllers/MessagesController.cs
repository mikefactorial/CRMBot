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
using CRMBot.Forms;

namespace CRMBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {

        public async Task<HttpResponseMessage> Post([FromBody]Activity message)
        {
            /* Bot out of service message
            if (message.Type == ActivityTypes.Message)
            {
                ConnectorClient connector = new ConnectorClient(new Uri(message.ServiceUrl));
                // return our reply to the user
                Activity reply = message.CreateReply("Hi there. CRM Bot is undergoing an upgrade of it's framework version to better serve your CRM needs. Check back soon. CRM you later...");
                await connector.Conversations.ReplyToActivityAsync(reply);
            }
            else
            {
                //HandleSystemMessage(activity);
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            return response;
            */
            MicrosoftAppCredentials.TrustServiceUrl(message.ServiceUrl, DateTime.MaxValue);
            ConnectorClient connector = new ConnectorClient(new Uri(message.ServiceUrl));
            try
            {
                if (message.Type == ActivityTypes.Message)
                {
                    Activity reply = message.CreateReply();
                    if (reply != null)
                    {
                        reply.Type = ActivityTypes.Typing;
                        reply.Text = null;
                        await connector.Conversations.ReplyToActivityAsync(reply);
                    }

                    if ((message.Text.ToLower().StartsWith("hi ") || message.Text.ToLower().StartsWith("hello ")) && (message.Text.ToLower().Contains("emaline") || message.Text.ToLower().Contains("emmy")))
                    {
                        await Conversation.SendAsync(message, EmmyForm.MakeRootDialog);
                    }
                    else if (message.Text.ToLower().StartsWith("bot1") || message.Text.ToLower().StartsWith("bot0"))
                    {
                        QueryExpression query = new QueryExpression("cobalt_crmorganization");
                        query.Criteria.AddCondition("cobalt_registrationcode", ConditionOperator.Equal, message.Text);
                        using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(Guid.Empty.ToString()))
                        {
                            EntityCollection collection = serviceProxy.RetrieveMultiple(query);
                            if (collection.Entities != null && collection.Entities.Count == 1)
                            {
                                if (message.From != null)
                                {
                                    if (message.ChannelId.ToLower() == "facebook" || message.ChannelId.ToLower() == "skype")
                                    {
                                        if (message.ChannelId.ToLower() == "facebook")
                                        {
                                            collection.Entities[0]["cobalt_facebookmessengerid"] = message.From.Id;
                                            collection.Entities[0]["cobalt_firstconversationid"] = message.Conversation.Id;
                                        }
                                        else if (message.ChannelId.ToLower() == "skype")
                                        {
                                            collection.Entities[0]["cobalt_skypeid"] = message.From.Id;
                                            collection.Entities[0]["cobalt_firstconversationid"] = message.Conversation.Id;
                                        }
                                        serviceProxy.Update(collection.Entities[0]);

                                        SetStateRequest setState = new SetStateRequest();
                                        setState.EntityMoniker = new EntityReference();
                                        setState.EntityMoniker.Id = collection.Entities[0].Id;

                                        setState.EntityMoniker.LogicalName = collection.Entities[0].LogicalName;
                                        setState.State = new OptionSetValue(0);
                                        setState.Status = new OptionSetValue(533470000);
                                        serviceProxy.Execute(setState);
                                        if (ChatState.SetChatState(message))
                                        {
                                            ChatState chatState = ChatState.RetrieveChatState(message.Conversation.Id);
                                            await connector.Conversations.ReplyToActivityAsync(message.CreateReply($"Welcome {chatState.UserFirstName}! Thanks for registering. To get started say something like {Dialogs.CrmDialog.BuildCommandList(CRMBot.Dialogs.CrmDialog.WelcomePhrases)}."));
                                        }
                                        else
                                        {
                                            await connector.Conversations.ReplyToActivityAsync(message.CreateReply("Hmmm... I found your registration but I couldn't connect to your CRM Organization. Make sure you've entered all information or try registering again at [http://www.cobalt.net/botregistration](http://www.cobalt.net/botregistration)"));
                                        }
                                    }
                                    else
                                    {
                                        await connector.Conversations.ReplyToActivityAsync(message.CreateReply("Hmmm... Unfortunately, I can't use this app to communicate right now. You can try sending the registraiton code using either Facebook Messenger or Skype."));
                                    }
                                }
                                else
                                {
                                    await connector.Conversations.ReplyToActivityAsync(message.CreateReply("Hmmm... I don't recognize that registration code. Make sure you sent the correct code and try again or try registering again at [http://www.cobalt.net/botregistration](http://www.cobalt.net/botregistration)"));
                                }
                            }
                            else
                            {
                                await connector.Conversations.ReplyToActivityAsync(message.CreateReply("Hmmm... I don't recognize that registration code. Make sure you sent the correct code and try again or try registering again at [http://www.cobalt.net/botregistration](http://www.cobalt.net/botregistration)"));
                            }
                        }
                    }
                    else if (!ChatState.SetChatState(message))
                    {
                        if (message.Text.ToLower().Contains("help"))
                        {
                            await connector.Conversations.ReplyToActivityAsync(message.CreateReply($"Before we can work together you'll need to go [here](http://www.cobalt.net/botregistration) to connect me to your CRM organization."));
                        }
                        else if (message.Text.ToLower().Contains("see ya") || message.Text.ToLower().Contains("bye") || message.Text.ToLower().Contains("later"))
                        {
                            ChatState.ClearChatState(message.Conversation.Id);
                            await connector.Conversations.ReplyToActivityAsync(message.CreateReply("CRM you later..."));
                        }
                        else if (message.Text.ToLower().Contains("thank"))
                        {
                            await connector.Conversations.ReplyToActivityAsync(message.CreateReply($"You're welcome!"));
                        }
                        else if (message.Text.ToLower().StartsWith("say"))
                        {
                            await connector.Conversations.ReplyToActivityAsync(message.CreateReply(message.Text.Substring(message.Text.ToLower().IndexOf("say") + 4)));
                        }
                        else
                        {
                            //await Conversation.SendAsync(message, LeadForm.MakeRootDialog);
                            await connector.Conversations.ReplyToActivityAsync(message.CreateReply($"Hey there, I don't believe we've met. Unfortunately, I can't talk to strangers. Before we can work together you'll need to go [here](http://www.cobalt.net/botregistration) to connect me to your CRM organization."));
                        }
                    }
                    else
                    {
                        if (message.Attachments != null && message.Attachments.Count > 0)
                        {
                            List<byte[]> attachments = new List<byte[]>();
                            foreach (Attachment attach in message.Attachments)
                            {
                                if (!string.IsNullOrEmpty(attach.ContentUrl))
                                {
                                    attachments.Add(new System.Net.WebClient().DownloadData(attach.ContentUrl));
                                }
                            }

                            Dialogs.CrmDialog dialog = new Dialogs.CrmDialog(message.Conversation.Id);
                            dialog.Attachments = attachments;

                            if (string.IsNullOrEmpty(message.Text))
                            {
                                await connector.Conversations.ReplyToActivityAsync(message.CreateReply($"I got your file. What would you like to do with it? You can say {string.Join(" or ", Dialogs.CrmDialog.AttachmentActionPhrases)}."));
                            }
                        }
                        else
                        {
                            await Conversation.SendAsync(message, () => new Dialogs.CrmDialog(message.Conversation.Id));
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
                using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(Guid.Empty.ToString()))
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

        public static Activity DoSetup(Activity message, ConnectorClient connector)
        {
            ChatState chatState = ChatState.RetrieveChatState(message.Conversation.Id);
            if (chatState.Lead == null)
            {
                chatState.Lead = new Microsoft.Xrm.Sdk.Entity("lead");
            }
            if (chatState.CrmOrganization == null)
            {
                chatState.CrmOrganization = new Microsoft.Xrm.Sdk.Entity("cobalt_crmorganization");
            }
            if (string.IsNullOrEmpty(chatState.UserFirstName) && chatState.SetupStage == AccountSetupStage.None)
            {
                chatState.SetupStage = chatState.SetupStage;
                return message.CreateReply($"Hey there, I don't believe we've met. Unfortunately, I can't talk to strangers. Before we can work together you'll need to connect me to your CRM Organization. What would you like me to call you?");
            }
            else if ((chatState.Lead["firstname"] == null || string.IsNullOrEmpty(chatState.Lead["firstname"].ToString())) && chatState.SetupStage == AccountSetupStage.FirstName)
            {
                chatState.SetupStage = AccountSetupStage.OrgUrl;
                chatState.Lead["firstname"] = message.Text;
                chatState.UserFirstName = message.Text;
                return message.CreateReply($"Thanks {message.Text}! Now enter your CRM orgaization URL (e.g. https://org.crm.dynamics.com)");
            }
            else if ((chatState.CrmOrganization["cobalt_organizationurl"] == null || string.IsNullOrEmpty(chatState.Lead["cobalt_organizationurl"].ToString())) && chatState.SetupStage == AccountSetupStage.OrgUrl)
            {
                //TODO MODEBUG Login dialog
                chatState.SetupStage = AccountSetupStage.Credentials;
                chatState.Lead["cobalt_organizationurl"] = message.Text;
                return message.CreateReply($"Got it. Whenever you send me a message I will connect to {message.Text}! All that's left is to sign in to your CRM orgaization.");
            }
            else if ((chatState.CrmOrganization["cobalt_username"] == null || string.IsNullOrEmpty(chatState.Lead["cobalt_username"].ToString())) && chatState.SetupStage == AccountSetupStage.Credentials)
            {
                chatState.SetupStage = AccountSetupStage.Complete;
                chatState.Lead["cobalt_organizationurl"] = message.Text;
                return message.CreateReply($"Got it. Whenever you send me a message I will connect to {message.Text}! All that's left is to sign in to your CRM orgaization.");
            }
            return null;
        }
    }
}