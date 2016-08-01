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

namespace CRMBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        public async Task<HttpResponseMessage> Post([FromBody]Activity message)
        {
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

            /*
            try
            {
                if (message.Type == "Message")
                {
                    if (message.Text.ToLower() == "ping")
                    {
                        QueryExpression query = new QueryExpression("systemuser");
                        query.PageInfo = new PagingInfo();
                        query.PageInfo.Count = 1;
                        query.PageInfo.PageNumber = 1;
                        query.ColumnSet = new ColumnSet(new string[] { "systeuserid" });
                        try
                        {
                            using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(message.Conversation.Id))
                            {
                                EntityCollection collection = serviceProxy.RetrieveMultiple(query);
                                return message.CreateReply("true");
                            }
                        }
                        catch
                        {
                            return message.CreateReply("false");
                        }
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
                                    }
                                    else
                                    {
                                        return message.CreateReply("Hmmm... Unfortunately, I can't use this app to communicate right now. You can try sending the registraiton code using either Facebook Messenger or Skype.");
                                    }
                                    if (ChatState.SetChatState(message))
                                    {
                                        ChatState chatState = ChatState.RetrieveChatState(message.Conversation.Id);
                                        return message.CreateReply($"Welcome {chatState.UserFirstName}! Your registration has been confirmed. To get started say something like {Dialogs.CrmDialog.BuildCommandList(CRMBot.Dialogs.CrmDialog.WelcomePhrases)}.");
                                    }
                                    else
                                    {
                                        return message.CreateReply("Hmmm... I found your registration but I couldn't connect to your CRM Organization. Make sure you've entered all information or try registering again at cobalt.net/BotRegistration");
                                    }
                                }
                            }
                        }
                        return message.CreateReply("Hmmm... I don't recognize that registration code. Make sure you sent the correct code and try again or try registering again at cobalt.net/BotRegistration");
                    }

                    if (!ChatState.SetChatState(message))
                    {
                        if (message.Text.ToLower().Contains("help"))
                        {
                            return message.CreateReply($"Before we can work together you'll need to go [here](http://www.cobalt.net/botregistration) to connect me to your CRM organization.");
                        }
                        else if (message.Text.ToLower().Contains("goodbye"))
                        {
                            return message.CreateReply("CRM you later...");
                        }
                        else if (message.Text.ToLower().Contains("thank"))
                        {
                            return message.CreateReply($"You're welcome!");
                        }
                        else if (message.Text.ToLower().StartsWith("say"))
                        {
                            return message.CreateReply(message.Text.Substring(message.Text.ToLower().IndexOf("say") + 4));
                        }
                        else
                        {
                            return message.CreateReply($"Hey there, I don't believe we've met. Unfortunately, I can't talk to strangers. Before we can work together you'll need to go [here](http://www.cobalt.net/botregistration) to connect me to your CRM organization.");
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
                                return message.CreateReply($"I got your file. What would you like to do with it? You can say {string.Join(" or ", Dialogs.CrmDialog.AttachmentActionPhrases)}.");
                            }
                        }
                        //TODO
                        return message.CreateReply($"TODO Luis.");
                        //return await Conversation.SendAsync(message, () => new Dialogs.CrmDialog(message.Conversation.Id));
                    }
                }
                else
                {
                    return HandleSystemMessage(message);
                }
            }
            catch (Exception ex)
            {
                return message.CreateReply($"Kabloooey! Well played human you just fried my circuits. Thanks for being patient, I'm still learning to do some things while in preview. Hopefully, I'll get this worked out soon. Here's your prize: {ex.Message}");
            }
            */
        }

        private static Activity HandleSystemMessage(Activity message)
        {
            if (message.Type.ToLower() == "botaddedtoconversation")
            {
                return message.CreateReply($"Hello {message.From?.Name}! To get started say something like {Dialogs.CrmDialog.BuildCommandList(CRMBot.Dialogs.CrmDialog.WelcomePhrases)}.");
            }
            else if (message.Type.ToLower() == "botremovedfromconversation")
            {
                return message.CreateReply($"See ya {message.From?.Name}!");
            }

            return null;
        }
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        /*
        public Message Post([FromBody]Message message)
        {
            try
            {
                if (message.Type == "Message")
                {
                    if (message.Text.ToLower().Contains("forget"))
                    {
                        ChatState.RetrieveChatState(message.Conversation.Id).Attachments = null;
                        if (ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity != null)
                        {
                            EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity.LogicalName);
                            string primaryAtt = ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity[metadata.PrimaryNameAttribute].ToString();
                            ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity = null;
                            return message.CreateReply($"Okay. We're done with {primaryAtt}");
                        }
                        return message.CreateReply($"Okay. We're done with that");
                    }
                    else if (message.Text.ToLower().Contains("say goodbye"))
                    {
                        ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity = null;
                        ChatState.RetrieveChatState(message.Conversation.Id).Attachments = null;
                        return message.CreateReply("I'll Be Back...");
                    }
                    else if (message.Text.ToLower().Contains("thank"))
                    {
                        return message.CreateReply("You're welcome!");
                    }
                    else if (message.Text.ToLower().Contains("say"))
                    {
                        return message.CreateReply(message.Text.Substring(message.Text.ToLower().IndexOf("say") + 4));
                    }
                    else
                    {
                        CRMBot.LuisResults.Result result = CRMBot.LuisResults.Result.Parse(message.Text);

                        if (result == null)
                        {
                            return message.CreateReply($"Hmmm...I can't seem to connect to the internet. Please check your connection.");
                        }

                        string bestIntention = result.RetrieveIntention();

                        string output = string.Empty;
                        if (bestIntention == "RejectLead")
                        {
                            rejectIndex++;
                            if (rejectIndex >= RejectionStrings.Length)
                            {
                                rejectIndex = 0;
                            }
                            return message.CreateReply(RejectionStrings[rejectIndex]);

                        }
                        else if (bestIntention == "Send")
                        {
                            //Send email
                            //string emailAddress = result.entities.Where(e => e.type == "Email").Max(e => e.score);

                            string email = string.Empty;
                        }
                        else if (bestIntention == "HowMany")
                        {
                            return MessageHandler.HandleHowMany(message, result);
                        }
                        else if (bestIntention == "FollowUp")
                        {
                            if (ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity != null)
                            {
                                Entity entity = new Entity("task");
                                EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity.LogicalName);
                                entity["subject"] = $"Follow up with {ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity[metadata.PrimaryNameAttribute]}";
                                entity["regardingobjectid"] = new EntityReference(ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity.LogicalName, ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity.Id);

                                DateTime date = DateTime.MinValue;
                                CRMBot.LuisResults.Entity dateEntity = result.RetrieveEntity(CRMBot.LuisResults.EntityTypeNames.DateTime);

                                if (dateEntity != null)
                                {
                                    DateTime[] dates = dateEntity.ParseDateTimes();
                                    if (dates != null && dates.Length > 0)
                                    {
                                        entity["scheduledend"] = dates[0];
                                        date = dates[0];
                                    }
                                }

                                using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService())
                                {
                                    serviceProxy.Create(entity);
                                }
                                if (date != DateTime.MinValue)
                                {
                                    return message.CreateReply($"Okay...I've created task to follow up with {ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity["firstname"]} {ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity["lastname"]} on { date.ToLongDateString() }");
                                }
                                else
                                {
                                    return message.CreateReply($"Okay...I've created task to follow up with {ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity["firstname"]} {ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity["lastname"]}");
                                }
                            }
                            else
                            {
                                return message.CreateReply($"Hmmm...I'm not sure who to follow up with. Say for example 'Locate contact John Smith'");
                            }
                        }
                        else if (bestIntention == "Locate" || bestIntention == "Select")
                        {
                            LuisResults.Entity entityTypeEntity = result.RetrieveEntity(LuisResults.EntityTypeNames.EntityType);
                            if (entityTypeEntity != null)
                            {
                                LuisResults.Entity dateEntity = result.RetrieveEntity(LuisResults.EntityTypeNames.DateTime);
                                if (dateEntity != null)
                                {
                                    return MessageHandler.HandleHowMany(message, result);
                                }
                                else
                                {
                                    string entityType = entityTypeEntity.entity;
                                    EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(entityType);
                                    Dictionary<string, string> atts = new Dictionary<string, string>();
                                    Dictionary<string, double> attScores = new Dictionary<string, double>();
                                    foreach (var entity in result.entities)
                                    {
                                        if (entity.score > .5)
                                        {
                                            if (entity.type != "EntityType")
                                            {
                                                string[] split = entity.type.Split(':');
                                                if (!atts.ContainsKey(split[split.Length - 1]))
                                                {
                                                    attScores.Add(split[split.Length - 1], entity.score);
                                                    atts.Add(split[split.Length - 1], entity.entity);
                                                }
                                                else if (entity.score > attScores[split[split.Length - 1]])
                                                {
                                                    atts[split[split.Length - 1]] = entity.entity;
                                                }
                                            }
                                        }
                                    }

                                    using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService())
                                    {
                                        QueryExpression expression = new QueryExpression(entityType);
                                        expression.ColumnSet = new ColumnSet(true);
                                        Dictionary<string, string>.Enumerator iEnum = atts.GetEnumerator();
                                        while (iEnum.MoveNext())
                                        {
                                            expression.Criteria.AddCondition(iEnum.Current.Key.ToLower(), ConditionOperator.Equal, iEnum.Current.Value);
                                        }
                                        EntityCollection collection = serviceProxy.RetrieveMultiple(expression);
                                        string entityDisplayName = entityType;
                                        if (metadata.DisplayName != null && metadata.DisplayName.UserLocalizedLabel != null && !string.IsNullOrEmpty(metadata.DisplayName.UserLocalizedLabel.Label))
                                        {
                                            entityDisplayName = metadata.DisplayName.UserLocalizedLabel.Label;
                                        }
                                        if (collection.Entities != null && collection.Entities.Count == 1)
                                        {
                                            ChatState.RetrieveChatState(message.Conversation.Id).SelectedEntity = collection.Entities[0];

                                            output = $"I found a {entityDisplayName} named {collection.Entities[0][metadata.PrimaryNameAttribute]} what would you like to do next? You can say {string.Join(" or ", ActionPhrases)}";
                                        }
                                        else
                                        {
                                            output = $"Hmmm...I couldn't find that {entityDisplayName}.";
                                        }
                                    }
                                }
                            }
                        }
                        //entityType = result.entities
                        // return our reply to the user
                        if (!string.IsNullOrEmpty(output))
                        {
                            return message.CreateReply(output);
                        }
                    }
                }
                return message.CreateReply("Sorry, I didn't understand that. I'm still learning. Hopefully my human trainers will help me understand that request next time.");
            }
            catch (Exception ex)
            {
                return message.CreateReply($"Kabloooey! Nice work you just fried my circuits. Well played human. Here's your prize: {ex.Message}");
            }
        }*/
    }
}