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
        public async Task<Message> Post([FromBody]Message message)
        {
            try
            {
                if (message.Type == "Message")
                {
                    if (message.Text.ToLower().StartsWith("portalreg"))
                    {
                        string[] split = message.Text.Split('|');
                        if (split.Length > 1)
                        {
                            //MODEBUG TODO
                            QueryExpression query = new QueryExpression("cobalt_crmorganization");
                            query.Criteria.AddCondition("cobalt_registrationcode", ConditionOperator.Equal, message.Text);
                            using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(split[1]))
                            {
                                EntityCollection collection = serviceProxy.RetrieveMultiple(query);
                                if (collection.Entities != null && collection.Entities.Count == 1)
                                {
                                    return message.CreateReplyMessage("You're all set. Let us know if you have any questions.");
                                }
                            }
                        }
                        return message.CreateReplyMessage("Huh?");
                    }
                    else if (message.Text.ToLower().StartsWith("crmbot-"))
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
                                    if (message.From.ChannelId.ToLower() == "facebook" || message.From.ChannelId.ToLower() == "skype")
                                    {
                                        if (message.From.ChannelId.ToLower() == "facebook")
                                        {
                                            collection.Entities[0]["cobalt_facebookmessengerid"] = message.From.Id;
                                            collection.Entities[0]["cobalt_firstconversationid"] = message.ConversationId;
                                        }
                                        else if (message.From.ChannelId.ToLower() == "skype")
                                        {
                                            collection.Entities[0]["cobalt_skypeid"] = message.From.Id;
                                            collection.Entities[0]["cobalt_firstconversationid"] = message.ConversationId;
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
                                        return message.CreateReplyMessage("Hmmm... Unfortunately, I can't use this app to communicate right now. You can try sending the registraiton code using either Facebook Messenger or Skype.");
                                    }
                                    return message.CreateReplyMessage("Got it! Your registration has been confirmed. Now go back to the registration portal to complete your setup.");
                                }
                            }
                        }
                        return message.CreateReplyMessage("Hmmm... I don't recognize that registration code. Make sure you sent the correct code and try again or try registering again at cobalt.net/BotRegistration");
                    }

                    if (!ChatState.SetChatState(message))
                    {
                        return message.CreateReplyMessage($"Hey there. I'm CRM Bot. I don't believe we've met. Before we can work together you'll need to go [here](http://www.cobalt.net/botregistration) to tell me more about yourself and your CRM organization.");
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
                            Dialogs.CrmDialog dialog = new Dialogs.CrmDialog(message.ConversationId);
                            dialog.Attachments = attachments;

                            if (string.IsNullOrEmpty(message.Text))
                            {
                                return message.CreateReplyMessage($"I got your file. What would you like to do with it? You can say {string.Join(" or ", Dialogs.CrmDialog.AttachmentActionPhrases)}.");
                            }
                        }
                        return await Conversation.SendAsync(message, () => new Dialogs.CrmDialog(message.ConversationId));
                    }
                }
                else
                {
                    return HandleSystemMessage(message);
                }
            }
            catch (Exception ex)
            {
                return message.CreateReplyMessage($"Kabloooey! Nice work you just fried my circuits. Well played human. Here's your prize: {ex.ToString()}");
            }
        }

        private static Message HandleSystemMessage(Message message)
        {
            if (message.Type == "BotAddedToConversation")
            {
                return message.CreateReplyMessage($"Hello {message.From?.Name}! To get started say something like {string.Join(" or ", Dialogs.CrmDialog.WelcomePhrases)}.");
            }
            else if (message.Type == "BotRemovedFromConversation")
            {
                return message.CreateReplyMessage($"See ya {message.From?.Name}!");
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
                        ChatState.RetrieveChatState(message.ConversationId).Attachments = null;
                        if (ChatState.RetrieveChatState(message.ConversationId).SelectedEntity != null)
                        {
                            EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.LogicalName);
                            string primaryAtt = ChatState.RetrieveChatState(message.ConversationId).SelectedEntity[metadata.PrimaryNameAttribute].ToString();
                            ChatState.RetrieveChatState(message.ConversationId).SelectedEntity = null;
                            return message.CreateReplyMessage($"Okay. We're done with {primaryAtt}");
                        }
                        return message.CreateReplyMessage($"Okay. We're done with that");
                    }
                    else if (message.Text.ToLower().Contains("say goodbye"))
                    {
                        ChatState.RetrieveChatState(message.ConversationId).SelectedEntity = null;
                        ChatState.RetrieveChatState(message.ConversationId).Attachments = null;
                        return message.CreateReplyMessage("I'll Be Back...");
                    }
                    else if (message.Text.ToLower().Contains("thank"))
                    {
                        return message.CreateReplyMessage("You're welcome!");
                    }
                    else if (message.Text.ToLower().Contains("say"))
                    {
                        return message.CreateReplyMessage(message.Text.Substring(message.Text.ToLower().IndexOf("say") + 4));
                    }
                    else
                    {
                        CRMBot.LuisResults.Result result = CRMBot.LuisResults.Result.Parse(message.Text);

                        if (result == null)
                        {
                            return message.CreateReplyMessage($"Hmmm...I can't seem to connect to the internet. Please check your connection.");
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
                            return message.CreateReplyMessage(RejectionStrings[rejectIndex]);

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
                            if (ChatState.RetrieveChatState(message.ConversationId).SelectedEntity != null)
                            {
                                Entity entity = new Entity("task");
                                EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.LogicalName);
                                entity["subject"] = $"Follow up with {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity[metadata.PrimaryNameAttribute]}";
                                entity["regardingobjectid"] = new EntityReference(ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.LogicalName, ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.Id);

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
                                    return message.CreateReplyMessage($"Okay...I've created task to follow up with {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity["firstname"]} {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity["lastname"]} on { date.ToLongDateString() }");
                                }
                                else
                                {
                                    return message.CreateReplyMessage($"Okay...I've created task to follow up with {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity["firstname"]} {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity["lastname"]}");
                                }
                            }
                            else
                            {
                                return message.CreateReplyMessage($"Hmmm...I'm not sure who to follow up with. Say for example 'Locate contact John Smith'");
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
                                            ChatState.RetrieveChatState(message.ConversationId).SelectedEntity = collection.Entities[0];

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
                            return message.CreateReplyMessage(output);
                        }
                    }
                }
                return message.CreateReplyMessage("Sorry, I didn't understand that. I'm still learning. Hopefully my human trainers will help me understand that request next time.");
            }
            catch (Exception ex)
            {
                return message.CreateReplyMessage($"Kabloooey! Nice work you just fried my circuits. Well played human. Here's your prize: {ex.Message}");
            }
        }*/
    }
}