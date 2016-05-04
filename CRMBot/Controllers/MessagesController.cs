using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Utilities;

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

namespace CRMBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {

        private string[] RejectionStrings = new string[]
        {
            "I'm sorry I can't do that. You must CLOSE ALL DEALS!",
            "Nope",
            "Nie",
            "Not gonna happen",
            "Uh-uh",
            "Sorry",
            "We're done here",
            "Stop that!"
        };

        private static int rejectIndex = -1;
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public Message Post([FromBody]Message message)
        {
            if (message.Type == "Message")
            {
                if (string.IsNullOrEmpty(message.Text) && message.Attachments != null && message.Attachments.Count > 0)
                {
                    return message.CreateReplyMessage("That's pretty cool, but why are you sending me pics? I'm a chatbot.");
                }
                else if (message.Text.ToLower().Contains("say"))
                {
                    return message.CreateReplyMessage(message.Text.Substring(message.Text.ToLower().IndexOf("say") + 4));
                }
                else if (message.Text.ToLower().Contains("thank"))
                {
                    return message.CreateReplyMessage("You're welcome!");
                }
                else if (message.Text.ToLower().Contains("say goodbye"))
                {
                    ChatState.RetrieveChatState(message.ConversationId).SelectedEntity = null;
                    return message.CreateReplyMessage("I'll Be Back...");
                }
                else
                {
                    var client = new RestClient("https://api.projectoxford.ai");
                    var request = new RestRequest("/luis/v1/application?id=cc421661-4803-4359-b19b-35a8bae3b466&subscription-key=70c9f99320804782866c3eba387d54bf&q=" + message.Text, Method.GET);
                    // automatically deserialize result
                    IRestResponse<CRMBot.LuisResults.Result> response = client.Execute<CRMBot.LuisResults.Result>(request);

                    string bestIntention = string.Empty;
                    double max = 0.00;
                    if (response.Data == null)
                    {
                        return message.CreateReplyMessage($"Damn you demo Gods!!! I can't seem to connect to the internet. Please check your connection.");
                    }
                    foreach (var intent in response.Data.intents)
                    {
                        if (max < intent.score)
                        {
                            bestIntention = intent.intent;
                            max = intent.score;
                        }
                    }

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
                    else if (bestIntention == "HowMany")
                    {
                        if (response.Data != null && response.Data.entities != null && response.Data.entities.Count > 0 && response.Data.entities[0] != null)
                        {
                            string parentEntity = string.Empty;

                            string entityType = response.Data.entities[0].entity;
                            //Hack fix to use plural name from metadata
                            if (entityType.EndsWith("s"))
                            {
                                entityType = entityType.Substring(0, entityType.Length - 1);
                            }
                            QueryExpression expression = new QueryExpression(entityType);
                            if (ChatState.RetrieveChatState(message.ConversationId).SelectedEntity != null)
                            {
                                expression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.Id);
                            }

                            using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService())
                            {
                                EntityCollection collection = serviceProxy.RetrieveMultiple(expression);
                                if (collection.Entities != null)
                                {
                                    if (ChatState.RetrieveChatState(message.ConversationId).SelectedEntity != null)
                                    {
                                        return message.CreateReplyMessage($"I found {collection.Entities.Count} {entityType} for the {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.LogicalName} {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity["firstname"]} {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity["lastname"]}");
                                    }
                                    else
                                    {
                                        return message.CreateReplyMessage($"I found {collection.Entities.Count} {entityType} in CRM.");
                                    }
                                }
                            }
                        }
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
                            if (response.Data != null && response.Data.entities != null && response.Data.entities.Count > 0 && response.Data.entities[0].resolution != null)
                            {
                                LuisResults.Entity dateEntity = response.Data.entities.FirstOrDefault(e => e.type == "builtin.datetime.date");
                                if (dateEntity != null)
                                {
                                    date = DateTime.Parse(dateEntity.resolution.date);
                                    entity["scheduledend"] = date;
                                }
                            }

                            try
                            {
                                using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService())
                                {
                                    serviceProxy.Create(entity);
                                }
                            }
                            catch (FaultException<OrganizationServiceFault> ex)
                            {
                                return message.CreateReplyMessage(ex.Message);
                                throw;
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

                        string entityType = string.Empty;
                        Dictionary<string, string> atts = new Dictionary<string, string>();
                        Dictionary<string, double> attScores = new Dictionary<string, double>();
                        foreach (var entity in response.Data.entities)
                        {
                            if (entity.type == "EntityType")
                            {
                                entityType = entity.entity;
                            }
                            else
                            {
                                string[] split = entity.type.Split(':');
                                if (!atts.ContainsKey(split[split.Length - 1]))
                                {
                                    attScores.Add(split[split.Length - 1], entity.score);
                                    atts.Add(split[split.Length - 1], entity.entity);
                                }
                                else if(entity.score > attScores[split[split.Length - 1]])
                                {
                                    atts[split[split.Length - 1]] = entity.entity;
                                }
                            }
                        }

                        try
                        {
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
                                if (collection.Entities != null && collection.Entities.Count == 1)
                                {
                                    ChatState.RetrieveChatState(message.ConversationId).SelectedEntity = collection.Entities[0];
                                    output = $"I found a {entityType} named {collection.Entities[0]["firstname"]} {collection.Entities[0]["lastname"]} from {collection.Entities[0]["address1_city"]} what would you like to do next? ";
                                }
                                else
                                {
                                    output = $"Hmmm...I couldn't find that {entityType}.";
                                }
                            }
                        }
                        catch (FaultException<OrganizationServiceFault> ex)
                        {
                            return message.CreateReplyMessage(ex.Message);
                            throw;
                        }
                    }
                    //entityType = response.Data.entities
                    // return our reply to the user
                    return message.CreateReplyMessage(output);
                }
            }
            return message.CreateReplyMessage("Sorry I didn't understand that.");
        }
    }
}