using Microsoft.Bot.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CRMBot
{
    public class MessageHandler
    {
        public const int MAX_RECORDS_TO_SHOW = 10;
        public static Message HandleHowMany(Message message, LuisResults.Result result)
        {
            if (result != null && result.entities != null && result.entities.Count > 0 && result.entities[0] != null)
            {
                string parentEntity = string.Empty;

                CRMBot.LuisResults.Entity entity = result.RetrieveEntity(CRMBot.LuisResults.EntityTypeNames.EntityType);
                if (entity != null && !string.IsNullOrEmpty(entity.entity))
                {
                    string entityType = entity.entity;
                    EntityMetadata entityMetadata = CrmHelper.RetrieveEntityMetadata(entityType);
                    QueryExpression expression = new QueryExpression(entityType);
                    expression.ColumnSet = new ColumnSet(new string[] { entityMetadata.PrimaryNameAttribute, entityMetadata.PrimaryIdAttribute });
                    //TODO make this smarter based on relationship metadata
                    if (ChatState.RetrieveChatState(message.ConversationId).SelectedEntity != null && ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.LogicalName == "systemuser" && entityMetadata.Attributes.Any(a => a.LogicalName == "createdby"))
                    {
                        expression.Criteria.AddCondition("createdby", ConditionOperator.Equal, ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.Id);
                    }
                    else if (ChatState.RetrieveChatState(message.ConversationId).SelectedEntity != null && entityMetadata.Attributes.Any(a => a.LogicalName == "regardingobjectid"))
                    {
                        expression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.Id);
                    }
                    else if (ChatState.RetrieveChatState(message.ConversationId).SelectedEntity != null && entityMetadata.Attributes.Any(a => a.LogicalName == "objectid"))
                    {
                        expression.Criteria.AddCondition("objectid", ConditionOperator.Equal, ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.Id);
                    }
                    CRMBot.LuisResults.Entity dateEntity = result.RetrieveEntity(LuisResults.EntityTypeNames.DateTime);
                    string whenString = string.Empty;
                    if (dateEntity != null)
                    {
                        DateTime[] dates = dateEntity.ParseDateTimes();
                        if (dates != null && dates.Length > 0)
                        {
                            string action = "created";
                            CRMBot.LuisResults.Entity actionEntity = result.RetrieveEntity(LuisResults.EntityTypeNames.Action);
                            if (actionEntity != null)
                            {
                                action = actionEntity.entity;
                            }
                            //TODO make this better? Currently taking created / modified and turning into createdon / modifiedon
                            if (!string.IsNullOrEmpty(action))
                            {
                                string field = action.ToLower() + "on";
                                expression.Criteria.AddCondition(field, ConditionOperator.OnOrAfter, new object[] { dates[0] });
                                if (dates.Length > 1)
                                {
                                    expression.Criteria.AddCondition(field, ConditionOperator.OnOrBefore, new object[] { dates[1] });
                                    whenString = $"{action} between {dates[0].ToString("MM/dd/yyyy")} and {dates[1].ToString("MM/dd/yyyy")}";
                                }
                                else
                                {
                                    whenString = $"{action} after {dates[0].ToString("MM/dd/yyyy")}";
                                }
                            }
                        }
                    }
                    using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService())
                    {
                        EntityCollection collection = serviceProxy.RetrieveMultiple(expression);
                        if (collection.Entities != null)
                        {
                            string displayName = entityType;
                            if (collection.Entities.Count == 1 && entityMetadata.DisplayName != null && entityMetadata.DisplayName.UserLocalizedLabel != null && !string.IsNullOrEmpty(entityMetadata.DisplayName.UserLocalizedLabel.Label))
                            {
                                displayName = entityMetadata.DisplayName.UserLocalizedLabel.Label;
                            }
                            else if (collection.Entities.Count != 1 && entityMetadata.DisplayCollectionName != null && entityMetadata.DisplayCollectionName.UserLocalizedLabel != null && !string.IsNullOrEmpty(entityMetadata.DisplayCollectionName.UserLocalizedLabel.Label))
                            {
                                displayName = entityMetadata.DisplayCollectionName.UserLocalizedLabel.Label;
                            }

                            System.Text.StringBuilder sb = new System.Text.StringBuilder();
                            for (int i = 0; i < collection.Entities.Count && i < MAX_RECORDS_TO_SHOW; i++)
                            {
                                sb.Append($"{i + 1}) ");
                                sb.Append(collection.Entities[i][entityMetadata.PrimaryNameAttribute]);
                                sb.Append("\r\n");
                                if (i == MAX_RECORDS_TO_SHOW - 1)
                                {
                                    sb.Append("...");
                                }
                            }

                            if (ChatState.RetrieveChatState(message.ConversationId).SelectedEntity != null)
                            {
                                return message.CreateReplyMessage($"I found {collection.Entities.Count} {displayName} for the {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity.LogicalName} {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity["firstname"]} {ChatState.RetrieveChatState(message.ConversationId).SelectedEntity["lastname"]} {whenString}\r\n{sb.ToString()}");
                            }
                            else
                            {
                                return message.CreateReplyMessage($"I found {collection.Entities.Count} {displayName} {whenString} in CRM.\r\n{sb.ToString()}");
                            }
                        }
                    }
                }
            }
            return CreateDefaultMessage(message);
        }

        public static Message CreateDefaultMessage(Message message)
        {
            return message.CreateReplyMessage("Sorry, I didn't understand that. I'm still learning. Hopefully my human trainers will help me understand that request next time.");
        }
    }
}