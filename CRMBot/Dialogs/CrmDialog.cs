using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace CRMBot.Dialogs
{
    [LuisModel("cc421661-4803-4359-b19b-35a8bae3b466", "70c9f99320804782866c3eba387d54bf")]
    [Serializable]
    public class CrmDialog : LuisDialog<object>
    {
        public static string[] AttachmentActionPhrases = new string[]
        {
            "\"Attach as subject 'Powerpoint Presentation'\"",
            "\"Attach as subject 'Profile Pic'\""
        };
        public static string[] WelcomePhrases = new string[]
        {
            "\"Find contact, lead etc. Susan Connor.\"",
            "\"How many opportunities, leads etc. have been created today?\"",
        };
        public static string[] ActionPhrases = new string[]
        {
            "\"How many tasks, emails etc.\"",
            "\"Follow up July 12 2016\"",
            "\"Follow up next Tuesday\"",
            "Send me an image and say \"Attach as 'Profile Pic'\""
        };

        public static string[] FindActionPhrases = new string[]
        {
            "Find contact Dave Grohl.\"",
            "Find lead Susan Connor.\"",
            "Find opportunity Roger Waters.\""
        };

        public static string defaultMessage = $"Sorry, I didn't understand that. {FormatPhrases(WelcomePhrases)}";

        public const int MAX_RECORDS_TO_SHOW = 10;

        private string conversationId = string.Empty;

        public CrmDialog(string conversationId)
        {
            this.conversationId = conversationId;
        }

        [LuisIntent("FollowUp")]
        public async Task FollowUp(IDialogContext context, LuisResult result)
        {
            if (this.SelectedEntity != null)
            {
                Entity entity = new Entity("task");
                entity["subject"] = $"Follow up with {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]}";
                entity["regardingobjectid"] = new EntityReference(this.SelectedEntity.LogicalName, this.SelectedEntity.Id);

                DateTime date = DateTime.MinValue;
                EntityRecommendation dateEntity = result.RetrieveEntity(EntityTypeNames.DateTime);

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
                    await context.PostAsync($"Okay...I've created task to follow up with {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]} on { date.ToLongDateString() }");
                    context.Wait(MessageReceived);
                }
                else
                {
                    await context.PostAsync($"Okay...I've created task to follow up with {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]}");
                    context.Wait(MessageReceived);
                }
            }
            else
            {
                await context.PostAsync($"Hmmm...I'm not sure who to follow up with. Say for example {FormatPhrases(FindActionPhrases)}");
                context.Wait(MessageReceived);
            }

        }
        [LuisIntent("Display")]
        public async Task Display(IDialogContext context, LuisResult result)
        {
            await context.PostAsync(defaultMessage);
            context.Wait(MessageReceived);
        }
        [LuisIntent("Send")]
        public async Task Send(IDialogContext context, LuisResult result)
        {
            await context.PostAsync(defaultMessage);
            context.Wait(MessageReceived);
        }
        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            if (result.Query.ToLower().Contains("forget"))
            {
                this.Attachments = null;
                if (this.SelectedEntity != null && this.SelectedEntityMetadata != null)
                {
                    string primaryAtt = this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute].ToString();
                    this.SelectedEntity = null;
                    await context.PostAsync($"Okay. We're done with {primaryAtt}");
                }
                await context.PostAsync($"Okay. We're done with that");
                context.Wait(MessageReceived);
            }
            else if (result.Query.ToLower().Contains("say goodbye"))
            {
                this.Attachments = null;
                this.SelectedEntity = null;
                await context.PostAsync("I'll Be Back...");
                context.Wait(MessageReceived);
            }
        }
        [LuisIntent("Open")]
        public async Task Open(IDialogContext context, LuisResult result)
        {
            await context.PostAsync(defaultMessage);
            context.Wait(MessageReceived);
        }
        [LuisIntent("Update")]
        public async Task Update(IDialogContext context, LuisResult result)
        {
            await context.PostAsync(defaultMessage);
            context.Wait(MessageReceived);
        }
        [LuisIntent("Create")]
        public async Task Create(IDialogContext context, LuisResult result)
        {
            await context.PostAsync(defaultMessage);
            context.Wait(MessageReceived);
        }
        [LuisIntent("Locate")]
        public async Task Locate(IDialogContext context, LuisResult result)
        {
            EntityRecommendation entityTypeEntity = result.RetrieveEntity(EntityTypeNames.EntityType);
            if (entityTypeEntity != null)
            {
                EntityRecommendation dateEntity = result.RetrieveEntity(EntityTypeNames.DateTime);
                if (dateEntity != null)
                {
                   await CountRecords(context, result);
                }
                else
                {
                    string entityType = entityTypeEntity.Entity;
                    EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(entityType);
                    Dictionary<string, string> atts = new Dictionary<string, string>();
                    Dictionary<string, double?> attScores = new Dictionary<string, double?>();
                    foreach (var entity in result.Entities)
                    {
                        if (entity.Score > .5)
                        {
                            if (entity.Type != "EntityType")
                            {
                                string[] split = entity.Type.Split(':');
                                if (!atts.ContainsKey(split[split.Length - 1]))
                                {
                                    attScores.Add(split[split.Length - 1], entity.Score);
                                    atts.Add(split[split.Length - 1], entity.Entity);
                                }
                                else if (entity.Score > attScores[split[split.Length - 1]])
                                {
                                    atts[split[split.Length - 1]] = entity.Entity;
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
                            this.SelectedEntity = collection.Entities[0];
                            await context.PostAsync($"I found a {entityDisplayName} named {collection.Entities[0][metadata.PrimaryNameAttribute]} what would you like to do next? You can say {string.Join(" or ", ActionPhrases)}");
                        }
                        else
                        {
                            await context.PostAsync($"Hmmm...I couldn't find that {entityDisplayName}.");
                        }
                    }
                }
            }
            context.Wait(MessageReceived);
        }

        [LuisIntent("HowMany")]
        public async Task CountRecords(IDialogContext context, LuisResult result)
        {
            string parentEntity = string.Empty;

            EntityRecommendation entityType = result.RetrieveEntity(EntityTypeNames.EntityType);
            if (entityType != null && !string.IsNullOrEmpty(entityType.Entity))
            {
                EntityMetadata entityMetadata = CrmHelper.RetrieveEntityMetadata(entityType.Entity);
                QueryExpression expression = new QueryExpression(entityType.Entity);
                expression.ColumnSet = new ColumnSet(new string[] { entityMetadata.PrimaryNameAttribute, entityMetadata.PrimaryIdAttribute });
                //TODO make this smarter based on relationship metadata
                if (this.SelectedEntity != null && this.SelectedEntity.LogicalName == "systemuser" && entityMetadata.Attributes.Any(a => a.LogicalName == "createdby"))
                {
                    expression.Criteria.AddCondition("createdby", ConditionOperator.Equal, this.SelectedEntity.Id);
                }
                else if (this.SelectedEntity != null && entityMetadata.Attributes.Any(a => a.LogicalName == "regardingobjectid"))
                {
                    expression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, this.SelectedEntity.Id);
                }
                else if (this.SelectedEntity != null && entityMetadata.Attributes.Any(a => a.LogicalName == "objectid"))
                {
                    expression.Criteria.AddCondition("objectid", ConditionOperator.Equal, this.SelectedEntity.Id);
                }

                string whenString = string.Empty;

                EntityRecommendation dateEntity = result.RetrieveEntity(EntityTypeNames.DateTime);
                if (dateEntity != null)
                {
                    DateTime[] dates = dateEntity.ParseDateTimes();
                    if (dates != null && dates.Length > 0)
                    {
                        string action = "created";
                        EntityRecommendation actionEntity = result.RetrieveEntity(EntityTypeNames.Action);
                        if (actionEntity != null)
                        {
                            action = actionEntity.Entity;
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
                        string displayName = entityType.Entity;
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

                        if (this.SelectedEntity != null)
                        {
                            await context.PostAsync($"I found {collection.Entities.Count} {displayName} for the {this.SelectedEntity.LogicalName} {this.SelectedEntity["firstname"]} {this.SelectedEntity["lastname"]} {whenString}\r\n{sb.ToString()}");
                            context.Wait(MessageReceived);
                        }
                        else
                        {
                            await context.PostAsync($"I found {collection.Entities.Count} {displayName} {whenString} in CRM.\r\n{sb.ToString()}");
                            context.Wait(MessageReceived);
                        }
                    }
                }
            }
            else
            {
                await context.PostAsync(defaultMessage);
                context.Wait(MessageReceived);
            }
        }
        [LuisIntent("Attach")]
        public async Task Attach(IDialogContext context, LuisResult result)
        {
            if (this.SelectedEntity == null)
            {
                await context.PostAsync($"There's nothing to attach the file to. Say {FormatPhrases(FindActionPhrases)} to find a record to attach this to.");
            }
            else if (this.Attachments == null || this.Attachments.Count == 0)
            {
                await context.PostAsync($"There's nothing to attach. Send me a file and Say {FormatPhrases(AttachmentActionPhrases)} to attach a file.");
            }
            else
            {
                EntityRecommendation subjectEntity = result.RetrieveEntity(EntityTypeNames.AttributeValue);
                string subject = "Attachment";
                if (subjectEntity != null)
                {
                    subject = subjectEntity.Entity;
                }
                int i = 1;
                foreach (byte[] attachment in this.Attachments)
                {
                    Entity annotation = new Entity("annotation");
                    annotation["objectid"] = new EntityReference() { Id = this.SelectedEntity.Id, LogicalName = this.SelectedEntity.LogicalName };
                    string encodedData = System.Convert.ToBase64String(attachment);

                    annotation["filename"] = subject + i;
                    annotation["subject"] = subject;
                    annotation["mimetype"] = "application /octet-stream";
                    annotation["documentbody"] = encodedData;

                    using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService())
                    {
                        serviceProxy.Create(annotation);
                        await context.PostAsync($"Okay. I've attached the file to {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]} as a note with the Subject '{annotation["subject"]}'");
                    }
                }
            }

            context.Wait(MessageReceived);
        }
        public static string FormatPhrases(string[] phrases)
        {
            return string.Join(" or ", phrases);
        }
        protected EntityMetadata SelectedEntityMetadata
        {
            get
            {
                return CrmHelper.RetrieveEntityMetadata(this.SelectedEntity.LogicalName);
            }
        }

        protected Entity SelectedEntity
        {
            get
            {
                return ChatState.RetrieveChatState(this.conversationId).Get(ChatState.SelectedEntity) as Entity;
            }
            set
            {
                ChatState.RetrieveChatState(this.conversationId).Set(ChatState.SelectedEntity, value);
            }
        }
        private List<byte[]> Attachments
        {
            get
            {
                return ChatState.RetrieveChatState(this.conversationId).Get(ChatState.Attachments) as List<byte[]>;
            }
            set
            {
                ChatState.RetrieveChatState(this.conversationId).Set(ChatState.Attachments, value);
            }
        }
    }
}