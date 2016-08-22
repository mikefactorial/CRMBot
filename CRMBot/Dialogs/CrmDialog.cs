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
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace CRMBot.Dialogs
{
    //[LuisModel("cc421661-4803-4359-b19b-35a8bae3b466", "70c9f99320804782866c3eba387d54bf")]
    [LuisModel("64c400cf-b36d-4874-bd01-1c7567e57d8a", "1d8d05db65364c67a7bf15a4bf855f03")]
    [Serializable]
    public class CrmDialog : LuisDialog<object>
    {
        public static string[] AttachmentActionPhrases = new string[]
        {
            "\"Attach as subject 'Powerpoint Presentation'\"",
            "\"Attach as Note\""
        };
        public static string[] HelpPhrases = new string[]
        {
            "\"Create lead Sarah Connor\"",
            "\"Update lead Sarah Connor set home phone to 234-234-2345\"",
            "\"How many leads were created this week\"",
            "\"Find lead Sarah Connor\"",
            "\"Open lead Sarah Connor\"",
            "\"Follow up with lead Sarah Connor next Tuesday\"",
            "For more help go [here](http://www.cobalt.net/crmbot)."
        };

        public static string[] WelcomePhrases = new string[]
        {
            "\"Create lead Sarah Connor\"",
            "\"Update lead Sarah Connor set home phone to 234-234-2345\"",
            "\"How many leads were created this week\"",
            "\"Find lead Sarah Connor\"",
            "\"Open lead Sarah Connor\"",
            "\"Follow up with lead Sarah Connor next Tuesday\"",
            "For more help go [here](http://www.cobalt.net/crmbot)."
        };
        public static string[] ActionPhrases = new string[]
        {
            "\"How many tasks, emails etc.\"",
            "\"Follow up July 12 2016\"",
            "\"Show me the name, status, date created etc.\"",
            "Send me an image and say \"Attach as 'Profile Pic'\"",
            "\"Forget current record\"",
        };

        public static string[] FindActionPhrases = new string[]
        {
            "\"Find contact John Connor.\"",
            "\"Find lead Sarah Connor.\"",
            "\"Find opportunity Kyle Reese.\""
        };

        public static string defaultMessage = $"Sorry, I didn't understand that. Try saying {CrmDialog.BuildCommandList(CrmDialog.WelcomePhrases)}";

        public static string waitMessage = "Got it...Give me a second while I ";
        public const int MAX_RECORDS_TO_SHOW_PER_PAGE = 5;

        private string conversationId = string.Empty;

        public CrmDialog(string conversationId)
        {
            this.conversationId = conversationId;
        }

        [LuisIntent("FollowUp")]
        public async Task FollowUp(IDialogContext context, LuisResult result)
        {
            this.FindEntity(result, true);
            if (this.SelectedEntity != null)
            {
                Entity entity = new Entity("task");
                string subject = $"Follow up with {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]}";
                entity["regardingobjectid"] = new EntityReference(this.SelectedEntity.LogicalName, this.SelectedEntity.Id);

                DateTime date = DateTime.MinValue;
                EntityRecommendation dateEntity = result.RetrieveEntity(this.conversationId, EntityTypeNames.DateTime);

                if (dateEntity != null)
                {
                    List<DateTime> dates = dateEntity.ParseDateTimes();
                    if (dates != null && dates.Count > 0)
                    {
                        entity["scheduledend"] = dates[0];
                        date = dates[0];
                    }
                }

                if (date != DateTime.MinValue)
                {
                    entity["subject"] = subject + " - " + date.ToShortDateString();
                }
                else
                {
                    entity["subject"] = subject;
                }
                using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(this.conversationId))
                {
                    serviceProxy.Create(entity);
                }
                if (date != DateTime.MinValue)
                {
                    await context.PostAsync($"Okay...I've created task to follow up with {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]} on { date.ToLongDateString() }");
                }
                else
                {
                    await context.PostAsync($"Okay...I've created task to follow up with {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]}");
                }
            }
            else
            {
                await context.PostAsync($"Hmmm...I'm not sure who to follow up with. Say for example {FormatPhrases(FindActionPhrases)}");
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Display")]
        public async Task Display(IDialogContext context, LuisResult result)
        {
            this.FindEntity(result, true);
            if (this.SelectedEntity != null)
            {
                EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, this.SelectedEntity.LogicalName);
                EntityRecommendation displayField = result.RetrieveEntity(this.conversationId, EntityTypeNames.DisplayField);
                string displayName = RetrieveEntityDisplayName(metadata, false);
                if (displayField != null)
                {
                    string att = CrmHelper.FindAttributeLogicalName(metadata, displayField.Entity);
                    string displayValue = GetAttributeDisplayValue(metadata, att);
                    if (!string.IsNullOrEmpty(displayValue))
                    {
                        await context.PostAsync($"{this.SelectedEntity[metadata.PrimaryNameAttribute]}'s {displayField.Entity} is {displayValue}");
                    }
                    else
                    {
                        await context.PostAsync($"The {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]} does not have a {displayField.Entity}");
                    }
                }
                else
                {
                    await context.PostAsync(defaultMessage);
                }
            }
            else if (this.FilteredEntities != null && this.FilteredEntities.Length > 0)
            {
                EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, this.FilteredEntities[0].LogicalName);
                string displayName = RetrieveEntityDisplayName(metadata, true);
                await context.PostAsync($"These are the currently selected {displayName} {this.BuildFilteredEntitiesList()}");
            }
            else
            {
                await context.PostAsync(defaultMessage);
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Send")]
        public async Task Send(IDialogContext context, LuisResult result)
        {
            //MODEBUG TODO
            await context.PostAsync(defaultMessage);
            context.Wait(MessageReceived);
        }
        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            ChatState chatState = ChatState.RetrieveChatState(conversationId);
            int selection = -1;

            EntityRecommendation ordinal = result.RetrieveEntity(this.conversationId, EntityTypeNames.Ordinal);
            if (ordinal != null || Int32.TryParse(result.Query, out selection))
            {
                bool gotIndex = true;
                if (ordinal != null)
                {
                    string number = string.Empty;
                    foreach (char c in ordinal.Entity)
                    {
                        if (c >= 48 || c <= 57)
                        {
                            number += c;
                        }
                    }
                    gotIndex = Int32.TryParse(number, out selection);
                }
                else
                {
                    gotIndex = Int32.TryParse(result.Query, out selection);
                }
                if (gotIndex && this.FilteredEntities != null && this.FilteredEntities.Length >= selection && selection >= 1)
                {
                    this.SelectedEntity = this.FilteredEntities[selection - 1];
                    if (this.SelectedEntity != null)
                    {
                        this.FilteredEntities = null;
                        EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, this.SelectedEntity.LogicalName);
                        string displayName = RetrieveEntityDisplayName(metadata, false);
                        await context.PostAsync($"Got it. You've selected the {displayName} named {this.SelectedEntity[metadata.PrimaryNameAttribute]}. Now say {BuildCommandList(ActionPhrases)}");
                    }
                    else
                    {
                        await context.PostAsync($"Hmmm. I couldn't select that record. Make sure the number is within the current range of selected records (i.e. 1 to {this.FilteredEntities.Length}).");
                    }
                }
                else if (this.FilteredEntities == null || this.FilteredEntities.Length == 0)
                {
                    await context.PostAsync($"Hmmm. It looks like you're trying to select a record. However, there are no filtered records to choose from. Say something like \"How many leads were created last week.\" to select a list of records.");
                }
                else
                {
                    await context.PostAsync($"Hmmm. I couldn't select that record. Make sure the number is within the current range of selected records (i.e. 1 to {this.FilteredEntities.Length}).");
                }
            }
            else if (result.Query.ToLower().StartsWith("next"))
            {
                this.CurrentPageIndex = this.CurrentPageIndex + 1;
                string filteredEntitiesList = this.BuildFilteredEntitiesList();
                await context.PostAsync($"{filteredEntitiesList}");
            }
            else if (result.Query.ToLower().StartsWith("back"))
            {
                if (this.CurrentPageIndex >= 0)
                {
                    this.CurrentPageIndex = this.CurrentPageIndex - 1;
                }
                string filteredEntitiesList = this.BuildFilteredEntitiesList();
                await context.PostAsync($"{filteredEntitiesList}");
            }
            else if (result.Query.ToLower().StartsWith("help"))
            {
                if (this.SelectedEntity == null)
                {
                    await context.PostAsync($"Hi {chatState.UserFirstName}! Here are some commands you can try... {CrmDialog.BuildCommandList(CrmDialog.HelpPhrases)}.");
                }
                else
                {
                    EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, this.SelectedEntity.LogicalName);
                    string displayName = RetrieveEntityDisplayName(metadata, false);
                    await context.PostAsync($"You've selected a {displayName} named {this.SelectedEntity[metadata.PrimaryNameAttribute]}? Now say {BuildCommandList(ActionPhrases)}");
                }
            }
            else if (result.Query.ToLower().StartsWith("forget") || result.Query.ToLower().Contains("start over") || result.Query.ToLower().Contains("done"))
            {
                this.Attachments = null;
                if (this.SelectedEntity != null && this.SelectedEntityMetadata != null)
                {
                    string primaryAtt = this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute].ToString();
                    this.SelectedEntity = null;
                    await context.PostAsync($"Okay. We're done with {primaryAtt}.");
                }
                else
                {
                    await context.PostAsync($"Okay. We're done with that.");
                }
            }
            else if (result.Query.ToLower().StartsWith("goodbye") || result.Query.ToLower().StartsWith("bye"))
            {
                this.FilteredEntities = null;
                this.Attachments = null;
                this.SelectedEntity = null;
                await context.PostAsync($"CRM you later {chatState.UserFirstName}...");
            }
            else if (result.Query.ToLower().StartsWith("thank"))
            {
                await context.PostAsync($"You're welcome {chatState.UserFirstName}!");
            }
            else if (result.Query.ToLower().StartsWith("say"))
            {
                await context.PostAsync(result.Query.Substring(result.Query.ToLower().IndexOf("say") + 4));
            }
            else
            {
                if (result.Query.ToLower().Contains("hello") || result.Query.ToLower().Contains("hi"))
                {
                    await context.PostAsync($"Hey there {chatState.UserFirstName}!");
                }
                else if (result.Query.ToLower().Contains("what's up") || result.Query.ToLower().Contains("waddup") || result.Query.ToLower().Contains("sup") || result.Query.ToLower().Contains("whats up"))
                {
                    await context.PostAsync($"Nothing much, just serving up some CRM data like it's my job.");
                }
                await context.PostAsync($"I'm here to help {chatState.UserFirstName}! To get started say something like {CrmDialog.BuildCommandList(CrmDialog.WelcomePhrases)}.");
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Open")]
        public async Task Open(IDialogContext context, LuisResult result)
        {
            this.FindEntity(result, false);

            if (this.SelectedEntity != null)
            {
                ChatState chatState = ChatState.RetrieveChatState(conversationId);
                //[here](ms-dynamicsxrm://?pagetype=view&etn={this.SelectedEntity.LogicalName}&id={this.SelectedEntity.Id.ToString()}
                //Open in mobile client: ms-dynamicsxrm://?pagetype=view&etn=contact&id=899D4FCF-F4D3-E011-9D26-00155DBA3819
                //Open in browser: http://myorg.crm.dynamics.com/main.aspx?etn=account&pagetype=entityrecord&id=%7B91330924-802A-4B0D-A900-34FD9D790829%7D
                //await context.PostAsync($"Click [here]({chatState.OrganizationUrl}/main.aspx?etn={this.SelectedEntity.LogicalName}&pagetype=entityrecord&id={this.SelectedEntity.Id.ToString()}) to open the record in your browser or click [here](ms-dynamicsxrm://?pagetype=view&etn={this.SelectedEntity.LogicalName}&id={this.SelectedEntity.Id.ToString()}) to open the record in the mobile client.");
                await context.PostAsync($"Click [here]({chatState.OrganizationUrl}/main.aspx?etn={this.SelectedEntity.LogicalName}&pagetype=entityrecord&id={this.SelectedEntity.Id.ToString()}) to open the record in your browser.");
            }
            else
            {
                await context.PostAsync(defaultMessage);
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Update")]
        public async Task Update(IDialogContext context, LuisResult result)
        {
            this.FindEntity(result, true);
            if (this.SelectedEntity != null)
            {
                EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, this.SelectedEntity.LogicalName);
                EntityRecommendation attributeName = result.RetrieveEntity(this.conversationId, EntityTypeNames.AttributeName);
                EntityRecommendation attributeValue = result.RetrieveEntity(this.conversationId, EntityTypeNames.AttributeValue);
                string displayName = RetrieveEntityDisplayName(metadata, false);
                if (attributeName != null && attributeValue != null)
                {
                    string att = CrmHelper.FindAttributeLogicalName(metadata, attributeName.Entity);
                    if (!string.IsNullOrEmpty(att))
                    {
                        this.SelectedEntity[att] = attributeValue.Entity;
                        using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(this.conversationId))
                        {
                            serviceProxy.Update(this.SelectedEntity);
                            await context.PostAsync($"I've update the {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]} record with the new {attributeName.Entity} {attributeValue.Entity}");
                        }
                    }
                    else
                    {
                        await context.PostAsync($"I wasn't able to update the {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]} record. I didn't recognize the field name.");
                    }
                }
                else
                {
                    await context.PostAsync($"I wasn't able to update the {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]} record. I didn't recognize the field name and value.");
                }
            }
            else
            {
                await context.PostAsync(defaultMessage);
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Create")]
        public async Task Create(IDialogContext context, LuisResult result)
        {
            EntityRecommendation entityTypeEntity = result.RetrieveEntity(this.conversationId, EntityTypeNames.EntityType);
            if (entityTypeEntity != null && !string.IsNullOrEmpty(entityTypeEntity.Entity))
            {
                EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, entityTypeEntity.Entity);
                if (metadata != null)
                {
                    EntityRecommendation firstName = result.RetrieveEntity(this.conversationId, EntityTypeNames.FirstName);
                    EntityRecommendation lastName = result.RetrieveEntity(this.conversationId, EntityTypeNames.LastName);
                    EntityRecommendation attributeName = result.RetrieveEntity(this.conversationId, EntityTypeNames.AttributeName);
                    EntityRecommendation attributeValue = result.RetrieveEntity(this.conversationId, EntityTypeNames.AttributeValue);

                    Entity entity = new Entity(entityTypeEntity.Entity);
                    if (attributeValue != null)
                    {
                        if (attributeName != null)
                        {
                            string att = CrmHelper.FindAttributeLogicalName(metadata, attributeName.Entity);
                            SetValue(entity, metadata, att, attributeValue);
                        }
                        else
                        {
                            SetValue(entity, metadata, metadata.PrimaryNameAttribute, attributeValue);
                        }
                    }
                    else
                    {
                        if (firstName != null)
                        {
                            SetValue(entity, metadata, "firstname", firstName);
                        }
                        if (lastName != null)
                        {
                            SetValue(entity, metadata, "lastname", lastName);
                        }
                    }

                    if (metadata.IsActivity.HasValue && metadata.IsActivity.Value && this.SelectedEntity != null)
                    {
                        if (metadata.Attributes.Any(a => a.LogicalName == "customerid") && (this.SelectedEntity.LogicalName == "account" || this.SelectedEntity.LogicalName == "contact"))
                        {
                            entity["customerid"] = new EntityReference(this.SelectedEntity.LogicalName, this.SelectedEntity.Id);
                        }
                    }
                    using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(this.conversationId))
                    {
                        Guid id = serviceProxy.Create(entity);
                        this.SelectedEntity = serviceProxy.Retrieve(entity.LogicalName, id, new ColumnSet(true));
                    }

                    string displayName = RetrieveEntityDisplayName(metadata, false);
                    await context.PostAsync($"I've created a new {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]}. If you want to set additional fields say something like 'Update {displayName} set home phone to 555-555-5555'");
                }
                else
                {
                    await context.PostAsync($"Hmmm. I couldn't find any records called {entityTypeEntity.Entity} in CRM.");
                }
            }
            else
            {
                await context.PostAsync(defaultMessage);
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Locate")]
        public async Task Locate(IDialogContext context, LuisResult result)
        {
            string entityDisplayName = this.FindEntity(result, false);

            if (this.SelectedEntity != null)
            {
                EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, this.SelectedEntity.LogicalName);
                await context.PostAsync($"I found a {entityDisplayName} named {this.SelectedEntity[metadata.PrimaryNameAttribute]} what would you like to do next? You can say {string.Join(" or ", ActionPhrases)}");
            }
            else if (this.FilteredEntities != null && this.FilteredEntities.Length > 0)
            {
                this.CurrentPageIndex = 0;
                string filteredEntitiesList = this.BuildFilteredEntitiesList();
                await context.PostAsync($"I found {this.FilteredEntities.Length} {entityDisplayName} that match. To select one say a number below.\r\n{filteredEntitiesList}");
            }
            else
            {
                await context.PostAsync($"Hmmm...I couldn't find that {entityDisplayName}.");
            }
            context.Wait(MessageReceived);
        }

        [LuisIntent("HowMany")]
        public async Task CountRecords(IDialogContext context, LuisResult result)
        {
            string parentEntity = string.Empty;

            EntityRecommendation entityType = result.RetrieveEntity(this.conversationId, EntityTypeNames.EntityType);
            if (entityType != null && !string.IsNullOrEmpty(entityType.Entity))
            {
                EntityMetadata entityMetadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, entityType.Entity);
                if (entityMetadata != null)
                {
                    bool associatedEntities = false;
                    QueryExpression expression = new QueryExpression(entityType.Entity);

                    /********************************************************************************************************
                    //Winner winner bot shirt
                    *********************************************************************************************************/
                    expression.AddOrder("createdon", OrderType.Ascending);
                    /********************************************************************************************************
                    //Winner winner bot shirt
                    *********************************************************************************************************/

                    expression.ColumnSet = new ColumnSet(true);
                    //TODO make this smarter based on relationship metadata
                    if (this.SelectedEntity != null && this.SelectedEntity.LogicalName == "systemuser" && entityMetadata.Attributes.Any(a => a.LogicalName == "createdby"))
                    {
                        associatedEntities = true;
                        expression.Criteria.AddCondition("createdby", ConditionOperator.Equal, this.SelectedEntity.Id);
                    }
                    else if (this.SelectedEntity != null && entityMetadata.Attributes.Any(a => a.LogicalName == "regardingobjectid"))
                    {
                        associatedEntities = true;
                        expression.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, this.SelectedEntity.Id);
                    }
                    else if (this.SelectedEntity != null && entityMetadata.Attributes.Any(a => a.LogicalName == "objectid"))
                    {
                        associatedEntities = true;
                        expression.Criteria.AddCondition("objectid", ConditionOperator.Equal, this.SelectedEntity.Id);
                    }

                    string whenString = string.Empty;

                    EntityRecommendation dateEntity = result.RetrieveEntity(this.conversationId, EntityTypeNames.DateTime);
                    if (dateEntity != null)
                    {
                        List<DateTime> dates = dateEntity.ParseDateTimes();
                        if (dates != null && dates.Count > 0)
                        {
                            string action = "created";
                            EntityRecommendation actionEntity = result.RetrieveEntity(this.conversationId, EntityTypeNames.Action);
                            if (actionEntity != null)
                            {
                                action = actionEntity.Entity;
                            }
                            //TODO make this better? Currently taking created / modified and turning into createdon / modifiedon
                            if (!string.IsNullOrEmpty(action))
                            {
                                string field = action.ToLower() + "on";
                                expression.Criteria.AddCondition(field, ConditionOperator.OnOrAfter, new object[] { dates[0] });
                                if (dates.Count > 1)
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
                    using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(this.conversationId))
                    {
                        EntityCollection collection = serviceProxy.RetrieveMultiple(expression);
                        if (collection.Entities != null)
                        {
                            this.FilteredEntities = collection.Entities.ToArray();
                            string displayName = RetrieveEntityDisplayName(entityMetadata, this.FilteredEntities.Length != 1);

                            string filteredEntitiesList = this.BuildFilteredEntitiesList();
                            if (associatedEntities)
                            {
                                await context.PostAsync($"I found {collection.Entities.Count} {displayName} for the {this.SelectedEntity.LogicalName} {this.SelectedEntity["firstname"]} {this.SelectedEntity["lastname"]} {whenString}\r\n{filteredEntitiesList}");
                            }
                            else
                            {
                                await context.PostAsync($"I found {collection.Entities.Count} {displayName} {whenString} in CRM.\r\n{filteredEntitiesList}");
                            }
                        }
                    }
                }
                else
                {
                    await context.PostAsync($"Hmmm. I couldn't find any records called {entityType.Entity} in CRM.");
                }
            }
            else
            {
                await context.PostAsync(defaultMessage);
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Attach")]
        public async Task Attach(IDialogContext context, LuisResult result)
        {
            this.FindEntity(result, true);
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
                EntityRecommendation subjectEntity = result.RetrieveEntity(this.conversationId, EntityTypeNames.AttributeValue);
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

                    using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(this.conversationId))
                    {
                        serviceProxy.Create(annotation);
                        await context.PostAsync($"Okay. I've attached the file to {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]} as a note with the Subject '{annotation["subject"]}'");
                    }
                }
                this.Attachments = null;
            }

            context.Wait(MessageReceived);
        }
        public static string FormatPhrases(string[] phrases)
        {
            return string.Join(" or ", phrases);
        }

        protected string BuildFilteredEntitiesList()
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            if (this.FilteredEntities != null)
            {
                EntityMetadata entityMetadata = null;
                List<string> columns = new List<string>();
                int start = 0;
                if (this.CurrentPageIndex > 0)
                {
                    start = MAX_RECORDS_TO_SHOW_PER_PAGE * this.CurrentPageIndex;
                }
                bool hasMore = false;
                for (int i = start; i < this.FilteredEntities.Length && i < start + MAX_RECORDS_TO_SHOW_PER_PAGE; i++)
                {
                    hasMore = this.FilteredEntities.Length > (i + 1);
                    if (entityMetadata == null)
                    {
                        entityMetadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, this.FilteredEntities[i].LogicalName);

                        using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(this.conversationId))
                        {
                            columns.Add(entityMetadata.PrimaryNameAttribute);
                            QueryExpression expression = new QueryExpression("savedquery");
                            expression.ColumnSet = new ColumnSet(new string[] { "savedqueryid", "layoutxml" });
                            expression.Criteria.AddCondition("isdefault", ConditionOperator.Equal, new object[] { true });
                            expression.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, new object[] { entityMetadata.LogicalName });
                            expression.AddOrder("querytype", OrderType.Ascending);
                            EntityCollection collection = serviceProxy.RetrieveMultiple(expression);
                            if (collection.Entities != null)
                            {
                                if (collection.Entities.Count > 0 && collection.Entities[0]["layoutxml"] != null)
                                {
                                    string layoutXml = collection.Entities[0]["layoutxml"].ToString();
                                    Grid grid = Grid.Deserialize(layoutXml);
                                    if (grid.Row != null && grid.Row.Cells != null)
                                    {
                                        foreach (Cell cell in grid.Row.Cells)
                                        {
                                            if (!columns.Contains(cell.Name))
                                            {
                                                columns.Add(cell.Name);
                                            }
                                            //Don't show more than 3 columns for brevity
                                            if (columns.Count > 2)
                                            {
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    sb.Append($"{i + 1}. ");
                    for (int j = 0; j < columns.Count; j++)
                    {
                        string column = columns[j];
                        string displayValue = this.GetAttributeDisplayValue(this.FilteredEntities[i], entityMetadata, column);
                        if (j == 1)
                        {
                            sb.Append("(");
                        }
                        if (!string.IsNullOrEmpty(displayValue))
                        {
                            if (j > 1)
                            {
                                sb.Append(", ");
                            }
                            sb.Append(displayValue);
                        }
                    }
                    if (columns.Count > 1)
                    {
                        sb.Append(")");
                    }
                    sb.Append("\r\n");
                }
                if (hasMore)
                {
                    sb.Append("...");
                }
            }

            return sb.ToString();
        }
        protected string FindEntity(LuisResult result, bool ignoreAttributeNameAndValue)
        {
            Entity previouslySelectedEntity = this.SelectedEntity;

            string entityDisplayName = string.Empty;
            this.SelectedEntity = null;
            EntityRecommendation entityTypeEntity = result.RetrieveEntity(this.conversationId, EntityTypeNames.EntityType);
            if (entityTypeEntity != null)
            {
                string entityType = entityTypeEntity.Entity;
                EntityMetadata metadata = CrmHelper.RetrieveEntityMetadata(this.conversationId, entityType);
                if (metadata != null)
                {
                    entityDisplayName = entityType;
                    if (metadata.DisplayName != null && metadata.DisplayName.UserLocalizedLabel != null && !string.IsNullOrEmpty(metadata.DisplayName.UserLocalizedLabel.Label))
                    {
                        entityDisplayName = metadata.DisplayName.UserLocalizedLabel.Label;
                    }

                    EntityRecommendation dateEntity = result.RetrieveEntity(this.conversationId, EntityTypeNames.DateTime);
                    if (dateEntity != null)
                    {
                        //MODEBUG TODO await CountRecords(context, result);
                    }
                    else
                    {
                        EntityRecommendation firstName = result.RetrieveEntity(this.conversationId, EntityTypeNames.FirstName);
                        EntityRecommendation lastName = result.RetrieveEntity(this.conversationId, EntityTypeNames.LastName);
                        EntityRecommendation attributeName = result.RetrieveEntity(this.conversationId, EntityTypeNames.AttributeName);
                        EntityRecommendation attributeValue = result.RetrieveEntity(this.conversationId, EntityTypeNames.AttributeValue);

                        using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(this.conversationId))
                        {
                            QueryExpression expression = new QueryExpression(entityType);
                            expression.ColumnSet = new ColumnSet(true);
                            this.AddFilter(expression, metadata, firstName);
                            this.AddFilter(expression, metadata, lastName);

                            if (!ignoreAttributeNameAndValue)
                            {
                                this.AddFilter(expression, metadata, attributeName, attributeValue);
                            }
                            EntityCollection collection = serviceProxy.RetrieveMultiple(expression);
                            if (collection.Entities != null)
                            {
                                if (collection.Entities.Count == 1)
                                {
                                    this.SelectedEntity = collection.Entities[0];
                                }
                                else if (collection.Entities.Count > 1)
                                {
                                    this.FilteredEntities = collection.Entities.ToArray();
                                }
                            }
                        }
                    }
                }
            }
            if (this.SelectedEntity == null)
            {
                this.SelectedEntity = previouslySelectedEntity;
            }
            return entityDisplayName;
        }

        public static string BuildCommandList(string[] phrases)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < phrases.Length; i++)
            {
                sb.Append("\r\n");
                sb.Append(i + 1);
                sb.Append(". ");
                sb.Append(phrases[i]);
            }

            return sb.ToString();
        }
        protected void SetValue(Entity entity, EntityMetadata metadata, string attributeName, EntityRecommendation attributeValue)
        {
            if (attributeValue != null && !string.IsNullOrEmpty(attributeValue.Entity))
            {
                string att = CrmHelper.FindAttributeLogicalName(metadata, attributeName);

                AttributeMetadata attMetadata = metadata.Attributes.FirstOrDefault(a => a.LogicalName == att);
                if (string.IsNullOrEmpty(att))
                {
                    att = metadata.PrimaryNameAttribute;
                }

                if (!string.IsNullOrEmpty(att))
                {
                    object value = null;
                    switch (attMetadata.AttributeType)
                    {
                        case AttributeTypeCode.Integer:
                        case AttributeTypeCode.BigInt:
                            value = Int32.Parse(attributeValue.Entity);
                            break;
                        case AttributeTypeCode.Boolean:
                            value = bool.Parse(attributeValue.Entity);
                            break;
                        case AttributeTypeCode.DateTime:
                            value = DateTime.Parse(attributeValue.Entity);
                            break;
                        case AttributeTypeCode.Decimal:
                        case AttributeTypeCode.Money:
                            value = Decimal.Parse(attributeValue.Entity);
                            break;
                        case AttributeTypeCode.Double:
                            value = double.Parse(attributeValue.Entity);
                            break;
                        case AttributeTypeCode.Uniqueidentifier:
                            value = Guid.Parse(attributeValue.Entity);
                            break;
                        case AttributeTypeCode.Picklist:
                            value = new OptionSetValue()
                            {
                                Value = Int32.Parse(attributeValue.Entity)
                            };
                            break;
                        default:
                            value = attributeValue.Entity;
                            break;
                    }
                    entity[att] = value;
                }
            }
        }

        protected string GetAttributeDisplayValue(Entity entity, EntityMetadata metadata, string attributeName)
        {
            string att = CrmHelper.FindAttributeLogicalName(metadata, attributeName);
            if (entity != null && entity.Attributes != null && entity.Attributes.Contains(att) && entity[att] != null)
            {
                object attributeValue = entity[att];
                AttributeMetadata attMetadata = metadata.Attributes.FirstOrDefault(a => a.LogicalName == att);

                if (!string.IsNullOrEmpty(att))
                {
                    if (attributeValue is EntityReference)
                    {
                        return ((EntityReference)attributeValue).Name;
                    }
                    else if (attributeValue is Money)
                    {
                        return ((Money)attributeValue).Value.ToString("c");
                    }
                    else if (attributeValue is OptionSetValue)
                    {
                        int value = ((OptionSetValue)attributeValue).Value;
                        OptionMetadata optMetadata = ((EnumAttributeMetadata)attMetadata).OptionSet.Options.FirstOrDefault(o => o.Value != null && o.Value.Value == value);
                        if (optMetadata.Label != null && optMetadata.Label.UserLocalizedLabel != null && optMetadata.Label.UserLocalizedLabel != null)
                        {
                            return optMetadata.Label.UserLocalizedLabel.Label;
                        }
                        return string.Empty;
                    }
                    else
                    {
                        return attributeValue.ToString();
                    }
                }

            }
            return string.Empty;
        }
        protected string GetAttributeDisplayValue(EntityMetadata entityMetadata, string attributeName)
        {
            return this.GetAttributeDisplayValue(this.SelectedEntity, entityMetadata, attributeName);
        }
        protected string RetrieveEntityDisplayName(EntityMetadata entityMetadata, bool showPlural)
        {
            string displayName = (showPlural) ? entityMetadata.LogicalCollectionName : entityMetadata.LogicalName;
            if (!showPlural && entityMetadata.DisplayName != null && entityMetadata.DisplayName.UserLocalizedLabel != null && !string.IsNullOrEmpty(entityMetadata.DisplayName.UserLocalizedLabel.Label))
            {
                displayName = entityMetadata.DisplayName.UserLocalizedLabel.Label;
            }
            else if (showPlural && entityMetadata.DisplayCollectionName != null && entityMetadata.DisplayCollectionName.UserLocalizedLabel != null && !string.IsNullOrEmpty(entityMetadata.DisplayCollectionName.UserLocalizedLabel.Label))
            {
                displayName = entityMetadata.DisplayCollectionName.UserLocalizedLabel.Label;
            }
            return displayName;
        }
        protected void AddFilter(QueryExpression expression, EntityMetadata metadata, EntityRecommendation attribute)
        {
            if (attribute != null)
            {
                string att = CrmHelper.FindAttributeLogicalName(metadata, attribute.Type);
                if (!string.IsNullOrEmpty(att))
                {
                    expression.Criteria.AddCondition(att, ConditionOperator.Equal, attribute.Entity);
                }
                else
                {
                    expression.Criteria.AddCondition(metadata.PrimaryNameAttribute, ConditionOperator.Equal, attribute.Entity);
                }
            }
        }

        protected void AddFilter(QueryExpression expression, EntityMetadata entity, EntityRecommendation attributeName, EntityRecommendation attributeValue)
        {
            if (attributeName != null && attributeValue != null)
            {
                string att = CrmHelper.FindAttributeLogicalName(entity, attributeName.Entity);
                if (!string.IsNullOrEmpty(att))
                {
                    expression.Criteria.AddCondition(att, ConditionOperator.Equal, attributeValue.Entity.Replace(" . ", ".").Replace(" - ", "-").Replace(" @ ", "@"));
                }
            }
        }

        protected EntityMetadata SelectedEntityMetadata
        {
            get
            {
                return CrmHelper.RetrieveEntityMetadata(this.conversationId, this.SelectedEntity.LogicalName);
            }
        }

        protected Entity[] FilteredEntities
        {
            get
            {
                return ChatState.RetrieveChatState(this.conversationId).Get(ChatState.FilteredEntities) as Entity[];
            }
            set
            {
                ChatState.RetrieveChatState(this.conversationId).Set(ChatState.FilteredEntities, value);
            }
        }

        protected int CurrentPageIndex
        {
            get
            {
                if (ChatState.RetrieveChatState(this.conversationId).Get(ChatState.CurrentPageIndex) == null)
                {
                    return 0;
                }
                return (int)ChatState.RetrieveChatState(this.conversationId).Get(ChatState.CurrentPageIndex);
            }
            set
            {
                ChatState.RetrieveChatState(this.conversationId).Set(ChatState.CurrentPageIndex, value);
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
        public List<byte[]> Attachments
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