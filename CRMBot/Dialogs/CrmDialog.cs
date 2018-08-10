using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.WebServiceClient;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace CRMBot.Dialogs
{
    public class CrmAttachment
    {
        public byte[] Attachment { get; set; }
        public string FileName { get; set; }
    }
    //[LuisModel("cc421661-4803-4359-b19b-35a8bae3b466", "70c9f99320804782866c3eba387d54bf")]
    [LuisModel("64c400cf-b36d-4874-bd01-1c7567e57d8a", "a03f8796d25a493dac9ff9e8ad2b15a6")]
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
            "\"Follow up July 12 2020\"",
            "\"Show me the name, status, date created etc.\"",
            "Send me an image and say \"Attach as 'Profile Pic'\""
        };

        public static string[] FindActionPhrases = new string[]
        {
            "\"Find contact John Connor.\"",
            "\"Find lead Sarah Connor.\"",
            "\"Find opportunity Kyle Reese.\""
        };

        public static string defaultMessage = $"Sorry, I didn't understand that. Try saying {CrmDialog.BuildCommandList(CrmDialog.WelcomePhrases)}";

        public static string waitMessage = "Got it...Give me a second while I ";
        public const int MAX_RECORDS_TO_SHOW_PER_PAGE = int.MaxValue;

        private string channelId = string.Empty;
        private string userId = string.Empty;

        public CrmDialog(Microsoft.Bot.Connector.Activity message)
        {
            this.channelId = message.ChannelId;
            this.userId = message.From.Id;
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
                EntityRecommendation dateEntity = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.DateTime);

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
                ChatState chatState = ChatState.RetrieveChatState(this.channelId, this.userId);
                using (OrganizationWebProxyClient serviceProxy = chatState.CreateOrganizationService())
                {
                    serviceProxy.Create(entity);
                }
                if (date != DateTime.MinValue)
                {
                    ShowCurrentRecordSelection(context, result, $"Okay...I've created task to follow up with {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]} on { date.ToLongDateString() }");
                }
                else
                {
                    ShowCurrentRecordSelection(context, result, $"Okay...I've created task to follow up with {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]}");
                }
            }
            else
            {
                ShowCurrentRecordSelection(context, result, $"Hmmm...I'm not sure who to follow up with. Say for example {BuildCommandList(FindActionPhrases)}");
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Send")]
        public async Task Send(IDialogContext context, LuisResult result)
        {
            ShowCurrentRecordSelection(context, result, defaultMessage);
            context.Wait(MessageReceived);
        }
        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            ChatState chatState = ChatState.RetrieveChatState(this.channelId, this.userId);
            int selection = -1;

            EntityRecommendation ordinal = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.Ordinal);
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
                        EntityMetadata metadata = chatState.RetrieveEntityMetadata(this.SelectedEntity.LogicalName);
                        string displayName = RetrieveEntityDisplayName(metadata, false);
                        ShowCurrentRecordSelection(context, result, $"Got it. You've selected the {displayName} named {this.SelectedEntity[metadata.PrimaryNameAttribute]}. Now say {BuildCommandList(ActionPhrases)}");
                    }
                    else
                    {
                        ShowCurrentRecordSelection(context, result, $"Hmmm. I couldn't select that record. Make sure the number is within the current range of selected records (i.e. 1 to {this.FilteredEntities.Length}).");
                    }
                }
                else if (this.FilteredEntities == null || this.FilteredEntities.Length == 0)
                {
                    ShowCurrentRecordSelection(context, result, $"Hmmm. It looks like you're trying to select a record. However, there are no filtered records to choose from. Say something like \"How many leads were created last week.\" to select a list of records.");
                }
                else
                {
                    ShowCurrentRecordSelection(context, result, $"Hmmm. I couldn't select that record. Make sure the number is within the current range of selected records (i.e. 1 to {this.FilteredEntities.Length}).");
                }
            }
            //Summit stuff
            else if (result.Query.ToLower().Contains("you ready"))
            {
                ShowCurrentRecordSelection(context, result, $"I was built ready {chatState.UserFirstName}! Don't you go screwing it up for me...");
            }
            else if (result.Query.ToLower().StartsWith("next"))
            {
                this.CurrentPageIndex = this.CurrentPageIndex + 1;
                BuildFilteredEntitiesList(context, result, "Select a record from the list below");
            }
            else if (result.Query.ToLower().StartsWith("back"))
            {
                if (this.CurrentPageIndex >= 0)
                {
                    this.CurrentPageIndex = this.CurrentPageIndex - 1;
                }
                this.BuildFilteredEntitiesList(context, result, "Select a record from the list below");
            }
            else if (result.Query.ToLower().StartsWith("help"))
            {
                if (this.SelectedEntity == null)
                {
                    await context.PostAsync($"Hi {chatState.UserFirstName}! Here are some commands you can try... {CrmDialog.BuildCommandList(CrmDialog.HelpPhrases)}.");
                }
                else
                {
                    EntityMetadata metadata = chatState.RetrieveEntityMetadata(this.SelectedEntity.LogicalName);
                    string displayName = RetrieveEntityDisplayName(metadata, false);
                    ShowCurrentRecordSelection(context, result, $"You've selected a {displayName} named {this.SelectedEntity[metadata.PrimaryNameAttribute]}? Now say {BuildCommandList(ActionPhrases)}");
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
            else if (result.Query.ToLower().StartsWith("thank"))
            {
                await context.PostAsync($"You're welcome {chatState.UserFirstName}!");
            }
            else if (result.Query.ToLower().StartsWith("say"))
            {
                await context.PostAsync(result.Query.Substring(result.Query.ToLower().IndexOf("say") + 4));
            }
            else if (result.Query.ToLower().StartsWith("bye") || result.Query.ToLower().Contains("see ya") || result.Query.ToLower().Contains("bye") || result.Query.ToLower().Contains("later"))
            {
                chatState.Data.Clear();
                await context.PostAsync($"See you later {chatState.UserFirstName}...");
            }
            else
            {
                if (result.Query.ToLower().Contains("hello") || result.Query.ToLower().Contains("hi"))
                {
                    await context.PostAsync($"Hey there {chatState.UserFirstName}!");
                }
                else if (result.Query.ToLower().Contains("what's up") || result.Query.ToLower().Contains("waddup") || result.Query.ToLower().Contains("sup") || result.Query.ToLower().Contains("whats up"))
                {
                    await context.PostAsync($"Nothing much, just serving up some Dynamics data like it's my job.");
                }
                ShowCurrentRecordSelection(context, result, $"To get started say something like {CrmDialog.BuildCommandList(CrmDialog.WelcomePhrases)}.");
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Open")]
        public async Task Open(IDialogContext context, LuisResult result)
        {
            this.FindEntity(result, false);

            if (this.SelectedEntity != null)
            {
                ChatState chatState = ChatState.RetrieveChatState(this.channelId, this.userId);
                if (context.Activity.From.Properties.ContainsKey("crmUrl"))
                {
                    ShowCurrentRecordSelection(context, result, "Give me a second. I'm opening that record...");

                    Microsoft.Bot.Connector.Activity backChannelReply = ((Microsoft.Bot.Connector.Activity)context.Activity).CreateReply();
                    backChannelReply.Text = $"{this.SelectedEntity.LogicalName}|{this.SelectedEntity.Id.ToString()}";
                    backChannelReply.Name = "openForm";
                    backChannelReply.Recipient = context.Activity.From;
                    backChannelReply.Type = Microsoft.Bot.Connector.ActivityTypes.Event;
                    await context.PostAsync(backChannelReply);
                }
                else
                {
                    //[here](ms-dynamicsxrm://?pagetype=view&etn={this.SelectedEntity.LogicalName}&id={this.SelectedEntity.Id.ToString()}
                    //Open in mobile client: ms-dynamicsxrm://?pagetype=view&etn=contact&id=899D4FCF-F4D3-E011-9D26-00155DBA3819
                    //Open in browser: http://myorg.crm.dynamics.com/main.aspx?etn=account&pagetype=entityrecord&id=%7B91330924-802A-4B0D-A900-34FD9D790829%7D
                    //await context.PostAsync($"Click [here]({chatState.OrganizationUrl}/main.aspx?etn={this.SelectedEntity.LogicalName}&pagetype=entityrecord&id={this.SelectedEntity.Id.ToString()}) to open the record in your browser or click [here](ms-dynamicsxrm://?pagetype=view&etn={this.SelectedEntity.LogicalName}&id={this.SelectedEntity.Id.ToString()}) to open the record in the mobile client.");
                    await context.PostAsync($"Click [here]({chatState.OrganizationUrl}/main.aspx?etn={this.SelectedEntity.LogicalName}&pagetype=entityrecord&id={this.SelectedEntity.Id.ToString()}) to open the record in your browser.");
                }
            }
            else
            {
                ShowCurrentRecordSelection(context, result, defaultMessage);
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Update")]
        public async Task Update(IDialogContext context, LuisResult result)
        {
            this.FindEntity(result, true);
            if (this.SelectedEntity != null)
            {
                ChatState chatState = ChatState.RetrieveChatState(this.channelId, this.userId);
                EntityMetadata metadata = chatState.RetrieveEntityMetadata(this.SelectedEntity.LogicalName);
                EntityRecommendation attributeName = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.AttributeName, EntityTypeNames.DisplayField);
                EntityRecommendation attributeValue = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.AttributeValue);
                string displayName = RetrieveEntityDisplayName(metadata, false);
                if (attributeName != null && attributeValue != null)
                {
                    string att = result.FindAttributeLogicalName(metadata, attributeName.Entity);
                    if (!string.IsNullOrEmpty(att))
                    {
                        this.SelectedEntity[att] = attributeValue.Entity;
                        using (OrganizationWebProxyClient serviceProxy = chatState.CreateOrganizationService())
                        {
                            serviceProxy.Update(this.SelectedEntity);
                            ShowCurrentRecordSelection(context, result, $"I've update the {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]} record with the new {attributeName.Entity} {attributeValue.Entity}");
                        }
                    }
                    else
                    {
                        ShowCurrentRecordSelection(context, result, $"I wasn't able to update the {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]} record. I didn't recognize the field name.");
                    }
                }
                else
                {
                    ShowCurrentRecordSelection(context, result, $"I wasn't able to update the {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]} record. I didn't recognize the field name and value.");
                }
            }
            else
            {
                ShowCurrentRecordSelection(context, result, defaultMessage);
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Create")]
        public async Task Create(IDialogContext context, LuisResult result)
        {
            ChatState chatState = ChatState.RetrieveChatState(channelId, userId);
            EntityRecommendation entityTypeEntity = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.EntityType);
            if (entityTypeEntity != null && !string.IsNullOrEmpty(entityTypeEntity.Entity))
            {
                EntityMetadata metadata = chatState.RetrieveEntityMetadata(entityTypeEntity.Entity);
                if (metadata != null)
                {
                    EntityRecommendation firstName = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.FirstName);
                    EntityRecommendation lastName = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.LastName);
                    EntityRecommendation attributeName = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.AttributeName, EntityTypeNames.DisplayField);
                    EntityRecommendation attributeValue = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.AttributeValue);

                    Entity entity = new Entity(entityTypeEntity.Entity);
                    if (attributeValue != null)
                    {
                        if (attributeName != null)
                        {
                            string att = result.FindAttributeLogicalName(metadata, attributeName.Entity);
                            SetValue(result, entity, metadata, att, attributeValue);
                        }
                        else
                        {
                            SetValue(result, entity, metadata, metadata.PrimaryNameAttribute, attributeValue);
                        }
                    }
                    else
                    {
                        if (firstName != null)
                        {
                            SetValue(result, entity, metadata, "firstname", firstName);
                        }
                        if (lastName != null)
                        {
                            SetValue(result, entity, metadata, "lastname", lastName);
                        }
                    }

                    if (metadata.IsActivity.HasValue && metadata.IsActivity.Value && this.SelectedEntity != null)
                    {
                        if (metadata.Attributes.Any(a => a.LogicalName == "customerid") && (this.SelectedEntity.LogicalName == "account" || this.SelectedEntity.LogicalName == "contact"))
                        {
                            entity["customerid"] = new EntityReference(this.SelectedEntity.LogicalName, this.SelectedEntity.Id);
                        }
                    }
                    using (OrganizationWebProxyClient serviceProxy = chatState.CreateOrganizationService())
                    {
                        Guid id = serviceProxy.Create(entity);
                        this.SelectedEntity = serviceProxy.Retrieve(entity.LogicalName, id, new ColumnSet(true));
                    }

                    string displayName = RetrieveEntityDisplayName(metadata, false);
                    ShowCurrentRecordSelection(context, result, $"I've created a new {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]}. If you want to set additional fields say something like 'Update {displayName} set home phone to 555-555-5555'");
                }
                else
                {
                    if (entityTypeEntity != null && !string.IsNullOrEmpty(entityTypeEntity.Entity))
                    {
                        ShowCurrentRecordSelection(context, result, $"Hmmm. I couldn't find any records called {entityTypeEntity.Entity} in Dynamics.");
                    }
                    else
                    {
                        ShowCurrentRecordSelection(context, result, $"Hmmm. I couldn't find any records that matched your query in Dynamics.");
                    }
                }
            }
            else
            {
                ShowCurrentRecordSelection(context, result, defaultMessage);
            }
            context.Wait(MessageReceived);
        }
        [LuisIntent("Locate")]
        public async Task Locate(IDialogContext context, LuisResult result)
        {
            ChatState chatState = ChatState.RetrieveChatState(this.channelId, this.userId);
            Guid previouslySelectedEntityId = (this.SelectedEntity != null) ? this.SelectedEntity.Id : Guid.Empty;
            string entityDisplayName = this.FindEntity(result, false);

            string parentEntity = string.Empty;
            EntityRecommendation entityTypeEntity = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.EntityType);
            if ((this.SelectedEntity == null || previouslySelectedEntityId == this.SelectedEntity.Id) && entityTypeEntity != null && !string.IsNullOrEmpty(entityTypeEntity.Entity))
            {
                EntityMetadata entityMetadata = chatState.RetrieveEntityMetadata(entityTypeEntity.Entity);
                if (entityMetadata != null)
                {
                    bool associatedEntities = false;
                    QueryExpression expression = new QueryExpression(entityTypeEntity.Entity);

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
                    else if (this.SelectedEntity != null)
                    {
                        var relationships = entityMetadata.ManyToOneRelationships.Where(m => m.ReferencedEntity == this.SelectedEntity.LogicalName).ToArray();
                        foreach (var relationship in relationships)
                        {
                            associatedEntities = true;
                            expression.Criteria.AddCondition(relationship.ReferencingAttribute, ConditionOperator.Equal, this.SelectedEntity.Id);
                        }
                    }

                    string whenString = string.Empty;

                    EntityRecommendation dateEntity = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.DateTime);
                    if (dateEntity != null)
                    {
                        List<DateTime> dates = dateEntity.ParseDateTimes();
                        if (dates != null && dates.Count > 0)
                        {
                            string action = "created";
                            EntityRecommendation actionEntity = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.Action);
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
                                expression.AddOrder(field, OrderType.Ascending);
                            }
                        }
                    }
                    if (expression.Orders == null || expression.Orders.Count <= 0)
                    {
                        expression.AddOrder(entityMetadata.PrimaryNameAttribute, OrderType.Ascending);
                    }
                    using (OrganizationWebProxyClient serviceProxy = chatState.CreateOrganizationService())
                    {
                        EntityCollection collection = serviceProxy.RetrieveMultiple(expression);
                        if (collection.Entities != null)
                        {
                            this.FilteredEntities = collection.Entities.ToArray();
                            string displayName = RetrieveEntityDisplayName(entityMetadata, this.FilteredEntities.Length != 1);

                            if (associatedEntities)
                            {
                                EntityMetadata selectedEntityMetadata = chatState.RetrieveEntityMetadata(this.SelectedEntity.LogicalName);

                                string selectedEntityDisplayName = RetrieveEntityDisplayName(selectedEntityMetadata, false);
                                this.BuildFilteredEntitiesList(context, result, $"I found {collection.Entities.Count} {displayName} for the {selectedEntityDisplayName} {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]} {whenString}");
                            }
                            else
                            {
                                this.BuildFilteredEntitiesList(context, result, $"I found {collection.Entities.Count} {displayName} {whenString} in Dynamics.");
                            }
                        }
                    }
                }
                else
                {
                    if (entityTypeEntity != null && !string.IsNullOrEmpty(entityTypeEntity.Entity))
                    {
                        ShowCurrentRecordSelection(context, result, $"Hmmm. I couldn't find any records called {entityTypeEntity.Entity} in Dynamics.");
                    }
                    else
                    {
                        ShowCurrentRecordSelection(context, result, $"Hmmm. I couldn't find any records that matched your query in Dynamics.");
                    }
                }
            }
            else
            {
                if (this.SelectedEntity != null)
                {
                    EntityMetadata metadata = ChatState.RetrieveChatState(this.channelId, this.userId).RetrieveEntityMetadata(this.SelectedEntity.LogicalName);
                    EntityRecommendation displayField = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.DisplayField, EntityTypeNames.AttributeName);
                    string displayName = RetrieveEntityDisplayName(metadata, false);
                    if (displayField != null)
                    {
                        string att = result.FindAttributeLogicalName(metadata, displayField.Entity);
                        string displayValue = GetAttributeDisplayValue(result, metadata, att);
                        if (!string.IsNullOrEmpty(displayValue))
                        {
                            ShowCurrentRecordSelection(context, result, $"{this.SelectedEntity[metadata.PrimaryNameAttribute]}'s {displayField.Entity} is {displayValue}");
                        }
                        else
                        {
                            ShowCurrentRecordSelection(context, result, $"The {displayName} {this.SelectedEntity[metadata.PrimaryNameAttribute]} does not have a {displayField.Entity}");
                        }
                    }
                    else
                    {
                        ShowCurrentRecordSelection(context, result, $"I found a {entityDisplayName} named {this.SelectedEntity[metadata.PrimaryNameAttribute]} what would you like to do next? You can say {BuildCommandList(ActionPhrases)}");
                    }
                }
                else if (this.FilteredEntities != null && this.FilteredEntities.Length > 0)
                {
                    this.BuildFilteredEntitiesList(context, result, $"I found {this.FilteredEntities.Length} {entityDisplayName} that match. Select a record from the list below");
                }
                else
                {
                    ShowCurrentRecordSelection(context, result, $"Hmmm...I couldn't find that {entityDisplayName}.");
                }
            }
            context.Wait(MessageReceived);
        }

        [LuisIntent("Attach")]
        public async Task Attach(IDialogContext context, LuisResult result)
        {
            this.FindEntity(result, true);
            if (this.SelectedEntity == null)
            {
                ShowCurrentRecordSelection(context, result, $"There's nothing to attach the file to. Say {BuildCommandList(FindActionPhrases)} to find a record to attach this to.");
            }
            else if (this.Attachments == null || this.Attachments.Count == 0)
            {
                ShowCurrentRecordSelection(context, result, $"There's nothing to attach. Send me a file and Say {BuildCommandList(AttachmentActionPhrases)} to attach a file.");
            }
            else
            {
                EntityRecommendation subjectEntity = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.AttributeValue);
                string subject = "Attachment";
                if (subjectEntity != null)
                {
                    subject = subjectEntity.Entity;
                }
                int i = 1;
                foreach (CrmAttachment attachment in this.Attachments)
                {
                    Entity annotation = new Entity("annotation");
                    annotation["objectid"] = new EntityReference() { Id = this.SelectedEntity.Id, LogicalName = this.SelectedEntity.LogicalName };
                    string encodedData = System.Convert.ToBase64String(attachment.Attachment);

                    annotation["filename"] = attachment.FileName;
                    annotation["subject"] = subject;
                    annotation["mimetype"] = MimeMapping.GetMimeMapping(attachment.FileName);
                    annotation["documentbody"] = encodedData;
                    ChatState chatState = ChatState.RetrieveChatState(this.channelId, this.userId);
                    using (OrganizationWebProxyClient serviceProxy = chatState.CreateOrganizationService())
                    {
                        serviceProxy.Create(annotation);
                        ShowCurrentRecordSelection(context, result, $"Okay. I've attached the file to {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]} as a note with the Subject '{annotation["subject"]}'");
                    }
                }
                this.Attachments = null;
            }

            context.Wait(MessageReceived);
        }

        protected async void ShowCurrentRecordSelection(IDialogContext context, LuisResult result, string messageText)
        {
            if (this.SelectedEntity != null)
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                var message = context.MakeMessage();
                message.Text = messageText;

                message.SuggestedActions = new Microsoft.Bot.Connector.SuggestedActions()
                {
                    Actions = new List<Microsoft.Bot.Connector.CardAction>()
                    {
                        new Microsoft.Bot.Connector.CardAction(){ Title = $"Forget {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]}", Type = Microsoft.Bot.Connector.ActionTypes.ImBack, Value="Forget" },
                        new Microsoft.Bot.Connector.CardAction(){ Title = $"Open {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]}", Type = Microsoft.Bot.Connector.ActionTypes.ImBack, Value="Open" },
                    }
                };
                await context.PostAsync(message);
            }
        }

        protected async void BuildFilteredEntitiesList(IDialogContext context, LuisResult result, string titleMessage)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            var message = context.MakeMessage();
            message.Text = titleMessage;

            AdaptiveCards.AdaptiveCard plCard = null;
            Microsoft.Bot.Connector.Attachment attachment = null;

            if (this.SelectedEntity != null)
            {
                message.SuggestedActions = new Microsoft.Bot.Connector.SuggestedActions()
                {
                    Actions = new List<Microsoft.Bot.Connector.CardAction>()
                    {
                        new Microsoft.Bot.Connector.CardAction(){ Title = $"Forget {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]}", Type = Microsoft.Bot.Connector.ActionTypes.ImBack, Value="Forget" },
                        new Microsoft.Bot.Connector.CardAction(){ Title = $"Open {this.SelectedEntity[this.SelectedEntityMetadata.PrimaryNameAttribute]}", Type = Microsoft.Bot.Connector.ActionTypes.ImBack, Value="Open" },
                    }
                };
            }
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
                        ChatState chatState = ChatState.RetrieveChatState(this.channelId, this.userId);
                        entityMetadata = chatState.RetrieveEntityMetadata(this.FilteredEntities[i].LogicalName);

                        using (OrganizationWebProxyClient serviceProxy = chatState.CreateOrganizationService())
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

                    StringBuilder cardText = new StringBuilder($"");
                    for (int j = 0; j < columns.Count; j++)
                    {
                        string column = columns[j];
                        string displayValue = this.GetAttributeDisplayValue(result, this.FilteredEntities[i], entityMetadata, column);

                        if (!string.IsNullOrEmpty(displayValue))
                        {
                            if (j > 0)
                            {
                                plCard.Body.Add(new AdaptiveCards.TextBlock() { Text = displayValue, Wrap = true, });
                            }
                            else
                            {
                                plCard = new AdaptiveCards.AdaptiveCard()
                                {
                                    Title = titleMessage,
                                };
                                plCard.Body.Add(new AdaptiveCards.TextBlock() { Text = displayValue, Wrap = true, Size = AdaptiveCards.TextSize.Large, Weight = AdaptiveCards.TextWeight.Bolder });
                            }
                        }
                    }


                    AdaptiveCards.SubmitAction recordButton = new AdaptiveCards.SubmitAction()
                    {
                        Data = (i + 1).ToString(),
                        Title = "Select this Record"
                    };
                    plCard.Actions.Add(recordButton);
                    attachment = new Microsoft.Bot.Connector.Attachment()
                    {
                        ContentType = AdaptiveCards.AdaptiveCard.ContentType,
                        Content = plCard
                    };
                    message.Attachments.Add(attachment);
                }

                plCard = new AdaptiveCards.AdaptiveCard()
                {
                    Title = titleMessage,
                };

                if (hasMore)
                {
                    if (start == 0)
                    {
                        AdaptiveCards.SubmitAction nextButton = new AdaptiveCards.SubmitAction()
                        {
                            Data = "next",
                            Title = "Show Next Page..."
                        };

                        plCard.Actions.Add(nextButton);
                    }
                    else
                    {
                        AdaptiveCards.SubmitAction backButton = new AdaptiveCards.SubmitAction()
                        {
                            Data = "back",
                            Title = "Show Previous Page..."
                        };

                        plCard.Actions.Add(backButton);
                        AdaptiveCards.SubmitAction nextButton = new AdaptiveCards.SubmitAction()
                        {
                            Data = "next",
                            Title = "Show Next Page..."
                        };

                        plCard.Actions.Add(nextButton);
                    }
                }
                else if (start > 0)
                {
                    AdaptiveCards.SubmitAction backButton = new AdaptiveCards.SubmitAction()
                    {
                        Data = "back",
                        Title = "Show Previous Page...",
                    };

                    plCard.Actions.Add(backButton);
                }
            }

            if (plCard.Actions.Count > 0)
            {
                attachment = new Microsoft.Bot.Connector.Attachment()
                {
                    ContentType = AdaptiveCards.AdaptiveCard.ContentType,
                    Content = plCard
                };
                message.Attachments.Add(attachment);
            }

            message.AttachmentLayout = "carousel";
            await context.PostAsync(message);
            if (FilteredEntities.Length == 0)
            {
                ShowCurrentRecordSelection(context, result, string.Empty);
            }
        }

        protected string FindEntity(LuisResult result, bool ignoreAttributeNameAndValue)
        {
            Entity previouslySelectedEntity = this.SelectedEntity;
            this.FilteredEntities = new Entity[] { };
            string entityDisplayName = string.Empty;
            this.SelectedEntity = null;
            EntityRecommendation entityTypeEntity = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.EntityType);
            if (entityTypeEntity != null && !string.IsNullOrEmpty(entityTypeEntity.Entity))
            {
                ChatState chatState = ChatState.RetrieveChatState(this.channelId, this.userId);
                EntityMetadata metadata = chatState.RetrieveEntityMetadata(entityTypeEntity.Entity);
                if (metadata != null)
                {
                    entityDisplayName = entityTypeEntity.Entity;
                    if (metadata.DisplayName != null && metadata.DisplayName.UserLocalizedLabel != null && !string.IsNullOrEmpty(metadata.DisplayName.UserLocalizedLabel.Label))
                    {
                        entityDisplayName = metadata.DisplayName.UserLocalizedLabel.Label;
                    }

                    EntityRecommendation dateEntity = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.DateTime);
                    EntityRecommendation firstName = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.FirstName);
                    EntityRecommendation lastName = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.LastName);
                    EntityRecommendation attributeName = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.AttributeName, EntityTypeNames.DisplayField);
                    EntityRecommendation attributeValue = result.RetrieveEntity(this.channelId, this.userId, EntityTypeNames.AttributeValue);

                    using (OrganizationWebProxyClient serviceProxy = chatState.CreateOrganizationService())
                    {
                        QueryExpression expression = new QueryExpression(entityTypeEntity.Entity);
                        expression.ColumnSet = new ColumnSet(true);
                        if (attributeValue == null || ignoreAttributeNameAndValue)
                        {
                            if ((firstName != null && attributeValue != null && firstName.Score != null && attributeValue.Score != null && firstName.Score > attributeValue.Score) || attributeValue == null || ignoreAttributeNameAndValue)
                            {
                                if (firstName != null)
                                {
                                    this.AddFilter(result, expression, metadata, firstName);
                                }
                            }
                            if ((lastName != null && attributeValue != null && lastName.Score != null && attributeValue.Score != null && lastName.Score > attributeValue.Score) || attributeValue == null || ignoreAttributeNameAndValue)
                            {
                                if (lastName != null)
                                {
                                    this.AddFilter(result, expression, metadata, lastName);
                                }
                            }

                        }
                        else if (!ignoreAttributeNameAndValue && attributeName != null && attributeValue != null)
                        {
                            this.AddFilter(result, expression, metadata, result.Query, entityTypeEntity.Entity, attributeName, attributeValue);
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
        protected void SetValue(LuisResult result, Entity entity, EntityMetadata metadata, string attributeName, EntityRecommendation attributeValue)
        {
            if (attributeValue != null && !string.IsNullOrEmpty(attributeValue.Entity))
            {
                string att = result.FindAttributeLogicalName(metadata, attributeName);

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
        public static string ParseCrmUrl(Microsoft.Bot.Connector.Activity message)
        {
            if (message.From.Properties.ContainsKey("crmUrl"))
            {
                return message.From.Properties["crmUrl"].ToString();
            }
            else if (!string.IsNullOrEmpty(message.Text))
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

        protected string GetAttributeDisplayValue(LuisResult result, Entity entity, EntityMetadata metadata, string attributeName)
        {
            string att = result.FindAttributeLogicalName(metadata, attributeName);
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
        protected string GetAttributeDisplayValue(LuisResult result, EntityMetadata entityMetadata, string attributeName)
        {
            return this.GetAttributeDisplayValue(result, this.SelectedEntity, entityMetadata, attributeName);
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
        protected void AddFilter(LuisResult result, QueryExpression expression, EntityMetadata metadata, EntityRecommendation attribute)
        {
            if (attribute != null)
            {
                string att = result.FindAttributeLogicalName(metadata, attribute.Type);
                if (!string.IsNullOrEmpty(att))
                {
                    expression.Criteria.AddCondition(att, ConditionOperator.Like, attribute.Entity + "%");
                }
                else
                {
                    expression.Criteria.AddCondition(metadata.PrimaryNameAttribute, ConditionOperator.Like, attribute.Entity + "%");
                }
            }
        }

        protected void AddFilter(LuisResult result, QueryExpression expression, EntityMetadata entity, string query, string entityType, EntityRecommendation attributeName, EntityRecommendation attributeValue)
        {
            string attrName = (attributeName != null) ? attributeName.Entity : entity.PrimaryNameAttribute;

            int attributeValueIndex = (attributeName != null) ? (query.IndexOf(attrName) + attrName.Length) : (query.IndexOf(entityType) + entityType.Length);
            int attributeValueLength = query.Length - attributeValueIndex;
            
            string attrValue = (attributeValue != null) ? attributeValue.Entity : query.Substring(attributeValueIndex, attributeValueLength);
            if (!string.IsNullOrEmpty(attrName) && !string.IsNullOrEmpty(attrValue))
            {
                string att = result.FindAttributeLogicalName(entity, attrName);
                if (!string.IsNullOrEmpty(att))
                {
                    expression.Criteria.AddCondition(att, ConditionOperator.Like, "%" + attrValue + "%");
                }
            }
        }

        protected EntityMetadata SelectedEntityMetadata
        {
            get
            {
                ChatState chatState = ChatState.RetrieveChatState(this.channelId, this.userId);
                return chatState.RetrieveEntityMetadata(this.SelectedEntity.LogicalName);
            }
        }

        protected Entity[] FilteredEntities
        {
            get
            {
                return ChatState.RetrieveChatState(this.channelId, this.userId).Get(ChatState.FilteredEntities) as Entity[];
            }
            set
            {
                this.CurrentPageIndex = 0;
                ChatState.RetrieveChatState(this.channelId, this.userId).Set(ChatState.FilteredEntities, value);
            }
        }

        protected int CurrentPageIndex
        {
            get
            {
                if (ChatState.RetrieveChatState(this.channelId, this.userId).Get(ChatState.CurrentPageIndex) == null)
                {
                    return 0;
                }
                return (int)ChatState.RetrieveChatState(this.channelId, this.userId).Get(ChatState.CurrentPageIndex);
            }
            set
            {
                ChatState.RetrieveChatState(this.channelId, this.userId).Set(ChatState.CurrentPageIndex, value);
            }
        }

        protected Entity SelectedEntity
        {
            get
            {
                return ChatState.RetrieveChatState(this.channelId, this.userId).Get(ChatState.SelectedEntity) as Entity;
            }
            set
            {
                ChatState.RetrieveChatState(this.channelId, this.userId).Set(ChatState.SelectedEntity, value);
            }
        }
        public List<CrmAttachment> Attachments
        {
            get
            {
                return ChatState.RetrieveChatState(this.channelId, this.userId).Get(ChatState.Attachments) as List<CrmAttachment>;
            }
            set
            {
                ChatState.RetrieveChatState(this.channelId, this.userId).Set(ChatState.Attachments, value);
            }
        }
    }
}