using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
#pragma warning disable 649
namespace CRMBot.Forms
{
    public enum LeadType { Partner, User };
    [Serializable]
    public class LeadForm
    {
        [Prompt("What's your {&}?")]
        public string Name;
        [Prompt("Thanks! Just one more easy question. Are you a CRM Partner or a CRM User? {||}")]
        public LeadType Type;
        public static IForm<LeadForm> BuildForm()
        {
            OnCompletionAsyncDelegate<LeadForm> processLead = async (context, state) =>
            {
                string[] nameSplit = state.Name.Split(' ');
                using (OrganizationServiceProxy serviceProxy = CrmHelper.CreateOrganizationService(Guid.Empty.ToString()))
                {
                    Microsoft.Xrm.Sdk.Entity newLead = new Microsoft.Xrm.Sdk.Entity("lead");
                    Guid leadId = Guid.NewGuid();
                    newLead["leadid"] = leadId;
                    newLead["subject"] = "CRM Bot Registrant";
                    if (nameSplit.Length > 0) newLead["firstname"] = nameSplit[0];
                    if (nameSplit.Length > 1) newLead["lastname"] = nameSplit[1];
                    newLead["cobalt_leadtype"] = new OptionSetValue() { Value = (state.Type == LeadType.Partner) ? 533470001 : 533470000 };
                    serviceProxy.Create(newLead);
                    await context.PostAsync($"Thanks for saying Hi {nameSplit[0]}! If you’d like to connect me to your CRM Organization go [here](http://www.cobalt.net/botregistration?id={leadId}).");
                }
            };
            return new FormBuilder<LeadForm>()
                    .Message("Hey there, I don't believe we've met.")
                    .OnCompletion(processLead)
                    .Build();
        }
        internal static IDialog<LeadForm> MakeRootDialog()
        {
            return Chain.From(() => FormDialog.FromForm(LeadForm.BuildForm));
        }
    }
}