using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.WebServiceClient;
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
        [Prompt("Nice to meet you {Name}! Just one more easy question. Are you a CRM Partner or a CRM User? {||}")]
        public LeadType? Type;
        public static IForm<LeadForm> BuildForm()
        {
            OnCompletionAsyncDelegate<LeadForm> processLead = async (context, state) =>
            {
                string[] nameSplit = state.Name.Split(' ');
                using (OrganizationWebProxyClient serviceProxy = CrmHelper.CreateOrganizationService(string.Empty, string.Empty))
                {
                    Microsoft.Xrm.Sdk.Entity newLead = new Microsoft.Xrm.Sdk.Entity("lead");
                    newLead["subject"] = "CRMUG Summit 2016 Lead";
                    if (nameSplit.Length > 0) newLead["firstname"] = nameSplit[0];
                    if (nameSplit.Length > 1) newLead["lastname"] = nameSplit[1];
                    newLead["cobalt_leadtype"] = new OptionSetValue() { Value = (state.Type == LeadType.Partner) ? 533470001 : 533470000 };
                    serviceProxy.Create(newLead);
                    await context.PostAsync($"Got it :) Thanks for saying Hi {newLead["firstname"]}! To register your CRM Organization go [here](http://www.cobalt.net/botregistration)");
                }
            };
            return new FormBuilder<LeadForm>()
                    .Message("Hey there, I don't believe we've met.")
                    .AddRemainingFields()
                    .Confirm("No verification will be shown", state => false)
                    .OnCompletion(processLead)
                    .Build();
        }
        internal static IDialog<LeadForm> MakeRootDialog()
        {
            return Chain.From(() => FormDialog.FromForm(LeadForm.BuildForm))
                .Do(async (context, order) =>
                {
                    try
                    {
                        var completed = await order;
                    }
                    catch (FormCanceledException<LeadForm> e)
                    {
                        string reply;
                        if (e.InnerException == null)
                        {
                            reply = $"Okay we're done with that for now";
                        }
                        else
                        {
                            reply = "Sorry, I've had a short circuit.  Please try again.";
                        }
                        await context.PostAsync(reply);
                    }
                });
        }
    }
}