using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
#pragma warning disable 649
namespace CRMBot.Forms
{
    [Serializable]
    public class EmmyForm
    {
        [Prompt("I missed you. How are you?")]
        public string HowAreYou;
        [Prompt(":) I had a pretty good day today. Your Dad taught me to talk to you. How was your day?")]
        public string HowWasYourDay;
        [Prompt("I was very busy today and learned a lot. What did you do today?")]
        public string WhatDidYouDoToday;
        [Prompt("That sounds like a nice day. What did you learn today?")]
        public string WhatDidYouLearnToday;


        public static IForm<EmmyForm> BuildForm()
        {
            OnCompletionAsyncDelegate<EmmyForm> sayGoodBye = async (context, state) =>
            {
                await context.PostAsync($"I think you are awesome. Thanks for talking to me Emmy! I have to go for now. Talk to you later :)");
            };
            return new FormBuilder<EmmyForm>()
                    .Message("Hi Emmy!")
                    .AddRemainingFields()
                    .Confirm("No verification will be shown", state => false)
                    .OnCompletion(sayGoodBye)
                    .Build();
        }
        internal static IDialog<EmmyForm> MakeRootDialog()
        {
            return Chain.From(() => FormDialog.FromForm(EmmyForm.BuildForm));
        }
    }
}