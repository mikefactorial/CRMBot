using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Bot.Connector;
using CRMBot;
using System.Threading.Tasks;

namespace CRMBotUnitTests
{
    [TestClass]
    public class UnitTest1
    {
        public static string conversationId = Guid.NewGuid().ToString();

        [TestMethod]
        public void TestLocateIntent()
        {
            /*
            MessagesController controller = new MessagesController();
            Message message = new Message();
            message.Type = "Message";
            message.Text = "Find contact Susan Jones."; 
            message.Conversation.Id = conversationId;

            Task<Message> responseMessage = controller.Post(message);
            Assert.IsNotNull(responseMessage.Result);
            Assert.IsTrue(responseMessage.Result.Text.Contains("I found a"));
            */
        }
    }
}
