using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StartingWithSpeechRecognition
{
    public class From
    {
        public string name { get; set; }
        public string channelId { get; set; }
        public string address { get; set; }
        public string id { get; set; }
        public bool isBot { get; set; }
    }

    public class To
    {
        public string name { get; set; }
        public string channelId { get; set; }
        public string address { get; set; }
        public string id { get; set; }
        public bool isBot { get; set; }
    }

    public class Participant
    {
        public string name { get; set; }
        public string channelId { get; set; }
        public string address { get; set; }
        public string id { get; set; }
        public bool isBot { get; set; }
    }

    public class BotRequest
    {
        public string type { get; set; }
        public string id { get; set; }
        public string conversationId { get; set; }
        public string created { get; set; }
        public string language { get; set; }
        public string text { get; set; }
        public List<object> attachments { get; set; }
        public From from { get; set; }
        public To to { get; set; }
        public List<Participant> participants { get; set; }
        public int totalParticipants { get; set; }
        public List<object> mentions { get; set; }
        public string channelMessageId { get; set; }
        public string channelConversationId { get; set; }
        public List<object> hashtags { get; set; }
    }
}
