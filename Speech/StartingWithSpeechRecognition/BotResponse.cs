using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StartingWithSpeechRecognition
{

    public class BotResponse
    {
        public string conversationId { get; set; }
        public string language { get; set; }
        public string text { get; set; }
        public From from { get; set; }
        public To to { get; set; }
        public string replyToMessageId { get; set; }
        public List<Participant> participants { get; set; }
        public int totalParticipants { get; set; }
        public string channelMessageId { get; set; }
        public string channelConversationId { get; set; }
    }
}
