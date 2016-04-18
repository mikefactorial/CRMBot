using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CRMBot.LuisResults
{
    public class Action
    {
        public bool triggered { get; set; }
        public string name { get; set; }
        public List<object> parameters { get; set; }
    }

    public class Intent
    {
        public string intent { get; set; }
        public double score { get; set; }
        public List<Action> actions { get; set; }
    }
    public class Resolution
    {
        public string date { get; set; }
    }
    public class Entity
    {
        public string entity { get; set; }
        public string type { get; set; }
        public int startIndex { get; set; }
        public int endIndex { get; set; }
        public double score { get; set; }
        public Resolution resolution { get; set; }
    }

    public class Result
    {
        public string query { get; set; }
        public List<Intent> intents { get; set; }
        public List<Entity> entities { get; set; }
    }
}