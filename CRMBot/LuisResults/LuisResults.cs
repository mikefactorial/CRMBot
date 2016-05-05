using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CRMBot.LuisResults
{
    public class EntityTypeNames
    {
        public const string EntityType = "EntityType";
        public const string Action = "Action";
        public const string CompanyName = "CompanyName";
        public const string AttributeName = "AttributeName";
        public const string AttributeValue = "AttributeValue";
        public const string EmailAddress = "EmailAddress";
        public const string DateTime = "builtin.datetime.date";
    }
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


        public string RetrieveIntention()
        {
            double max = 0.00;
            string bestIntention = string.Empty;
            foreach (var intent in this.intents)
            {
                if (max < intent.score)
                {
                    bestIntention = intent.intent;
                    max = intent.score;
                }
            }
            return bestIntention;
        }

        public Entity RetrieveEntity(string entityType)
        {
            double max = 0.00;
            Entity bestEntity = null;
            foreach (var entity in this.entities.Where(e => e.type == entityType))
            {
                if (max < entity.score)
                {
                    bestEntity = entity;
                    max = entity.score;
                }
            }
            return bestEntity;
        }

        public static Result Parse(string message)
        {
            var client = new RestClient("https://api.projectoxford.ai");
            var request = new RestRequest("/luis/v1/application?id=cc421661-4803-4359-b19b-35a8bae3b466&subscription-key=70c9f99320804782866c3eba387d54bf&q=" + message, Method.GET);
            // automatically deserialize result
            IRestResponse<Result> response = client.Execute<Result>(request);
            return response.Data;
        }
    }
}