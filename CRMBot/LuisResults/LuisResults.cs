using Microsoft.Xrm.Sdk.Metadata;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;

namespace CRMBot.LuisResults
{
    public class EntityTypeNames
    {
        public string EntityTypeName;
        public double EntityThreashold;

        public static EntityTypeNames EntityType = new EntityTypeNames() { EntityTypeName = "EntityType", EntityThreashold = .1 };
        public static EntityTypeNames Action = new EntityTypeNames() { EntityTypeName = "Action", EntityThreashold = .1 };
        public static EntityTypeNames CompanyName = new EntityTypeNames() { EntityTypeName = "CompanyName", EntityThreashold = .1 };
        public static EntityTypeNames AttributeName = new EntityTypeNames() { EntityTypeName = "AttributeName", EntityThreashold = .1 };
        public static EntityTypeNames AttributeValue = new EntityTypeNames() { EntityTypeName = "AttributeValue", EntityThreashold = .1 };
        public static EntityTypeNames EmailAddress = new EntityTypeNames() { EntityTypeName = "EmailAddress", EntityThreashold = .1 };
        public static EntityTypeNames DateTime = new EntityTypeNames() { EntityTypeName = "builtin.datetime.date", EntityThreashold = 0 };

        public static EntityTypeNames FirstName = new EntityTypeNames() { EntityTypeName = "ContactName:FirstName", EntityThreashold = 0 };
        public static EntityTypeNames LastName = new EntityTypeNames() { EntityTypeName = "ContactName:LastName", EntityThreashold = 0 };
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

        public DateTime[] ParseDateTimes()
        {
            if (this.resolution != null && !string.IsNullOrEmpty(this.resolution.date))
            {
                DateTime singleDate;
                if (DateTime.TryParse(this.resolution.date, out singleDate))
                {
                    return new DateTime[] { singleDate };
                }
                else
                {
                    if (this.resolution.date.Contains("-W"))
                    {
                        int year = Int32.Parse(this.resolution.date.Substring(0, 4));
                        int index = this.resolution.date.IndexOf("-W") + 2;
                        int week = Int32.Parse(this.resolution.date.Substring(index, this.resolution.date.Length - index));
                        return new DateTime[] { FirstDateOfWeekISO8601(year, week, 0), FirstDateOfWeekISO8601(year, week, 7) };
                    }
                }
            }
            return null;
        }

        public DateTime FirstDateOfWeekISO8601(int year, int weekOfYear, int daysToAdd)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            DateTime firstThursday = jan1.AddDays(daysOffset);
            var cal = CultureInfo.CurrentCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            var weekNum = weekOfYear;
            if (firstWeek <= 1)
            {
                weekNum -= 1;
            }
            var result = firstThursday.AddDays(weekNum * 7);
            return result.AddDays(-3).AddDays(daysToAdd);
        }

    }

    public class Result
    {
        public string query { get; set; }
        public List<Intent> intents { get; set; }
        public List<Entity> entities { get; set; }


        public string RetrieveIntention()
        {
            double max = -1.00;
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

        public Entity RetrieveEntity(EntityTypeNames entityType)
        {
            double max = -1.00;
            Entity bestEntity = null;
            foreach (var entity in this.entities.Where(e => e.type == entityType.EntityTypeName && e.score >= entityType.EntityThreashold))
            {
                if (max < entity.score)
                {
                    bestEntity = entity;
                    max = entity.score;
                }
            }
            if (bestEntity != null && bestEntity.type == EntityTypeNames.EntityType.EntityTypeName)
            {
                EntityMetadata entity = CrmHelper.RetrieveMetadata().FirstOrDefault(e => e.LogicalName.ToLower() == bestEntity.entity.Replace(" ", "").ToLower());
                if (entity == null)
                {
                    //Retrieve by display name
                    entity = CrmHelper.RetrieveMetadata().FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == bestEntity.entity.ToLower());
                    if (entity == null)
                    {
                        //Retrieve by plural display name
                        entity = CrmHelper.RetrieveMetadata().FirstOrDefault(e => e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == bestEntity.entity.ToLower());
                    }
                }
                if (entity != null)
                {
                    bestEntity.entity = entity.LogicalName;
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