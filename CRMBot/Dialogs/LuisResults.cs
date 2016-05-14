using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;

namespace CRMBot.Dialogs
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

    public static class LuisResultExtensions
    {
        public static EntityRecommendation RetrieveEntity(this LuisResult result, EntityTypeNames entityType)
        {
            double? max = -1.00;
            EntityRecommendation bestEntity = null;
            foreach (var entity in result.Entities.Where(e => e.Type == entityType.EntityTypeName && e.Score >= entityType.EntityThreashold))
            {
                if (max < entity.Score)
                {
                    bestEntity = entity;
                    max = entity.Score;
                }
            }
            if (bestEntity != null && bestEntity.Type == EntityTypeNames.EntityType.EntityTypeName)
            {
                EntityMetadata entity = CrmHelper.RetrieveMetadata().FirstOrDefault(e => e.LogicalName.ToLower() == bestEntity.Entity.Replace(" ", "").ToLower());
                if (entity == null)
                {
                    //Retrieve by display name
                    entity = CrmHelper.RetrieveMetadata().FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == bestEntity.Entity.ToLower());
                    if (entity == null)
                    {
                        //Retrieve by plural display name
                        entity = CrmHelper.RetrieveMetadata().FirstOrDefault(e => e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == bestEntity.Entity.ToLower());
                    }
                }
                if (entity != null)
                {
                    bestEntity.Entity = entity.LogicalName;
                }
            }
            return bestEntity;
        }

        public static DateTime[] ParseDateTimes(this EntityRecommendation dateEntity)
        {
            /*
            if (dateEntity.Resolution != null && !string.IsNullOrEmpty(dateEntity.Resolution.date))
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
            }*/
            return null;
        }

        public static DateTime FirstDateOfWeekISO8601(int year, int weekOfYear, int daysToAdd)
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
}