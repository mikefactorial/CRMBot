﻿using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Connector;
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
        public static EntityTypeNames AttributeName = new EntityTypeNames() { EntityTypeName = "AttributeName", EntityThreashold = .1 };
        public static EntityTypeNames AttributeValue = new EntityTypeNames() { EntityTypeName = "AttributeValue", EntityThreashold = .1 };
        public static EntityTypeNames DateTime = new EntityTypeNames() { EntityTypeName = "builtin.datetime.date", EntityThreashold = 0 };
        public static EntityTypeNames Ordinal = new EntityTypeNames() { EntityTypeName = "builtin.ordinal", EntityThreashold = 0 };
        public static EntityTypeNames DisplayField = new EntityTypeNames() { EntityTypeName = "DisplayField", EntityThreashold = .1 };

        public static EntityTypeNames FirstName = new EntityTypeNames() { EntityTypeName = "ContactName::FirstName", EntityThreashold = 0 };
        public static EntityTypeNames LastName = new EntityTypeNames() { EntityTypeName = "ContactName::LastName", EntityThreashold = 0 };
    }

    public static class LuisResultExtensions
    {
        public static EntityRecommendation RetrieveEntity(this LuisResult result, string channelId, string userId, EntityTypeNames entityType)
        {
            double? max = -1.00;
            EntityRecommendation bestEntity = null;
            foreach (var entity in result.Entities.Where(e => e.Type == entityType.EntityTypeName && (e.Score >= entityType.EntityThreashold) || (e.Score == null && entityType.EntityThreashold <= 0)))
            {
                if ((entity.Score != null && max < entity.Score) || (max <= 0 && entity.Score == null))
                {
                    bestEntity = entity;
                    max = entity.Score;
                }
            }
            if (bestEntity != null)
            {
                if (bestEntity.Type == EntityTypeNames.FirstName.EntityTypeName || bestEntity.Type == EntityTypeNames.LastName.EntityTypeName)
                {
                    bestEntity.Entity = bestEntity.Entity.Replace("'s", string.Empty);
                }
                else if (bestEntity.Type == EntityTypeNames.AttributeValue.EntityTypeName)
                {
                    if (bestEntity.Entity.Contains("@"))
                    {
                        bestEntity.Entity = bestEntity.Entity.Replace(" ", string.Empty);
                    }
                    bestEntity.Entity = bestEntity.Entity.Replace("'s", string.Empty).Replace(" '", string.Empty).Replace("' ", string.Empty);
                }
                else if (bestEntity.Type == EntityTypeNames.EntityType.EntityTypeName)
                {
                    bestEntity.Entity = CrmHelper.FindEntityLogicalName(channelId, userId, bestEntity.Entity);
                }
                if (bestEntity.Type.Contains(":"))
                {
                    bestEntity.Type = bestEntity.Type.Split(':')[bestEntity.Type.Split(':').Length - 1];
                }
            }
            return bestEntity;
        }

        public static List<DateTime> ParseDateTimes(this EntityRecommendation dateEntity)
        {
            List<DateTime> ret = new List<DateTime>();
            foreach (var vals in dateEntity.Resolution.Values)
            {
                Dictionary<string, object> values = (Dictionary<string, object>)((List<object>)vals)[0];
                if (values["type"].ToString() == "daterange")
                {
                    DateTime start;
                    DateTime end;

                    if (values.ContainsKey("start") && DateTime.TryParse(values["start"].ToString(), out start))
                    {
                        ret.Add(start);
                    }
                    if (values.ContainsKey("end") && DateTime.TryParse(values["end"].ToString(), out end))
                    {
                        ret.Add(end);
                    }
                }
            }
            return ret;
        }

        public static DateTime GetNextWeekday(DateTime start, DayOfWeek day)
        {
            // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
            int daysToAdd = ((int)day - (int)start.DayOfWeek + 7) % 7;
            return start.AddDays(daysToAdd);
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