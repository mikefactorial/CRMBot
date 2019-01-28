using Microsoft.Bot.Builder.Luis.Models;
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
        public static EntityTypeNames DateTime = new EntityTypeNames() { EntityTypeName = "builtin.datetimeV2.daterange", EntityThreashold = 0 };
        public static EntityTypeNames Ordinal = new EntityTypeNames() { EntityTypeName = "builtin.ordinal", EntityThreashold = 0 };
        public static EntityTypeNames DisplayField = new EntityTypeNames() { EntityTypeName = "DisplayField", EntityThreashold = .1 };

        public static EntityTypeNames FirstName = new EntityTypeNames() { EntityTypeName = "ContactName::FirstName", EntityThreashold = 0 };
        public static EntityTypeNames LastName = new EntityTypeNames() { EntityTypeName = "ContactName::LastName", EntityThreashold = 0 };
    }

    public static class LuisResultExtensions
    {
        private const int MIN_TEXTLENGTHFORFIELDSEARCH = 4;
        private const int MIN_TEXTLENGTHFORENTITYSEARCH = 4;

        public static string FindAttributeLogicalName(this LuisResult result, EntityMetadata entity, string text)
        {
            string subText = text.ToLower();
            //Equals
            AttributeMetadata att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
            if (att != null)
            {
                return att.LogicalName;
            }
            att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
            if (att != null)
            {
                return att.LogicalName;
            }
            att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
            if (att != null)
            {
                return att.LogicalName;
            }
            att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
            if (att != null)
            {
                return att.LogicalName;
            }

            //Substring Equals

            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower() == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.LogicalName == subText);
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }

            //Substring Contains
            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.DisplayName != null && a.DisplayName.UserLocalizedLabel != null && a.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.LogicalName.Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(1);
            }

            subText = text.Replace(" ", "").ToLower();
            while (subText.Length >= MIN_TEXTLENGTHFORFIELDSEARCH)
            {
                att = entity.Attributes.FirstOrDefault(a => a.LogicalName.Contains(subText));
                if (att != null)
                {
                    return att.LogicalName;
                }
                subText = subText.Substring(0, subText.Length - 1);
            }


            return string.Empty;
        }

        public static string FindEntityLogicalName(this LuisResult result, string channelId, string userId, EntityRecommendation entityTypeEntity)
        {
            ChatState state = ChatState.RetrieveChatState(channelId, userId);
            string text = entityTypeEntity.Entity.ToLower();

            if (!state.EntityMapping.ContainsKey(text))
            {
                string subText = text;
                EntityMetadata[] metadata = state.RetrieveMetadata();
                //state.EntityMapping[text] = "contact"; //If can't find a match default to contact records
                //Equals
                EntityMetadata entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
                if (entity != null)
                {
                    state.EntityMapping[text] = entity.LogicalName;
                    return state.EntityMapping[text];
                }
                entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
                if (entity != null)
                {
                    state.EntityMapping[text] = entity.LogicalName;
                    return state.EntityMapping[text];
                }
                entity = metadata.FirstOrDefault(e => e.LogicalName == subText);
                if (entity != null)
                {
                    state.EntityMapping[text] = entity.LogicalName;
                    return state.EntityMapping[text];
                }
                //Substring Equals
                while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
                {
                    entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
                    if (entity != null)
                    {
                        state.EntityMapping[text] = entity.LogicalName;
                        return state.EntityMapping[text];
                    }
                    subText = subText.Substring(1);
                }

                subText = text.ToLower();
                while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
                {
                    entity = metadata.FirstOrDefault(e => (e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower() == subText) || (e.DisplayCollectionName != null && e.DisplayCollectionName.UserLocalizedLabel != null && e.DisplayCollectionName.UserLocalizedLabel.Label.ToLower() == subText));
                    if (entity != null)
                    {
                        state.EntityMapping[text] = entity.LogicalName;
                        return state.EntityMapping[text];
                    }
                    subText = subText.Substring(0, subText.Length - 1);
                }


                subText = text.Replace(" ", "").ToLower();
                while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
                {
                    entity = metadata.FirstOrDefault(e => e.LogicalName == subText);
                    if (entity != null)
                    {
                        state.EntityMapping[text] = entity.LogicalName;
                        return state.EntityMapping[text];
                    }
                    subText = subText.Substring(1);
                }

                subText = text.Replace(" ", "").ToLower();
                while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
                {
                    entity = metadata.FirstOrDefault(e => e.LogicalName == subText);
                    if (entity != null)
                    {
                        state.EntityMapping[text] = entity.LogicalName;
                        return state.EntityMapping[text];
                    }
                    subText = subText.Substring(0, subText.Length - 1);
                }

                //Contains
                subText = text.ToLower();
                while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
                {
                    entity = metadata.FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                    if (entity != null)
                    {
                        state.EntityMapping[text] = entity.LogicalName;
                        return state.EntityMapping[text];
                    }
                    subText = subText.Substring(1);
                }

                subText = text.ToLower();
                while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
                {
                    entity = metadata.FirstOrDefault(e => e.DisplayName != null && e.DisplayName.UserLocalizedLabel != null && e.DisplayName.UserLocalizedLabel.Label.ToLower().Contains(subText));
                    if (entity != null)
                    {
                        state.EntityMapping[text] = entity.LogicalName;
                        return state.EntityMapping[text];
                    }
                    subText = subText.Substring(0, subText.Length - 1);
                }


                subText = text.Replace(" ", "").ToLower();
                while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
                {
                    entity = metadata.FirstOrDefault(e => e.LogicalName.Contains(subText));
                    if (entity != null)
                    {
                        state.EntityMapping[text] = entity.LogicalName;
                        return state.EntityMapping[text];
                    }
                    subText = subText.Substring(1);
                }

                subText = text.Replace(" ", "").ToLower();
                while (subText.Length >= MIN_TEXTLENGTHFORENTITYSEARCH)
                {
                    entity = metadata.FirstOrDefault(e => e.LogicalName.Contains(subText));
                    if (entity != null)
                    {
                        state.EntityMapping[text] = entity.LogicalName;
                        return state.EntityMapping[text];
                    }
                    subText = subText.Substring(0, subText.Length - 1);
                }
            }
            return state.EntityMapping[text];
        }

        public static EntityRecommendation RetrieveEntity(this LuisResult result, string channelId, string userId, bool originalCasing, params EntityTypeNames[] entityTypes)
        {
            double? max = -1.00;
            EntityRecommendation bestEntity = null;
            for (int i = 0; i < entityTypes.Length; i++)
            {
                foreach (var entity in result.Entities.Where(e => e.Type == entityTypes[i].EntityTypeName && ((e.Score >= entityTypes[i].EntityThreashold) || (e.Score == null && entityTypes[i].EntityThreashold <= 0))))
                {
                    if ((entity.Score != null && max < entity.Score) || (max <= 0 && entity.Score == null))
                    {
                        bestEntity = entity;
                        max = entity.Score;
                    }
                }
                if (bestEntity != null)
                {
                    break;
                }
            }
            if (bestEntity != null)
            {
                if (originalCasing)
                {
                    bestEntity.Entity = result.Query.Substring((int)bestEntity.StartIndex, ((int)bestEntity.EndIndex - (int)bestEntity.StartIndex) + 1);
                }
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
                    bestEntity.Entity = bestEntity.Entity.Replace("'s", string.Empty).Replace(" '", string.Empty).Replace("' ", string.Empty).Replace(" - ", "-").Replace(" . ", ".").Replace(" / ", "/").Replace(" ,", ",");
                }
                else if (bestEntity.Type == EntityTypeNames.EntityType.EntityTypeName)
                {
                    bestEntity.Entity = result.FindEntityLogicalName(channelId, userId, bestEntity);
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
                else if (values.ContainsKey("value"))
                {
                    DateTime value;
                    if (DateTime.TryParse(values["value"].ToString(), out value))
                    {
                        ret.Add(value);
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