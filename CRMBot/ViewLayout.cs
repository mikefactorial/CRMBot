using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Serialization;

namespace CRMBot
{
    [XmlRoot(ElementName = "cell")]
    public class Cell
    {
        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }
        [XmlAttribute(AttributeName = "width")]
        public string Width { get; set; }
    }

    [XmlRoot(ElementName = "row")]
    public class Row
    {
        [XmlElement(ElementName = "cell")]
        public List<Cell> Cells { get; set; }
        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }
        [XmlAttribute(AttributeName = "id")]
        public string Id { get; set; }
    }

    [XmlRoot(ElementName = "grid")]
    public class Grid
    {
        [XmlElement(ElementName = "row")]
        public Row Row { get; set; }
        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }
        [XmlAttribute(AttributeName = "object")]
        public string Object { get; set; }
        [XmlAttribute(AttributeName = "jump")]
        public string Jump { get; set; }
        [XmlAttribute(AttributeName = "select")]
        public string Select { get; set; }
        [XmlAttribute(AttributeName = "icon")]
        public string Icon { get; set; }
        [XmlAttribute(AttributeName = "preview")]
        public string Preview { get; set; }

        public static Grid Deserialize(string xml)
        {
            // DeSerialize the XML into an Entity and return the Entity
            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(Grid));
            // Declare a CRM Entity
            System.IO.StringReader reader = new System.IO.StringReader(xml);
            // Deserialize the Entity object
            return serializer.Deserialize(reader) as Grid;

        }
    }
}