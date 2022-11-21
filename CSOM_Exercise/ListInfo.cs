using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

[Serializable]

[XmlRoot("List")]
public class ListInfo
{
    [XmlElement]
    public List<Item> item { get; set; }
    [XmlAttribute]
    public string name { get; set; }
    [XmlAttribute]
    public string url { get; set; }
    [XmlAttribute]
    public string type { get; set; }
}


[Serializable]
public class Item
{
    [XmlAttribute]
    public string name { get; set; }

    [XmlAttribute]
    public string Url { get; set; }

    [XmlAttribute]
    public string ID { get; set; }
    
    [XmlElement]
    public List<column> Column { get; set; }
}


[Serializable]
public class column
{
    [XmlAttribute]
    public string name { get; set; }

    [XmlAttribute]
    public string value { get; set; }
}