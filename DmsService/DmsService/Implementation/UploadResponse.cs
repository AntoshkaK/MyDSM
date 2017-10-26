using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace DmsService.Implementation
{
    [XmlRoot(ElementName = "Message")]
    public class Message
    {
        [XmlElement(ElementName = "MessageId")]
        public string MessageId { get; set; }
        [XmlElement(ElementName = "MessageCode")]
        public string MessageCode { get; set; }
        [XmlElement(ElementName = "Description")]
        public string Description { get; set; }
    }

    [XmlRoot(ElementName = "Response")]
    public class UploadResponse
    {
        [XmlElement(ElementName = "DocumentResponse")]
        public string DocumentResponse { get; set; }
        [XmlElement(ElementName = "Message")]
        public Message Message { get; set; }
    }

}