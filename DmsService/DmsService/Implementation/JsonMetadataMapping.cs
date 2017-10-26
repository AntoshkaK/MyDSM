using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace DmsService.Implementation
{
    public class JsonMetadataMapping
    {
        public RootObject JsonParameters { get; set; }
        public string DocumentTypeId { get; set; }
    }

    [DataContract]
    public class ParametersData
    {
        [DataMember]
        public string integration_name { get; set; }
        [DataMember]
        public string parameterName { get; set; }
        [DataMember]
        public string type { get; set; }
    }
    [DataContract]
    public class RootObject
    {
        [DataMember]
        public List<ParametersData> parametersData { get; set; }
    }
}
