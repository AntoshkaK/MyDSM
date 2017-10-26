using DmsService.Implementation;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
         
            var connection = new CrmConnection("Crm");
            var _service = new OrganizationService(connection);            

            var dmsService = new DmsWebService(_service);
            JsonMetadataMapping jsonMapping = null;
            Entity fetchResult = null;

            var document = _service.Retrieve(dot_documents.EntityLogicalName, new Guid("308DF494-0545-E711-80C9-005056B90C62"), new ColumnSet(true));

            if (dmsService.CheckRequiredParametersForIntegration(document, out jsonMapping, out fetchResult))
            {
                dmsService.DmsServiceUploadAnntotations(jsonMapping, fetchResult);
            }


            //3D4E35CA-E692-E711-80CD-005056B90C62     
            var anntotation2 = _service.Retrieve(Annotation.EntityLogicalName, new Guid("9577BE91-8F93-E711-80CD-005056B90C62"), new ColumnSet(true));
            var anntotation1 = _service.Retrieve(Annotation.EntityLogicalName, new Guid("A9EBAA09-8E93-E711-80CD-005056B90C62"), new ColumnSet(true));

            //JsonMetadataMapping jsonMapping = new JsonMetadataMapping();
            //Entity fetchResult = null;
            //dmsService.CheckRequiredParametersForRetrieve(anntotation1, out jsonMapping, out fetchResult);

            //var discription = anntotation1.GetAttributeValue<string>("notetext");
            //if (discription != null && discription.Split('^').Length == 2)
            //{
            //    var documentDMSIdBase64EncodedBytes = Convert.FromBase64String(discription.Split('^')[1]);
            //    var documentDMSId = Encoding.UTF8.GetString(documentDMSIdBase64EncodedBytes);
            //    dmsService.DmsServiceUpdateJobe(anntotation2, documentDMSId, jsonMapping, fetchResult);             
            //}         


            // string documentBody = dmsService.DmsServiceDownloadJobe(anntotation);

        }
    }
}
