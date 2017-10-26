using DmsService.Implementation;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DmsService
{
    public class DmsServiceForAnnotationPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));         
            IOrganizationService _service = factory.CreateOrganizationService(context.UserId);

            var dmsService = new DmsWebService(_service);

            if (context.MessageName == "Create" && context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity document = (Entity)context.InputParameters["Target"];

                if (document.Contains("objectid") && document.GetAttributeValue<EntityReference>("objectid") != null && document.GetAttributeValue<EntityReference>("objectid").LogicalName == "dot_documents")
                {
                    JsonMetadataMapping jsonMapping = new JsonMetadataMapping();
                    Entity fetchResult = null;                                 
                    if (dmsService.CheckRequiredParametersForIntegration(document, out jsonMapping, out fetchResult))
                    {
                        string uploadDocumentId = dmsService.DmsServiceUploadJobe(document, jsonMapping, fetchResult);
                        if (uploadDocumentId != string.Empty)
                        {
                            var uploadDocumentIdBytes = Encoding.UTF8.GetBytes(uploadDocumentId);
                            document["documentbody"] = Convert.ToBase64String(uploadDocumentIdBytes);
                        }
                    }
                    dmsService.UpdateDocumentIntegrationStatus(document.GetAttributeValue<EntityReference>("objectid").Id);                   
                }
            }
            else if (context.MessageName == "Update" && context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PreEntityImages.Contains("PreImage") && context.PreEntityImages["PreImage"] is Entity)
            {                          
                Entity document = (Entity)context.InputParameters["Target"];
                var preMessageImage = (Entity)context.PreEntityImages["PreImage"];

                JsonMetadataMapping jsonMapping = new JsonMetadataMapping();
                Entity fetchResult = null;

                if (preMessageImage.Contains("objectid") && preMessageImage.GetAttributeValue<EntityReference>("objectid") != null && preMessageImage.GetAttributeValue<EntityReference>("objectid").LogicalName == "dot_documents" &&
                    document.Contains("documentbody") && document.GetAttributeValue<string>("documentbody") != null && dmsService.CheckRequiredParametersForRetrieve(preMessageImage, out jsonMapping, out fetchResult))   
                {
                    if (preMessageImage.Contains("documentbody") && preMessageImage.GetAttributeValue<string>("documentbody") != null && preMessageImage.GetAttributeValue<string>("documentbody") != string.Empty)
                    {
                        if (document.GetAttributeValue<string>("documentbody") != string.Empty)
                        {
                            var base64EncodedBytes = Convert.FromBase64String(document.GetAttributeValue<string>("documentbody"));
                            var docId = Encoding.UTF8.GetString(base64EncodedBytes);
                            bool checkIsGuid = false;
                            try
                            {
                                new Guid(docId);
                            }
                            catch (Exception ex)
                            {
                                checkIsGuid = true;
                            }

                            if (checkIsGuid)
                            {
                                var documentDMSIdBase64EncodedBytes = Convert.FromBase64String(preMessageImage.GetAttributeValue<string>("documentbody"));
                                var documentDMSId = Encoding.UTF8.GetString(documentDMSIdBase64EncodedBytes);
                                dmsService.DmsServiceUpdateJobe(document, documentDMSId, jsonMapping, fetchResult);
                                document["documentbody"] = preMessageImage.GetAttributeValue<string>("documentbody");
                            }
                            dmsService.UpdateDocumentIntegrationStatus(preMessageImage.GetAttributeValue<EntityReference>("objectid").Id);
                        }
                        else
                        {
                            var discription = preMessageImage.GetAttributeValue<string>("notetext");
                            document["notetext"] = string.Format("{0}^{1}", discription, preMessageImage.GetAttributeValue<string>("documentbody"));
                        }
                    }
                    else
                    {
                        var discription = preMessageImage.GetAttributeValue<string>("notetext");
                        if (discription != null && discription.Split('^').Length == 2)
                        {                            
                            var documentDMSIdBase64EncodedBytes = Convert.FromBase64String(discription.Split('^')[1]);
                            var documentDMSId = Encoding.UTF8.GetString(documentDMSIdBase64EncodedBytes);
                            dmsService.DmsServiceUpdateJobe(document, documentDMSId, jsonMapping, fetchResult);
                            document["documentbody"] = discription.Split('^')[1];
                            document["notetext"] = discription.Split('^')[0];
                        }
                        dmsService.UpdateDocumentIntegrationStatus(preMessageImage.GetAttributeValue<EntityReference>("objectid").Id);
                    }                    
                }
            }

            else if (context.MessageName == "Retrieve" && context.OutputParameters.Contains("BusinessEntity") && context.OutputParameters["BusinessEntity"] is Entity)
            {               
                var document = (Entity)context.OutputParameters["BusinessEntity"];
                JsonMetadataMapping jsonMapping = new JsonMetadataMapping();
                Entity fetchResult = null;
                if (document.GetAttributeValue<string>("documentbody") != null && dmsService.CheckRequiredParametersForRetrieve(document, out jsonMapping, out fetchResult))
                {
                    string documentBody = dmsService.DmsServiceDownloadJobe(document);
                    
                    if (documentBody != string.Empty)
                    {
                        document["documentbody"] = documentBody;
                    }
                }
                //dmsService.UpdateDocumentIntegrationStatus(document.GetAttributeValue<EntityReference>("objectid").Id);
            }
        }
    }
}