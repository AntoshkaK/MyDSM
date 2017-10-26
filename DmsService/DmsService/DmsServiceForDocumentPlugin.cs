using DmsService.Implementation;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DmsService
{
    public class DmsServiceForDocumentPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationService _service = factory.CreateOrganizationService(context.UserId);
           
            if (context.MessageName == "Update" && context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity document = (Entity)context.InputParameters["Target"];
                if (document.Contains("dot_validated") && document.GetAttributeValue<bool>("dot_validated") == true)
                {                   
                    var dmsService = new DmsWebService(_service);
                    JsonMetadataMapping jsonMapping = new JsonMetadataMapping();                 
                    Entity fetchResult = null;
                    if (dmsService.CheckRequiredParametersForIntegration(document, out jsonMapping, out fetchResult))
                    {                                    
                        dmsService.DmsServiceUploadAnntotations(jsonMapping, fetchResult);
                    }
                    if (dmsService.error != string.Empty)
                    {
                        document["dot_integrationerror"] = dmsService.error;
                        document["dot_integrationstatus"] = new OptionSetValue(180000002);
                    }
                    else
                    {
                        document["dot_integrationstatus"] = new OptionSetValue(180000001);
                        document["dot_integrationerror"] = string.Empty;
                    }
                }
            }
        }
    }
}