using DmsService.Repository;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace DmsService.Implementation
{
    public class DmsWebService
    {
        private IOrganizationService _service;
        private CrmRepository<dot_dotcysetting> _dotcySettingRepository;
        private CrmRepository<dot_documents> _documentRepository;
        private CrmRepository<Contact> _contactRepository;
        private CrmRepository<Lead> _leadRepository;
        private CrmRepository<Account> _accountRepository;
        private CrmRepository<SystemUser> _systemUserRepository;
        private CrmRepository<Annotation> _annotationRepository;
        private CrmRepository<dot_documenttype> _documentTypeRepository;

        

        bool authorization = false;
        string dmsUrl = string.Empty;
        string userName = string.Empty;
        string password = string.Empty;
        string applicationId = string.Empty;

        public string error = string.Empty;

        List<Annotation> annotations = new List<Annotation>();        

        public DmsWebService(IOrganizationService service)
        {
            _service = service;
            _dotcySettingRepository = new CrmRepository<dot_dotcysetting>(_service);
            _documentRepository = new CrmRepository<dot_documents>(_service);
            _contactRepository = new CrmRepository<Contact>(_service);
            _leadRepository = new CrmRepository<Lead>(_service);
            _accountRepository = new CrmRepository<Account>(_service);
            _systemUserRepository = new CrmRepository<SystemUser>(_service);
            _annotationRepository = new CrmRepository<Annotation>(_service);
            _documentTypeRepository = new CrmRepository<dot_documenttype>(_service);

            authorization = CheckIsAuthorization();
        }  

        private bool CheckIsAuthorization()
        {
            var dotcySetting = _dotcySettingRepository.GetAll().FirstOrDefault();
            if (dotcySetting != null && dotcySetting.dot_dmsUrl != null && dotcySetting.dot_dmsLogin != null && dotcySetting.dot_dmsPassword != null &&
                dotcySetting.dot_dmsUrl != string.Empty && dotcySetting.dot_dmsLogin != string.Empty && dotcySetting.dot_dmsPassword != string.Empty &&
                dotcySetting.dot_dmsApplicationId != null && dotcySetting.dot_dmsApplicationId != string.Empty)
            {
                dmsUrl = dotcySetting.dot_dmsUrl;
                userName = dotcySetting.dot_dmsLogin;
                password = dotcySetting.dot_dmsPassword;
                applicationId = dotcySetting.dot_dmsApplicationId;
                return true;
            }
            else
            {
                error = "Authorization Parameters Failed";
                return false;
            }            
        }

        public bool CheckRequiredParametersForIntegration(Entity target, out JsonMetadataMapping jsonMapping, out Entity fetchResult)
        {
            dot_documents document = new dot_documents();
            jsonMapping = new JsonMetadataMapping();
            fetchResult = null;
            bool isExistMetadada = false;
            if (target.LogicalName == "annotation")
            {
                document = _documentRepository.GetCrmEntityById(target.GetAttributeValue<EntityReference>("objectid").Id);
                isExistMetadada = CheckIsDocumentHaveMetadata(document.Id, out jsonMapping, out fetchResult);
            }
            else if (target.LogicalName == "dot_documents")
            {               
                annotations = _annotationRepository.GetEntitiesByField("objectid", target.Id, new ColumnSet(true)).ToList();
                isExistMetadada = CheckIsDocumentHaveMetadata(target.Id, out jsonMapping, out fetchResult);
            }            
            if (authorization && isExistMetadada && ((document != null && document.dot_Validated != null && document.dot_Validated.Value == true) || (annotations != null && annotations.Count > 0)))
            {               
                return true;
            }
       
            return false; 
        }
        public bool CheckRequiredParametersForRetrieve(Entity target, out JsonMetadataMapping jsonMapping, out Entity fetchResult)
        {
            jsonMapping = new JsonMetadataMapping();
            fetchResult = null;

            if (authorization)
            {
                string fetch = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
                                  "<entity name='dot_documents'>" +
                                    "<attribute name='dot_documentsid' />" +
                                    "<attribute name='dot_name' />" +
                                    "<attribute name='createdon' />" +
                                    "<attribute name='dot_validated' />" +
                                    "<order attribute='dot_name' descending='false' />" +
                                    "<link-entity name='annotation' from='objectid' to='dot_documentsid' alias='ag'>" +
                                      "<filter type='and'>" +
                                        "<condition attribute='annotationid' operator='eq' uitype='annotation' value='" + target.Id + "' />" +
                                      "</filter>" +
                                    "</link-entity>" +
                                  "</entity>" +
                                "</fetch> ";

                EntityCollection result = _service.RetrieveMultiple(new FetchExpression(fetch));

                if (result != null)
                {
                    if (result.Entities[0].Contains("dot_validated") && result.Entities[0].GetAttributeValue<bool>("dot_validated") == true)
                    {
                        var isMetadataExist = CheckIsDocumentHaveMetadata(result.Entities[0].Id, out jsonMapping, out fetchResult);                                       
                        return isMetadataExist;
                    }
                }
            }
            return false;           
        }

        public bool CheckIsDocumentHaveMetadata(Guid documentId, out JsonMetadataMapping jsonMapping, out Entity fetchResult)
        {
            jsonMapping = new JsonMetadataMapping();
            fetchResult = null;
            var document = _documentRepository.GetCrmEntityById(documentId);
            if(document != null && document.dot_DocumentTypeid != null)
            {
                var documentType = _documentTypeRepository.GetCrmEntityById(document.dot_DocumentTypeid.Id);
                if (documentType != null && documentType.dot_dmsintegrationquery != null && documentType.dot_dmsintegrationquery != string.Empty &&
                    documentType.dot_dmsintegrationmapping != null && documentType.dot_dmsintegrationmapping != string.Empty &&
                    documentType.dot_dmsdocumenttype != null && documentType.dot_dmsdocumenttype != string.Empty)
                {
                    var jsonMappingString = documentType.dot_dmsintegrationmapping;
                    try
                    {
                        using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(jsonMappingString)))
                        {
                            DataContractJsonSerializer serialiser = new DataContractJsonSerializer(typeof(RootObject));
                            jsonMapping.JsonParameters = (RootObject)serialiser.ReadObject(ms);
                            jsonMapping.DocumentTypeId = documentType.dot_dmsdocumenttype;
                        }
                    }
                    catch (Exception ex)
                    {
                        error = "Can't read dms integration mapping";
                        throw new InvalidPluginExecutionException("Can't read dms integration mapping");
                    }
                    if (jsonMapping != null)
                    {
                        try
                        {
                            var fetch = string.Format(documentType.dot_dmsintegrationquery, documentId);
                            fetchResult = _service.RetrieveMultiple(new FetchExpression(fetch)).Entities.FirstOrDefault();
                            if (fetchResult == null)
                            {
                                error = "Fetch for metadata don't given results";
                                return false;
                            }
                            else return true;
                        }
                        catch (Exception ex)
                        {
                            error = "Fetch for metadata broken";
                            return false;                            
                        }
                    }                   
                }
                else
                {
                    error = "One of required field in Document type is empty";
                    return false;
                }
            }
            return false;
        }

        #region XmlRequest
        private string XmlRequestForUploadOrUpdateDocuments(string method, Entity annotation, string documentDMSId, JsonMetadataMapping jsonMapping, Entity fetchResult)
        {
            XmlDocument xmlForUpload = new XmlDocument();
            //XmlDeclaration xmldecl;
            //xmldecl = xmlForUpload.CreateXmlDeclaration("1.0", null, null);
            //xmldecl.Encoding = "windows-1252";           

            XmlElement el = (XmlElement)xmlForUpload.AppendChild(xmlForUpload.CreateElement(method));
            el.AppendChild(xmlForUpload.CreateElement("documentBinary")).InnerText = annotation.GetAttributeValue<string>("documentbody");
            if (method == "UploadDocument") el.AppendChild(xmlForUpload.CreateElement("fileName")).InnerText = annotation.GetAttributeValue<string>("filename");
            if (method == "UpdateDocument") el.AppendChild(xmlForUpload.CreateElement("DocumentId")).InnerText = documentDMSId;
            el.AppendChild(xmlForUpload.CreateElement("UserName")).InnerText = userName;
            el.AppendChild(xmlForUpload.CreateElement("Password")).InnerText = password;
            el.AppendChild(xmlForUpload.CreateElement("ApplicationId")).InnerText = applicationId;
            XmlElement el2 = xmlForUpload.CreateElement("metadata");
            el.AppendChild(el2);
            XmlElement el3 = xmlForUpload.CreateElement("DocumentMetadata");
            el2.AppendChild(el3);
            el3.AppendChild(xmlForUpload.CreateElement("Id")).InnerText = "Document Type";
            el3.AppendChild(xmlForUpload.CreateElement("Value")).InnerText = jsonMapping.DocumentTypeId;

            if (method != "UpdateDocument")
            {
                foreach (var obj in jsonMapping.JsonParameters.parametersData)
                {
                    if (fetchResult.Contains(obj.parameterName))
                    {
                        var xmlElem = xmlForUpload.CreateElement("DocumentMetadata");
                        el2.AppendChild(xmlElem);
                        xmlElem.AppendChild(xmlForUpload.CreateElement("Id")).InnerText = obj.integration_name;
                        var value = string.Empty;
                        switch (obj.type)
                        {
                            case ("date"):
                                value = DateTime.Parse(fetchResult.GetAttributeValue<AliasedValue>(obj.parameterName).Value.ToString()).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                break;
                            case ("optionsetlabel"):
                                value = GetOptionSetStringValue(fetchResult.GetAttributeValue<AliasedValue>(obj.parameterName));
                                break;
                            default:
                                value = fetchResult.GetAttributeValue<AliasedValue>(obj.parameterName).Value.ToString();
                                break;
                        };
                        xmlElem.AppendChild(xmlForUpload.CreateElement("Value")).InnerText = value;                            
                    }
                }
            }               

            //xmlForUpload.InsertBefore(xmldecl, el);

            return xmlForUpload.OuterXml;
        }
        private string XmlRequestForDownloadDocuments(string documentId)
        {
            XmlDocument xmlForDownload = new XmlDocument();
            XmlElement el = (XmlElement)xmlForDownload.AppendChild(xmlForDownload.CreateElement("DownloadDocument"));
            el.AppendChild(xmlForDownload.CreateElement("UserName")).InnerText = userName;
            el.AppendChild(xmlForDownload.CreateElement("Password")).InnerText = password;
            el.AppendChild(xmlForDownload.CreateElement("ApplicationId")).InnerText = applicationId;
            el.AppendChild(xmlForDownload.CreateElement("DocumentId")).InnerText = documentId;

            return xmlForDownload.OuterXml;
        }
        #endregion XmlRequest

        public string DmsServiceUploadJobe(Entity anntotation, JsonMetadataMapping jsonMapping, Entity fetchResult)
        {
            string xmlRequestForUpload = XmlRequestForUploadOrUpdateDocuments("UploadDocument", anntotation, string.Empty, jsonMapping, fetchResult);
            return DMSServicePostRequest(xmlRequestForUpload, "UploadDocument");            
        }
        public string DmsServiceUpdateJobe(Entity anntotation, string documentDMSId, JsonMetadataMapping jsonMapping, Entity fetchResult)
        {            
            string xmlRequestForUpload = XmlRequestForUploadOrUpdateDocuments("UpdateDocument", anntotation, documentDMSId, jsonMapping, fetchResult);
            //throw new InvalidPluginExecutionException(xmlRequestForUpload);
            return DMSServicePostRequest(xmlRequestForUpload, "UpdateDocument");
        }
        public void DmsServiceUploadAnntotations(JsonMetadataMapping jsonMapping, Entity fetchResult)
        {
            foreach(var annotation in annotations)
            {
                var documentId = DMSServicePostRequest(XmlRequestForUploadOrUpdateDocuments("UploadDocument", annotation, string.Empty, jsonMapping, fetchResult), "UploadDocument");                
                if(documentId != string.Empty)
                {                   
                    var uploadDocumentIdBytes = Encoding.UTF8.GetBytes(documentId);
                    Annotation annotationForApdate = new Annotation()
                    {
                        Id = annotation.Id,                                            
                        DocumentBody = Convert.ToBase64String(uploadDocumentIdBytes), 
                    };
                    _annotationRepository.Update(annotationForApdate);
                }
            }            
        }
        public string DmsServiceDownloadJobe(Entity annotation)
        {
            var documetnId = annotation.GetAttributeValue<string>("documentbody");
            var base64EncodedBytes = Convert.FromBase64String(documetnId);

            string xmlRequestForDownload = XmlRequestForDownloadDocuments(Encoding.UTF8.GetString(base64EncodedBytes));
            var downloadDocument = DMSServicePostRequest(xmlRequestForDownload, "DownloadDocument");            
            return downloadDocument;
        }
              
        private string DMSServicePostRequest(string xmlRequest, string path)
        {
            string postResponse = string.Empty;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(dmsUrl + path);
            byte[] bytes = Encoding.ASCII.GetBytes(xmlRequest);
            request.Headers.Add("cache-control", "no-cache");
            request.ContentType = "application/xml; encoding='utf-8'";
            request.ContentLength = bytes.Length;
            request.Method = "POST";
            if(path == "UploadDocument") request.Accept = "application/json, application/xml, text/json, text/x-json, text/javascript, text/xml";
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            requestStream.Close();
            HttpWebResponse response;
            response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream responseStream = response.GetResponseStream();
                string responseStr = new StreamReader(responseStream).ReadToEnd();
                postResponse = GetObjectFromResponce(responseStr);
            }
            return postResponse;
        }
        public void UpdateDocumentIntegrationStatus(Guid documentId)
        {
            var document = _documentRepository.GetCrmEntityById(documentId);
            var documentForUpdate = new dot_documents()
            {
                Id = document.Id,
                dot_IntegrationStatus = error != string.Empty ? new OptionSetValue(180000002) : new OptionSetValue(180000001),
                dot_IntegrationError = error,
            };
            _documentRepository.Update(documentForUpdate);
        }

        private string GetObjectFromResponce (string responseStr)
        {            
            string documentResponse = string.Empty;
            XDocument xdoc = XDocument.Parse(responseStr);
            XmlSerializer serializer = new XmlSerializer(typeof(UploadResponse));
            var response = (UploadResponse)serializer.Deserialize(new StringReader(xdoc.ToString()));
            if (response.Message.MessageCode == "2000") documentResponse = response.DocumentResponse;
            else error = response.Message.Description;
            return documentResponse;
        }
        private string GetOptionSetStringValue(AliasedValue optionSetParameter)
        {
            string optionSetLable = string.Empty;
            var response = (RetrieveAttributeResponse)_service.Execute(new RetrieveAttributeRequest()
            {
                EntityLogicalName = optionSetParameter.EntityLogicalName,
                LogicalName = optionSetParameter.AttributeLogicalName,
                RetrieveAsIfPublished = false

            });
            if(response != null)
            {
                dynamic attributeMetadata = null;
                var optionSetValue = (OptionSetValue)optionSetParameter.Value;
                if (response.AttributeMetadata is PicklistAttributeMetadata) attributeMetadata = (PicklistAttributeMetadata)response.AttributeMetadata;
                OptionMetadata[] optionList = attributeMetadata.OptionSet.Options.ToArray();
                optionSetLable = optionList.Where(o => o.Value == optionSetValue.Value).Select(o => o.Label.UserLocalizedLabel.Label).FirstOrDefault();
            }
            return optionSetLable;
        }
    }
}
