using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DmsService.Repository
{
    public class CrmRepository<T> where T : Entity, new()
    {
        private readonly IOrganizationService _service;
        private readonly string _entityName;

        public CrmRepository(IOrganizationService service)
        {
            _service = service;
            _entityName = new T().LogicalName;
        }

        protected string EntityName
        {
            get
            {
                return _entityName;
            }
        }

        public T GetCrmEntityById(Guid id)
        {
            return (T)_service.Retrieve(_entityName, id, new ColumnSet(true)).ToEntity<T>();
        }

        public T GetCrmEntityById(Guid id, ColumnSet columns)
        {
            return (T)_service.Retrieve(_entityName, id, columns).ToEntity<T>();
        }

        public EntityReference GetCrmEntityRefByAttribute(string attribute, object value)
        {
            Entity result = null;

            if (value != null)
            {
                var query = new QueryByAttribute(_entityName)
                {
                    ColumnSet = new ColumnSet(),
                    Attributes = { attribute },
                    Values = { value }
                };

                result = _service.RetrieveMultiple(query).Entities.FirstOrDefault();
            }

            return result != null ? result.ToEntityReference() : null;
        }

        public List<T> GetEntitiesByField(string field, object value, ColumnSet columns)
        {
            var queryByAttribute = new QueryByAttribute(_entityName)
            {
                ColumnSet = columns,
                Attributes = { field },
                Values = { value }
            };

            return _service.RetrieveMultiple(queryByAttribute).Entities.Select(e => e.ToEntity<T>()).ToList();
        }

        public IList<T> GetAll()
        {
            var query = new QueryExpression(_entityName)
            {
                ColumnSet = new ColumnSet(true)
            };

            return _service.RetrieveMultiple(query).Entities.Select(e => e.ToEntity<T>()).ToList();
        }

        public IList<T> GetAll(ColumnSet columnSet)
        {
            var query = new QueryExpression(_entityName)
            {
                ColumnSet = columnSet
            };

            return _service.RetrieveMultiple(query).Entities.Select(e => e.ToEntity<T>()).ToList();
        }

        public IList<T> GetEntitiesByField(string field, object value, ColumnSet columns, OrderExpression order)
        {
            var queryByAttribute = new QueryByAttribute(_entityName)
            {
                ColumnSet = columns,
                Attributes = { field },
                Values = { value },
                Orders = { order }
            };

            return _service.RetrieveMultiple(queryByAttribute).Entities.Select(e => e.ToEntity<T>()).ToList();
        }

        private IList<Entity> GetEntitiesByField(string linkName, string field, object value, ColumnSet columns)
        {
            var queryByAttribute = new QueryByAttribute(linkName)
            {
                ColumnSet = columns,
                Attributes = { field },
                Values = { value }
            };

            return _service.RetrieveMultiple(queryByAttribute).Entities;
        }

        public IList<T> GetEntitiesByFieldActive(string field, object value, ColumnSet columns, OrderExpression order)
        {
            var queryByAttribute = new QueryByAttribute(_entityName)
            {
                ColumnSet = columns,
                Attributes = { field, "statecode" },
                Values = { value, 0 },
                Orders = { order }
            };

            return _service.RetrieveMultiple(queryByAttribute).Entities.Select(e => e.ToEntity<T>()).ToList();
        }

        public IList<T> GetEntitiesByQuery(QueryBase query)
        {
            return _service.RetrieveMultiple(query).Entities.Select(e => e.ToEntity<T>()).ToList();
        }

        public int GetCountByFilterExpression(string filterExpression)
        {
            string fetchQuery =
                @"<fetch distinct='false' mapping='logical' aggregate='true'>" +
                    "<entity name='" + _entityName + "'> " +
                        "<attribute name='" + _entityName + "id' aggregate='count' alias='countAttribute'/>" +
                        filterExpression +
                    "</entity>" +
                 "</fetch>";

            return (int)((AliasedValue)_service.RetrieveMultiple(new FetchExpression(fetchQuery)).Entities[0]["countAttribute"]).Value;
        }

        protected OrganizationResponse Execute(OrganizationRequest query)
        {
            return _service.Execute(query);
        }

        public void Update(T entity)
        {
            _service.Update(entity.ToEntity<Entity>());
        }

        public void Delete(T entity)
        {
            _service.Delete(entity.LogicalName, entity.Id);
        }

        public Guid Create(T entity)
        {
            return _service.Create(entity.ToEntity<Entity>());
        }

        public void SetStatus(EntityReference entityRef, int state, int status = -1)
        {
            _service.Execute(new SetStateRequest()
            {
                EntityMoniker = entityRef,
                State = new OptionSetValue(state),
                Status = new OptionSetValue(status)
            });
        }

        public IList<Guid> GetLinkedEntities(string linkName, string attributefrom, string attributeto, Guid id)
        {
            return GetEntitiesByField(linkName, attributeto, id, new ColumnSet(attributefrom)).
                Select(link => (Guid)link[attributefrom]).ToList();
        }

        public void Disassociate(string relationship, EntityReference target, params EntityReference[] entities)
        {
            _service.Execute(new DisassociateRequest()
            {
                Target = target,
                RelatedEntities = new EntityReferenceCollection(entities),
                Relationship = new Relationship(relationship)
            });
        }

        public void Associate(string relationship, EntityReference target, params EntityReference[] entities)
        {
            _service.Execute(new AssociateRequest()
            {
                Target = target,
                RelatedEntities = new EntityReferenceCollection(entities),
                Relationship = new Relationship(relationship)
            });
        }

        public IEnumerable<PrincipalAccess> GetSharedPrincipals(EntityReference target)
        {
            var req = new RetrieveSharedPrincipalsAndAccessRequest()
            {
                Target = target
            };
            var resp = (RetrieveSharedPrincipalsAndAccessResponse)this.Execute(req);
            return resp.PrincipalAccesses;
        }

        public void ModifyAccess(EntityReference target, PrincipalAccess access)
        {
            var req = new ModifyAccessRequest()
            {
                Target = target,
                PrincipalAccess = access
            };
            this.Execute(req);
        }

        public void RevokeAccess(EntityReference target, EntityReference revokee)
        {
            var req = new RevokeAccessRequest()
            {
                Target = target,
                Revokee = revokee
            };
            this.Execute(req);
        }

        public void Assign(EntityReference target, EntityReference assignee)
        {
            var req = new AssignRequest()
            {
                Target = target,
                Assignee = assignee
            };
            this.Execute(req);
        }

        public string GetOptionsetText(string optionsetName, int optionsetValue)
        {
            string optionsetSelectedText = string.Empty;

            RetrieveOptionSetRequest retrieveOptionSetRequest =
                new RetrieveOptionSetRequest
                {
                    Name = optionsetName
                };

            // Execute the request.
            RetrieveOptionSetResponse retrieveOptionSetResponse = null;
            retrieveOptionSetResponse = (RetrieveOptionSetResponse)_service.Execute(retrieveOptionSetRequest);

            // Access the retrieved OptionSetMetadata.
            OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

            // Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
            foreach (OptionMetadata optionMetadata in optionList)
            {
                if (optionMetadata.Value == optionsetValue)
                {
                    optionsetSelectedText = optionMetadata.Label.UserLocalizedLabel.Label.ToString();
                    break;
                }
            }

            return optionsetSelectedText;
        }

        public OptionMetadata[] GetAllOptionsetValues(string optionsetName)
        {
            string optionsetSelectedText = string.Empty;

            RetrieveOptionSetRequest retrieveOptionSetRequest =
                new RetrieveOptionSetRequest
                {
                    Name = optionsetName
                };

            // Execute the request.
            RetrieveOptionSetResponse retrieveOptionSetResponse = null;
            retrieveOptionSetResponse = (RetrieveOptionSetResponse)_service.Execute(retrieveOptionSetRequest);

            // Access the retrieved OptionSetMetadata.
            OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

            // Get the current options list for the retrieved attribute.
            return retrievedOptionSetMetadata.Options.ToArray();
        }

        public void MultiUpdate(List<T> entityList)
        {
            ExecuteMultipleRequest request = new ExecuteMultipleRequest()
            {
                Requests = new OrganizationRequestCollection(),
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = true,
                    ReturnResponses = true
                }
            };

            foreach (var entity in entityList)
            {
                UpdateRequest updateRequest = new UpdateRequest { Target = entity.ToEntity<Entity>() };
                request.Requests.Add(updateRequest);
            }

            try
            {
                ExecuteMultipleResponse response = (ExecuteMultipleResponse)_service.Execute(request);
            }
            catch (Exception ex)
            { }
        }

        public void MultiDelete(List<T> Entities)
        {
            ExecuteMultipleRequest request = new ExecuteMultipleRequest()
            {
                Requests = new OrganizationRequestCollection(),
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = true,
                    ReturnResponses = true
                }
            };

            foreach (var entity in Entities)
            {
                DeleteRequest deleteRequest = new DeleteRequest { Target = new EntityReference(entity.LogicalName, entity.Id) };
                request.Requests.Add(deleteRequest);
            }

            ExecuteMultipleResponse response = (ExecuteMultipleResponse)_service.Execute(request);
        }

        public void MultiCreate(List<T> Entities)
        {
            var multipleRequest = new ExecuteMultipleRequest()
            {
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                Requests = new OrganizationRequestCollection()
            };

            foreach (var entity in Entities)
            {
                CreateRequest createRequest = new CreateRequest { Target = entity.ToEntity<Entity>() };
                multipleRequest.Requests.Add(createRequest);
            }

            ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)_service.Execute(multipleRequest);

        }
    }
}
