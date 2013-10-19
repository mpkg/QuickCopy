using System;
using System.Collections.Generic;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Unizap.Addon.QuickCopy
{
    public class PostCreateCopy : IPlugin
    {
        private Guid currentRecordId = Guid.Empty;
        private Guid parentRecordId = Guid.Empty;
        private DataCollection<Entity> configRecords = null;
        private Entity targetEntity = null;
        private string entityPrimaryFieldName = string.Empty;
        private string primaryEntityName = string.Empty;
        private string relatedEntityName = string.Empty;
        private string relationshipName = string.Empty;
        private string intersectEntityName = string.Empty;
        private string referencingAttributeName = string.Empty;
        private string intersectFieldName = string.Empty;
        private int relatedEntityCount = 0;
        private string[] excludeFieldNames;

        private IPluginExecutionContext context = null;
        private IOrganizationServiceFactory serviceFactory = null;
        private IOrganizationService service = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);

            primaryEntityName = context.PrimaryEntityName;
            currentRecordId = context.PrimaryEntityId;

            if (context.InputParameters.Contains("Target"))
                targetEntity = context.InputParameters["Target"] as Entity;
            else
                return;

            if (targetEntity.Contains("new_quickcopyparentid"))
            {
                if (Guid.TryParse((string)targetEntity["new_quickcopyparentid"], out parentRecordId))
                {
                    QueryByAttribute query = new QueryByAttribute("unizap_quickcopyrelatedentityconfig");
                    query.AddAttributeValue("unizap_primaryentityname", primaryEntityName);
                    query.ColumnSet = new ColumnSet(true);

                    EntityCollection recordsCollection = service.RetrieveMultiple(query);
                    configRecords = recordsCollection.Entities;
                    relatedEntityCount = configRecords.Count;

                    if (relatedEntityCount > 0)
                    {
                        relatedEntityCount = configRecords.Count;
                    }

                    CopyRelatedEntity();

                    AddNtoNReference();
                }
            }
        }

        private void AddNtoNReference()
        {
            for (int i = 0; i < relatedEntityCount; i++)
            {
                if (((OptionSetValue)configRecords[i].Attributes["unizap_relationshiptype"]).Value == 2)
                {
                    if (configRecords[i].Attributes.Contains("unizap_referencingattributename"))
                    {
                        intersectFieldName = configRecords[i].Attributes["unizap_referencingattributename"].ToString();
                        intersectEntityName = configRecords[i].Attributes["unizap_intersectentityname"].ToString();
                        relatedEntityName = configRecords[i].Attributes["unizap_relatedentityname"].ToString();
                        relationshipName = configRecords[i].Attributes["unizap_relationshipname"].ToString();
                        QueryExpression query = new QueryExpression()
                        {
                            EntityName = relatedEntityName,
                            ColumnSet = new ColumnSet(true),
                            LinkEntities =
                                {
                                    new LinkEntity
                                    {
                                        LinkFromEntityName = relatedEntityName,
                                        LinkFromAttributeName = intersectFieldName,
                                        LinkToEntityName = intersectEntityName,
                                        LinkToAttributeName = intersectFieldName,
                                        LinkCriteria = new FilterExpression
                                        {
                                            FilterOperator = LogicalOperator.And,
                                            Conditions =
                                            {
                                                new ConditionExpression
                                                {
                                                    AttributeName = primaryEntityName + "id",
                                                    Operator = ConditionOperator.Equal,
                                                    Values = { parentRecordId }
                                                }
                                            }
                                        }
                                    }
                                }
                        };
                        EntityCollection records = service.RetrieveMultiple(query);

                        AssociateRequest associationRequest = new AssociateRequest();
                        associationRequest.RelatedEntities = new EntityReferenceCollection();
                        associationRequest.Target = new EntityReference(primaryEntityName.ToLower(), targetEntity.Id);
                        for (int j = 0; j < records.Entities.Count; j++)
                        {
                            associationRequest.RelatedEntities.Add(new EntityReference(records.Entities[j].LogicalName, records.Entities[j].Id));
                        }
                        associationRequest.Relationship = new Relationship(relationshipName);
                        service.Execute(associationRequest);
                    }
                }
            }
        }

        private void CopyRelatedEntity()
        {
            try
            {
                List<string> relatedEntites = new List<string>();

                for (int i = 0; i < relatedEntityCount; i++)
                {
                    if (((OptionSetValue)configRecords[i].Attributes["unizap_relationshiptype"]).Value == 1)
                    {
                        if (configRecords[i].Attributes.Contains("unizap_relatedentityname"))

                            relatedEntites.Add(configRecords[i].Attributes["unizap_relatedentityname"].ToString());
                    }
                }

                foreach (string relatedEntityName in relatedEntites)
                {
                    // Create the request object.
                    RetrieveMultipleRequest req = new RetrieveMultipleRequest();

                    RetrieveMultipleResponse resp = new RetrieveMultipleResponse();

                    //Retrive the Child entity whose Original parent ID is specified in the iotap_parentid
                    QueryExpression query = new QueryExpression(relatedEntityName);//no record

                    query.ColumnSet = new ColumnSet();
                    query.ColumnSet.AllColumns = true;

                    ConditionExpression condition = new ConditionExpression();

                    referencingAttributeName = primaryEntityName + "id";

                    for (int i = 0; i < relatedEntityCount; i++)
                    {
                        if (relatedEntityName.Equals(configRecords[i].Attributes["unizap_relatedentityname"].ToString(), StringComparison.OrdinalIgnoreCase))
                        {
                            if (configRecords[i].Attributes.Contains("unizap_referencingattributename"))
                            {
                                referencingAttributeName = configRecords[i].Attributes["unizap_referencingattributename"].ToString();
                                break;
                            }
                        }
                    }

                    if (String.IsNullOrEmpty(referencingAttributeName)) continue;

                    condition.AttributeName = referencingAttributeName;
                    condition.Operator = ConditionOperator.Equal;
                    condition.Values.Add(parentRecordId);

                    FilterExpression filter = new FilterExpression();
                    filter.AddCondition(condition);

                    query.Criteria = filter;

                    // Set the properties of the request object.
                    req.Query = query;

                    resp = (RetrieveMultipleResponse)service.Execute(req);

                    if (resp.EntityCollection.Entities.Count > 0)
                    {
                        foreach (Entity entity in resp.EntityCollection.Entities)
                        {
                            SetExcludeFieldNames(relatedEntityName);
                            CopyRecord(entity, relatedEntityName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new SystemException("An unknown exception was received. " + ex.Message);
            }
        }

        private void CopyRecord(Entity retrievedRecord, string child)
        {
            try
            {
                retrievedRecord.Id = Guid.Empty;

                if (excludeFieldNames != null)
                {
                    foreach (string excludeFieldName in excludeFieldNames)
                    {
                        if (retrievedRecord.Attributes.Contains(excludeFieldName.Trim().ToLower()))
                        {
                            retrievedRecord.Attributes.Remove(excludeFieldName);
                        }
                    }
                }

                //Removing the attribires which are not necessary

                if (retrievedRecord.Attributes.Contains("ownerid"))
                    retrievedRecord.Attributes.Remove("ownerid");

                if (retrievedRecord.Attributes.Contains("activityid"))
                    retrievedRecord.Attributes.Remove("activityid");

                if (retrievedRecord.Attributes.Contains(retrievedRecord.LogicalName.ToLower() + "id"))
                    retrievedRecord.Attributes.Remove(retrievedRecord.LogicalName.ToLower() + "id");

                if (retrievedRecord.Attributes.Contains("address1_addressid"))
                    retrievedRecord.Attributes.Remove("address1_addressid");

                if (retrievedRecord.Attributes.Contains("address2_addressid"))
                    retrievedRecord.Attributes.Remove("address2_addressid");

                if (retrievedRecord.Attributes.Contains("owningbusinessunit"))
                    retrievedRecord.Attributes.Remove("owningbusinessunit");

                if (retrievedRecord.Attributes.Contains("statuscode"))
                    retrievedRecord.Attributes.Remove("statuscode");

                if (retrievedRecord.Attributes.Contains("statecode"))
                    retrievedRecord.Attributes.Remove("statecode");

                if (retrievedRecord.Attributes.Contains("opportunitystatecode"))
                    retrievedRecord.Attributes.Remove("opportunitystatecode");

                if (retrievedRecord.Attributes.Contains("invoicestatecode"))
                    retrievedRecord.Attributes.Remove("invoicestatecode");

                if (retrievedRecord.Attributes.Contains("statuscode"))
                    retrievedRecord.Attributes.Remove("statuscode");

                if (retrievedRecord.Attributes.Contains("invoicestatecode"))
                    retrievedRecord.Attributes.Remove("invoicestatecode");

                if (retrievedRecord.Attributes.Contains("quotestatecode"))
                    retrievedRecord.Attributes.Remove("quotestatecode");

                if (retrievedRecord.Attributes.Contains("salesorderstatecode"))
                    retrievedRecord.Attributes.Remove("salesorderstatecode");

                if (retrievedRecord.Attributes.Contains("productnumber"))
                    retrievedRecord.Attributes.Remove("productnumber");

                if (retrievedRecord.Attributes.Contains("quotenumber"))
                    retrievedRecord.Attributes.Remove("quotenumber");

                if (retrievedRecord.Attributes.Contains("ordernumber"))
                    retrievedRecord.Attributes.Remove("ordernumber");

                if (retrievedRecord.Attributes.Contains("invoicenumber"))
                    retrievedRecord.Attributes.Remove("invoicenumber");

                if (retrievedRecord.Attributes.Contains("ticketnumber"))
                    retrievedRecord.Attributes.Remove("ticketnumber");

                //for custom entity
                if (retrievedRecord.LogicalName.Equals(child, StringComparison.OrdinalIgnoreCase))
                {
                    if (retrievedRecord.Attributes.Contains(referencingAttributeName))
                    {
                        retrievedRecord.Attributes[referencingAttributeName] = new EntityReference(primaryEntityName, currentRecordId);
                    }
                }

                Guid newCreatedREcordId = service.Create(retrievedRecord);
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                throw new FaultException(e.Message.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
        }

        public void SetExcludeFieldNames(string entityName)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = "unizap_quickcopyexcludefieldsconfig";
            query.ColumnSet = new ColumnSet() { AllColumns = true };
            query.Criteria = new FilterExpression();
            query.Criteria.AddCondition("unizap_entityname", ConditionOperator.Equal, entityName);

            EntityCollection records = service.RetrieveMultiple(query);

            //Iterate across results and output fullname
            foreach (Entity excludeFieldsConfig in records.Entities)
            {
                excludeFieldNames = excludeFieldsConfig.GetAttributeValue<string>("unizap_excludefieldslist").Split(',');
            }
        }
    }
}