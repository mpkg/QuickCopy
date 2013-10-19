using System;
using System.Collections.Generic;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Unizap.Addon.QuickCopy
{
    public class RelatedEntityConfiguration : IPlugin
    {
        private IPluginExecutionContext context = null;
        private IOrganizationService service = null;
        private IOrganizationServiceFactory serviceFactory = null;

        private string primaryEntityName = string.Empty;
        private string relatedEntityName = string.Empty;
        private string relationshipName = string.Empty;
        private string referencingAttribute = string.Empty;

        private Entity configuration = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.MessageName == "Create")
            {
                configuration = context.InputParameters["Target"] as Entity;
                primaryEntityName = configuration["unizap_primaryentityname"].ToString().ToLower();
                relationshipName = configuration["unizap_relationshipname"].ToString();

                CreateParentIdAttribute();
                SetConfigurationFields();
            }
        }

        public void CreateParentIdAttribute()
        {
            try
            {
                StringAttributeMetadata stringAttribute = new StringAttributeMetadata
                {
                    // Set base properties
                    SchemaName = "new_quickcopyparentid",
                    DisplayName = new Label("Quick Copy Parent Id", 1033),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Description = new Label("Quick Copy- Stores GUID of the parent record.", 1033),

                    // Set extended properties
                    MaxLength = 100,
                    IsAuditEnabled = new BooleanManagedProperty(false)
                };

                // Create the request.
                CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                {
                    EntityName = primaryEntityName,
                    Attribute = stringAttribute
                };

                // Execute the request.
                service.Execute(createAttributeRequest);
            }
            catch (Exception ex)
            {
            }
        }

        public void SetConfigurationFields()
        {
            try
            {
                //Retrieve primary field name of the primary entity
                RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
                    LogicalName = primaryEntityName
                };
                RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);

                configuration["unizap_primaryfieldname"] = retrieveEntityResponse.EntityMetadata.PrimaryNameAttribute;

                if (!string.IsNullOrEmpty(relationshipName))
                {
                    //Retrieve the Many-to-many relationship using the Name.
                    RetrieveRelationshipRequest retrieveOneToManyRequest =
                        new RetrieveRelationshipRequest { Name = relationshipName };
                    RetrieveRelationshipResponse retrieveOneToManyResponse =
                        (RetrieveRelationshipResponse)service.Execute(retrieveOneToManyRequest);

                    DataCollection<string, object> results = retrieveOneToManyResponse.Results;
                    System.Collections.IEnumerator en = results.GetEnumerator();
                    while (en.MoveNext())
                    {
                        if (((KeyValuePair<string, object>)(en.Current)).Value.GetType().Name == "OneToManyRelationshipMetadata")
                        {
                            configuration["unizap_relationshiptype"] = new OptionSetValue(1);
                            configuration["unizap_referencingattributename"] = ((OneToManyRelationshipMetadata)(((KeyValuePair<string, object>)(en.Current)).Value)).ReferencingAttribute;
                            configuration["unizap_relatedentityname"] = ((OneToManyRelationshipMetadata)(((KeyValuePair<string, object>)(en.Current)).Value)).ReferencingEntity;
                        }
                        else if (((KeyValuePair<string, object>)(en.Current)).Value.GetType().Name == "ManyToManyRelationshipMetadata")
                        {
                            configuration["unizap_relationshiptype"] = new OptionSetValue(2);
                            configuration["unizap_intersectentityname"] = ((ManyToManyRelationshipMetadata)(((KeyValuePair<string, object>)(en.Current)).Value)).IntersectEntityName;
                            if (primaryEntityName == ((ManyToManyRelationshipMetadata)(((KeyValuePair<string, object>)(en.Current)).Value)).Entity1LogicalName)
                            {
                                configuration["unizap_relatedentityname"] = ((ManyToManyRelationshipMetadata)(((KeyValuePair<string, object>)(en.Current)).Value)).Entity2LogicalName;
                                configuration["unizap_referencingattributename"] = ((ManyToManyRelationshipMetadata)(((KeyValuePair<string, object>)(en.Current)).Value)).Entity2IntersectAttribute;
                            }
                            else if (primaryEntityName == ((ManyToManyRelationshipMetadata)(((KeyValuePair<string, object>)(en.Current)).Value)).Entity2LogicalName)
                            {
                                configuration["unizap_relatedentityname"] = ((ManyToManyRelationshipMetadata)(((KeyValuePair<string, object>)(en.Current)).Value)).Entity1LogicalName;
                                configuration["unizap_referencingattributename"] = ((ManyToManyRelationshipMetadata)(((KeyValuePair<string, object>)(en.Current)).Value)).Entity1IntersectAttribute;
                            }
                        }
                    }
                }
            }
            catch (FaultException faultEx)
            {
                throw new Exception("SetConfigurationFields() - An unknown exception was received. " + faultEx.Message);
            }
        }
    }
}