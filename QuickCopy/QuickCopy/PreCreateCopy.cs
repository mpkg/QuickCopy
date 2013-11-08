using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace Unizap.Addon.QuickCopy
{
    public class PreCreateCopy : IPlugin
    {
        private IOrganizationService service = null;
        private IPluginExecutionContext context = null;
        private IOrganizationServiceFactory serviceFactory = null;
        private Entity targetEntity = null;
        private Entity parentEntity = null;
        private Guid parentId;
        private string primaryNameAttribute, primaryIdAttribute;
        private string[] excludeFieldNames;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                targetEntity = context.InputParameters["Target"] as Entity;
                if (targetEntity.Contains("new_quickcopyparentid"))
                {
                    if (Guid.TryParse((string)targetEntity["new_quickcopyparentid"], out parentId))
                    {
                        LicenseManager lic = new LicenseManager(ref service, ref context);
                        int result = lic.ValidateLicense("unizap_/QuickCopy/QuickCopy_License.xml", "QuickCopy");
                        if (result == Result.LicenseValid || result == Result.TrialLicense)
                        {
                            //Retreive record with id = parentid
                            parentEntity = service.Retrieve(context.PrimaryEntityName, parentId, new ColumnSet(true));

                            //set primary attributes
                            SetPrimaryNameAttribute();

                            SetExcludeFieldNames(context.PrimaryEntityName);

                            #region remove fields

                            if (excludeFieldNames != null)
                            {
                                foreach (string field in excludeFieldNames)
                                {
                                    if (parentEntity.Attributes.Contains(field.Trim().ToLower()))
                                    {
                                        parentEntity.Attributes.Remove(field);
                                    }
                                }
                            }

                            if (parentEntity.Attributes.Contains("new_quickcopyparentid"))
                                parentEntity.Attributes.Remove("new_quickcopyparentid");

                            if (parentEntity.Attributes.Contains("ownerid"))
                                parentEntity.Attributes.Remove("ownerid");

                            if (parentEntity.Attributes.Contains("activityid"))
                                parentEntity.Attributes.Remove("activityid");

                            if (parentEntity.Attributes.Contains(primaryIdAttribute))
                                parentEntity.Attributes.Remove(primaryIdAttribute);

                            if (parentEntity.Attributes.Contains("address1_addressid"))
                                parentEntity.Attributes.Remove("address1_addressid");

                            if (parentEntity.Attributes.Contains("address2_addressid"))
                                parentEntity.Attributes.Remove("address2_addressid");

                            if (parentEntity.Attributes.Contains("owningbusinessunit"))
                                parentEntity.Attributes.Remove("owningbusinessunit");

                            if (parentEntity.Attributes.Contains("organizationid"))
                                parentEntity.Attributes.Remove("organizationid");

                            if (parentEntity.Attributes.Contains("statuscode"))
                                parentEntity.Attributes.Remove("statuscode");

                            if (parentEntity.Attributes.Contains("statecode"))
                                parentEntity.Attributes.Remove("statecode");

                            if (parentEntity.Attributes.Contains("opportunitystatecode"))
                                parentEntity.Attributes.Remove("opportunitystatecode");

                            if (parentEntity.Attributes.Contains("invoicestatecode"))
                                parentEntity.Attributes.Remove("invoicestatecode");

                            if (parentEntity.Attributes.Contains("quotestatecode"))
                                parentEntity.Attributes.Remove("quotestatecode");

                            if (parentEntity.Attributes.Contains("salesorderstatecode"))
                                parentEntity.Attributes.Remove("salesorderstatecode");

                            if (parentEntity.Attributes.Contains("productnumber"))
                                parentEntity.Attributes.Remove("productnumber");

                            if (parentEntity.Attributes.Contains("quotenumber"))
                                parentEntity.Attributes.Remove("quotenumber");

                            if (parentEntity.Attributes.Contains("ordernumber"))
                                parentEntity.Attributes.Remove("ordernumber");

                            if (parentEntity.Attributes.Contains("invoicenumber"))
                                parentEntity.Attributes.Remove("invoicenumber");

                            if (parentEntity.Attributes.Contains("ticketnumber"))
                                parentEntity.Attributes.Remove("ticketnumber");

                            #endregion remove fields

                            foreach (var attribute in parentEntity.Attributes)
                            {
                                targetEntity[attribute.Key] = attribute.Value;
                            }
                            //append [Copy] to primary name attribute
                            targetEntity[primaryNameAttribute] = "[Copy] " + (string)parentEntity[primaryNameAttribute];
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("Oops!Quick Copy evaluation period over. Please contact info@unizap.com for license!");
                        }
                    }
                }
            }
        }

        private void SetPrimaryNameAttribute()
        {
            RetrieveEntityRequest req = new RetrieveEntityRequest();
            RetrieveEntityResponse res = new RetrieveEntityResponse();

            req.LogicalName = context.PrimaryEntityName;
            req.EntityFilters = EntityFilters.Entity;
            req.RetrieveAsIfPublished = true;

            res = (RetrieveEntityResponse)service.Execute(req);

            EntityMetadata entityMetaData = res.EntityMetadata;
            primaryIdAttribute = entityMetaData.PrimaryIdAttribute;
            primaryNameAttribute = entityMetaData.PrimaryNameAttribute;
        }

        private void SetExcludeFieldNames(string entityName)
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