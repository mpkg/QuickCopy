﻿using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace Unizap.Addon.QuickCopy
{
    public class PreCreateClone : IPlugin
    {
        private IOrganizationService service = null;
        private IPluginExecutionContext context = null;
        private IOrganizationServiceFactory serviceFactory = null;
        private Entity targetEntity = null;
        private Entity parentEntity = null;
        private Guid parentId;
        private string primaryNameAttribute, primaryIdAttribute;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                targetEntity = context.InputParameters["Target"] as Entity;
                if (targetEntity.Contains("new_parentid"))
                {
                    if (Guid.TryParse((string)targetEntity["new_parentid"], out parentId))
                    {
                        LicenseManager lic = new LicenseManager(ref service, ref context);
                        int result = lic.ValidateLicense("unizap_/QuickCopy/QuickCopy_License.xml", "QuickCopy");
                        if (result == Result.LicenseValid || result == Result.TrialLicense)
                        {
                            //Retreive record with id = parentid
                            parentEntity = service.Retrieve(context.PrimaryEntityName, parentId, new ColumnSet(true));

                            //set primary attributes
                            FetchPrimaryNameAttribute();

                            #region remove primary key fields

                            if (parentEntity.Attributes.Contains("new_parentid"))
                                parentEntity.Attributes.Remove("new_parentid");

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

                            if (parentEntity.Attributes.Contains("statuscode"))
                                parentEntity.Attributes.Remove("statuscode");

                            if (parentEntity.Attributes.Contains("statecode"))
                                parentEntity.Attributes.Remove("statecode");

                            if (parentEntity.Attributes.Contains("opportunitystatecode"))
                                parentEntity.Attributes.Remove("opportunitystatecode");

                            if (parentEntity.Attributes.Contains("invoicestatecode"))
                                parentEntity.Attributes.Remove("invoicestatecode");

                            if (parentEntity.Attributes.Contains("statuscode"))
                                parentEntity.Attributes.Remove("statuscode");

                            if (parentEntity.Attributes.Contains("invoicestatecode"))
                                parentEntity.Attributes.Remove("invoicestatecode");

                            if (parentEntity.Attributes.Contains("quotestatecode"))
                                parentEntity.Attributes.Remove("quotestatecode");

                            if (parentEntity.Attributes.Contains("salesorderstatecode"))
                                parentEntity.Attributes.Remove("salesorderstatecode");

                            if (parentEntity.Attributes.Contains("quotenumber"))
                                parentEntity.Attributes.Remove("quotenumber");

                            #endregion remove primary key fields

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

        private void FetchPrimaryNameAttribute()
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
    }
}