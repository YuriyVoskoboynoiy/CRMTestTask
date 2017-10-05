using System;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace CRMTestTask
{


    public class EmailCounterPlugin : PluginBase
    {

        public EmailCounterPlugin(string unsecure, string secure)
            : base(typeof(EmailCounterPlugin))
        {

            // TODO: Implement your custom configuration handling.
        }


        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new ArgumentNullException("localContext");
            }

            if (localContext.PluginExecutionContext.MessageName == "Create" || localContext.PluginExecutionContext.MessageName == "Update")
            {
                if (localContext.PluginExecutionContext.InputParameters.Contains("Target") && localContext.PluginExecutionContext.InputParameters["Target"] is Entity)
                {
                    Entity target = (Entity)localContext.PluginExecutionContext.InputParameters["Target"];

                    if (target.LogicalName != "email")
                    {
                        return;
                    }

                    if (target.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                    {

                        try
                        {
                            if (target.Contains("from"))
                            {

                                Entity from = target.GetAttributeValue<EntityCollection>("from").Entities.Cast<Entity>().First();
                                localContext.TracingService.Trace("ActivityPartyId: " + from.Id.ToString());
                               
                                EntityReference partId = from.GetAttributeValue<EntityReference>("partyid");
                                localContext.TracingService.Trace("UserId: " + partId.Id.ToString());

                                if (partId.LogicalName == "systemuser")
                                {

                                    Entity systemUser = localContext.OrganizationService.Retrieve("systemuser", partId.Id, new ColumnSet("new_emailcounter"));
                                    localContext.TracingService.Trace("SystemUserCardId: " + systemUser.Id.ToString());

                                    Entity updateUser = new Entity("systemuser");
                                    updateUser.Id = systemUser.Id;
                                    if (systemUser.Contains("new_emailcounter"))
                                    {
                                        updateUser["new_emailcounter"] = (int)systemUser["new_emailcounter"] + 1;
                                    }
                                    else
                                    {
                                        updateUser["new_emailcounter"] = 1;
                                    }


                                    localContext.OrganizationService.Update(updateUser);

                                }

                            }
                        }
                        catch (FaultException<OrganizationServiceFault> ex)
                        {

                            throw new InvalidPluginExecutionException("An error occurred in the FollowupPlugin plug-in.", ex);

                        }
                    }
                }
            }

        }
    }
}
