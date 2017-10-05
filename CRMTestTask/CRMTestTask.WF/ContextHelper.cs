using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace CRMTestTask.WF
{
    static class ContextHelper
    {
        public static ITracingService tracingService;
        public static IWorkflowContext context;
        public static IOrganizationServiceFactory serviceFactory;
        public static IOrganizationService _orgService;

        public static void Get(CodeActivityContext executionContext)
        {
            //Create the tracing service
            tracingService = executionContext.GetExtension<ITracingService>();
            // Get the context service.
            context = executionContext.GetExtension<IWorkflowContext>();
            serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.
            _orgService = serviceFactory.CreateOrganizationService(context.UserId);

        }

        public static void GetWithAdminAccess(CodeActivityContext executionContext)
        {
            //Create the tracing service
            tracingService = executionContext.GetExtension<ITracingService>();
            // Get the context service.
            context = executionContext.GetExtension<IWorkflowContext>();
            serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.
            _orgService = serviceFactory.CreateOrganizationService(null);

        }
    }
}
