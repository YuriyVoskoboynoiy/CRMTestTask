using System;
using System.Activities;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Crm.Sdk.Messages;

namespace CRMTestTask.WF
{
    public class SentEmailWithAttachment : CodeActivity
    {
        [Input("SourceEmail")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> SourceEmail { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            #region Context
            ContextHelper.Get(executionContext);
            #endregion

           Entity emailRecordWithAttacment = AddAttachmentToEmailRecord(SourceEmail.Get<EntityReference>(executionContext));

            SendEmailToTeamMembers(emailRecordWithAttacment);

        }

        private Entity AddAttachmentToEmailRecord(EntityReference sourceEmail)
        {
            Entity emailRecord = ContextHelper._orgService.Retrieve("email", sourceEmail.Id, new ColumnSet(true));

            QueryExpression QueryAnnotation = new QueryExpression("annotation");
            QueryAnnotation.ColumnSet = new ColumnSet(new string[] { "subject", "mimetype", "filename", "documentbody" });
            QueryAnnotation.Criteria = new FilterExpression();
            QueryAnnotation.Criteria.FilterOperator = LogicalOperator.And;
            QueryAnnotation.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, ContextHelper.context.PrimaryEntityId));

            EntityCollection annotations = ContextHelper._orgService.RetrieveMultiple(QueryAnnotation);

            if (annotations.Entities.Count() > 0)
            {
                foreach (Entity item in annotations.Entities)
                {
                    if (item.Contains("documentbody"))
                    {
                        Entity emailAttachment = new Entity("activitymimeattachment");
                        if (item.Contains("subject"))
                            emailAttachment.Attributes["subject"] = item.GetAttributeValue<string>("subject");

                        emailAttachment.Attributes["objectid"] = new EntityReference("email", emailRecord.Id);
                        emailAttachment.Attributes["objecttypecode"] = "email";
                        emailAttachment.Attributes["filename"] = item.GetAttributeValue<string>("filename");
                        emailAttachment.Attributes["body"] = item.GetAttributeValue<string>("documentbody");
                        emailAttachment.Attributes["mimetype"] = item.GetAttributeValue<string>("mimetype");

                        ContextHelper._orgService.Create(emailAttachment);

                    }
                }
            }

            return emailRecord;
        }

        private void SendEmailToTeamMembers(Entity emailRecord)
        {

            EntityCollection To = new EntityCollection();

            Entity CurrentContact = ContextHelper._orgService.Retrieve("contact", ContextHelper.context.PrimaryEntityId, new ColumnSet("new_teamcontactid"));

            if (CurrentContact.Contains("new_teamcontactid"))
            {
                EntityReference contactTeam = CurrentContact.GetAttributeValue<EntityReference>("new_teamcontactid");
                ContextHelper.tracingService.Trace("Contact Team: " + contactTeam.Name);

                #region Get team members

                QueryExpression teamMemberQuery = new QueryExpression("systemuser");
                teamMemberQuery.ColumnSet = new ColumnSet(true);
                LinkEntity teamLink = new LinkEntity("systemuser", "teammembership", "systemuserid", "systemuserid", JoinOperator.Inner);
                ConditionExpression teamCondition = new ConditionExpression("teamid", ConditionOperator.Equal, contactTeam.Id);
                teamLink.LinkCriteria.AddCondition(teamCondition);
                teamMemberQuery.LinkEntities.Add(teamLink);

                EntityCollection teamMembers = ContextHelper._orgService.RetrieveMultiple(teamMemberQuery);

                if (teamMembers.Entities.Count() > 0)
                {
                    foreach (Entity member in teamMembers.Entities)
                    {
                        Entity partyMember = new Entity("activityparty");
                        partyMember["partyid"] = new EntityReference("systemuser", member.Id);

                        To.Entities.Add(partyMember);
                    }

                }
                #endregion

                #region Update email
                Entity email = new Entity("email");
                email.Id = emailRecord.Id;
                email["to"] = To;
                ContextHelper._orgService.Update(email);
                #endregion

                #region Send email
                SendEmailRequest sendEmail = new SendEmailRequest();
                sendEmail.EmailId = emailRecord.Id;
                sendEmail.TrackingToken = "";
                sendEmail.IssueSend = true;

                SendEmailResponse response = (SendEmailResponse)ContextHelper._orgService.Execute(sendEmail);
                ContextHelper.tracingService.Trace("Response subject: " + response.Subject);
                #endregion
            }

        }
    }
}
