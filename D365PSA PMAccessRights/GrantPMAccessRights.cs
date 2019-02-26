using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;


namespace D365PSA_PMAccessRights
{
    public class GrantPMAccessRights : IPlugin
    {

        IOrganizationService service; Entity currentEntity, preImageEntity, postImageEntity;
        IPluginExecutionContext context; ITracingService trace; Guid previousPM = Guid.Empty, currentPM = Guid.Empty;

        /// <summary>
        /// Generic method to connect to the Dynamics 365 environment.
        /// </summary>
        /// <param name="serviceProvider">Object of IServiceProvider class</param>
        private void GetOrganizationService(IServiceProvider serviceProvider)
        {
            try
            {
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                service = ((IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory))).CreateOrganizationService(context.UserId);
                trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService)); //((ITracingService)serviceProvider.GetService(typeof(ITracingService)));
            }
            catch (Exception e)
            {
                //   service.Create(ErrorLogger.LogError(e.Message, e.StackTrace, "AddMembersToAccount", "GetOrganizationServices"));
            }
        }

        /// <summary>
        /// Gets the current record from the Execution context
        /// </summary>
        private void GetCurrentRecord()
        {
            try
            {
                currentEntity = (Entity)context.InputParameters["Target"];
                if (context.PreEntityImages.Contains("ProjectImages"))
                {
                    preImageEntity = (Entity)context.PreEntityImages["ProjectImages"];
                    if (preImageEntity.Attributes.Contains(Constants.Project.Attr_ProjectManager))
                    {
                        previousPM = preImageEntity.GetAttributeValue<EntityReference>(Constants.Project.Attr_ProjectManager).Id;
                    }
                }
                if (context.PostEntityImages.Contains("ProjectImages"))
                {
                    postImageEntity = (Entity)context.PostEntityImages["ProjectImages"];
                    if (postImageEntity.Attributes.Contains(Constants.Project.Attr_ProjectManager))
                    {
                        currentPM = postImageEntity.GetAttributeValue<EntityReference>(Constants.Project.Attr_ProjectManager).Id;
                    }
                }
                
                //if (currentEntity.Attributes.Contains(Constants.Project.Attr_ProjectManager))
                //{
                //    newPM = currentEntity.GetAttributeValue<EntityReference>(Constants.Project.Attr_ProjectManager).Id;
                //}
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Generic method to Grant Access to the users with whom the Project is supposed to be shared.
        /// </summary>
        /// <param name="userId">Primary key of the Systemuser record.</param>
        private void GrantAccessToNewPM(Guid userId)
        {
            try
            {
                var systemUser1Ref = new EntityReference(Constants.Entity.Ent_SystemUser, userId);

                var projectReference = new EntityReference(Constants.Entity.Ent_Project, currentEntity.Id);
                // Grant the first user read access to the created lead.
                var grantAccessRequest1 = new GrantAccessRequest
                {
                    PrincipalAccess = new PrincipalAccess
                    {
                        AccessMask = AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess | AccessRights.AssignAccess | AccessRights.ShareAccess,
                        Principal = systemUser1Ref
                    },
                    Target = projectReference
                };
                service.Execute(grantAccessRequest1);

            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Generic method to remove the access from the User who are removed from the Project.
        /// </summary>
        /// <param name="userId">Primary key System user record.</param>
        private void RevokeAccessFromPreviousPM(Guid userId)
        {
            try
            {
                var systemUser1Ref = new EntityReference(Constants.Entity.Ent_SystemUser, userId);

                var projectReference = new EntityReference(Constants.Entity.Ent_Project, currentEntity.Id);
                // Grant the first user read access to the created lead.
                var grantAccessRequest1 = new RevokeAccessRequest
                {
                    Revokee = systemUser1Ref,
                    Target = projectReference
                };
                service.Execute(grantAccessRequest1);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// First method of the plugin exeuction in Dynamics 365
        /// </summary>
        /// <param name="serviceProvider"></param>
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                GetOrganizationService(serviceProvider);
                GetCurrentRecord();
                if (currentPM != Guid.Empty)
                {
                    GrantAccessToNewPM(currentPM);
                    if (previousPM != Guid.Empty && previousPM != currentPM)
                    {
                        RevokeAccessFromPreviousPM(previousPM);
                    }
                }
               
            }
            catch (Exception e)
            {

                throw;
            }
        }
    }
}
