using Gits.NourNet.EarlyBound;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;

namespace LeadActiveQuestions
{
    class PluginDemo
    {
        static void Main()
        {
            Uri organizationUri = new Uri("https://nn-erp-dev.nour.net.sa/NourNetTestOrg/XRMServices/2011/Organization.svc");
            Uri homeRealmUri = null;
            ClientCredentials credentials = new ClientCredentials();
            // set default credentials for OrganizationService
            //credentials.Windows.ClientCredential = (NetworkCredential)CredentialCache.DefaultCredentials;
            // or
            credentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;
            OrganizationServiceProxy orgProxy = new OrganizationServiceProxy(organizationUri, homeRealmUri, credentials, null);
            // This statement is required to enable early-bound type support.
            orgProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());

            IOrganizationService _service = (IOrganizationService)orgProxy;

            EntityCollection activeQuestions = GetActiveQuestions(_service);
            if (isNoneSale(new Guid("1B4A8F34-2B98-E511-80C8-005056A66147"), "Non Sales Team", _service))
                Console.WriteLine("has noneSale!");
            if (activeQuestions.Entities.Count == 0)
                Console.WriteLine("There are no Active Questions to be Added!");


            foreach (var rec in activeQuestions.Entities)
            {
                Guid questionGuid = new Guid(rec.Attributes[gits_questionnaire.Fields.gits_questionnaireId].ToString());
                Console.WriteLine(questionGuid.ToString() + " -> GUID");
                if (questionGuid != null)
                {
                    gits_leadquesans question = new gits_leadquesans();
                    question.gits_Lead = new EntityReference(Lead.EntityLogicalName, new Guid("0A0F2286-36EF-E511-80CF-005056A6101C"));
                    question.gits_Questions = new EntityReference(gits_questionnaire.EntityLogicalName, questionGuid);
                    Guid id = _service.Create(question);
                    if (id != null)
                        Console.WriteLine(id.ToString() + " -> gits_leadquesans creted with guid");
                }

            }
            // Keep the console window open in debug mode.
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        public static EntityCollection GetActiveQuestions(IOrganizationService _service)
        {
            ConditionExpression c = new ConditionExpression(gits_questionnaire.Fields.statuscode, ConditionOperator.Equal, 1);
            FilterExpression filter = new FilterExpression();
            filter.Conditions.Add(c);

            QueryExpression query = new QueryExpression(gits_questionnaire.EntityLogicalName);
            query.ColumnSet.AddColumns(gits_questionnaire.Fields.gits_questionnaireId);
            query.Criteria.AddFilter(filter);

            return _service.RetrieveMultiple(query);
        }

        public static bool isOnlyNoneSale(Guid user, string roleName, IOrganizationService service)
        {
            bool hasRole = false;
            roleName = "Non Sales Team";
            QueryExpression query = new QueryExpression();
            query.EntityName = "role"; //role entity name
            ColumnSet cols = new ColumnSet();
            cols.AddColumn("name"); //We only need role name
            query.ColumnSet = cols;
            ConditionExpression ce = new ConditionExpression();
            ce.AttributeName = "systemuserid";
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add(user);

            //system roles

            LinkEntity linkRole = new LinkEntity();
            linkRole.LinkFromAttributeName = "roleid";
            linkRole.LinkFromEntityName = "role"; //FROM
            linkRole.LinkToEntityName = "systemuserroles";
            linkRole.LinkToAttributeName = "roleid";

            //system users
            LinkEntity linkSystemusers = new LinkEntity();
            linkSystemusers.LinkFromEntityName = "systemuserroles";
            linkSystemusers.LinkFromAttributeName = "systemuserid";
            linkSystemusers.LinkToEntityName = "systemuser";
            linkSystemusers.LinkToAttributeName = "systemuserid";

            linkSystemusers.LinkCriteria = new FilterExpression();
            linkSystemusers.LinkCriteria.Conditions.Add(ce);
            linkRole.LinkEntities.Add(linkSystemusers);
            query.LinkEntities.Add(linkRole);

            // query returns records means user has this role
            EntityCollection result = service.RetrieveMultiple(query);
            foreach (var item in result.Entities)
            {
                if (item.Attributes["name"] == roleName) hasRole = true;

            }
            // has more than one role but among them is sales person
            return result.Entities.Count > 1 && hasRole ? false : true;
        }

    }
}
