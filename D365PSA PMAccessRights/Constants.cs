using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace D365PSA_PMAccessRights
{
    static class Constants
    {

        internal static class Entity
        {
            internal const string Ent_SystemUser = "systemuser";
            internal const string Ent_Project = "msdyn_project";
        }

        internal static class Project
        {
            internal const string Attr_ProjectManager = "msdyn_projectmanager";
        }
    }
}
