using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for Global
/// </summary>
/// 

namespace ProCS.Infrastructure
{
    public class Global
    {
        public Global()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public enum TaskType
        {
            All,
            Any
        }

        public enum TaskStatus
        {
            Assigned,
            Pending,
            Approved,
            Rejected
        }
    }
}