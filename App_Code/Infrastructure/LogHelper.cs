using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

/// <summary>
/// Summary description for LogHelper
/// </summary>

namespace FilesToPDFConvertor
{
    public class LogHelper
    {
        public void AddExceptionLogs(string errorMessage, string errorSource, string errorStackTrace,string pageName, string methodName,  string createdBy, Int32 moduleId, Int32 companyid)
        {
            SqlParameter[] parameters = new SqlParameter[8];
            parameters[0] = new SqlParameter("@ERROR_MESSAGE",errorMessage);
            parameters[1] = new SqlParameter("@ERROR_SOURCE",errorSource);
            parameters[2] = new SqlParameter("@ERROR_STACK_TRACE",errorStackTrace);
            parameters[3] = new SqlParameter("@PAGE_NAME",pageName);
            parameters[4] = new SqlParameter("@METHOD_NAME",methodName);
            parameters[5] = new SqlParameter("@CREATED_BY",createdBy);
            parameters[6] = new SqlParameter("@MODULE_ID", moduleId);
            parameters[7] = new SqlParameter("@COMPANY_ID", companyid);
            SQLHelper.ExecuteScalar(SQLHelper.GetConnString(), CommandType.StoredProcedure, "SP_PROCS_LOG_EXCEPTION", "PROCS_ADMIN", parameters);
        }
    }
}
