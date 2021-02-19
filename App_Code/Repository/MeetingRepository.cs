using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using ProCS.Infrastructure;
using FilesToPDFConvertor;
using System.Configuration;

namespace FilesToPDFConvertor
{
    public class MeetingRepository
    {
        private static String connectionString = SQLHelper.GetConnString();
        private static String dbName = SQLHelper.GetDBName();

        #region "Get Agenda Item Documents"

        public Meeting GetAgendaDocs(Int32 companyId, String meetingId, String agendaId)
        {
            Meeting objMeeting = new Meeting();
            List<AgendaItems> lstAgendaItems = null;
            try
            {
                objMeeting.ID = Convert.ToInt32(meetingId);
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@MODE", "GET_ITEM_DOCUMENTS_FOR_PDF_CONVERSION"));
                        cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(new SqlParameter("@MEETING_ID", meetingId));
                        cmd.Parameters.Add(new SqlParameter("@AGENDA_ID_CSV", agendaId));
                        SqlDataReader rdr = cmd.ExecuteReader();
                        if (rdr.HasRows)
                        {
                            lstAgendaItems = new List<AgendaItems>();
                            while (rdr.Read())
                            {
                                AgendaItems obj = new AgendaItems();
                                obj.ID = Convert.ToInt32(rdr["AGENDA_ID"]);
                                obj.agendaDoc = (!String.IsNullOrEmpty(Convert.ToString(rdr["AGENDA_DOC"]))) ? Convert.ToString(rdr["AGENDA_DOC"]) : String.Empty;
                                lstAgendaItems.Add(obj);
                            }
                        }
                        rdr.Close();

                        if (lstAgendaItems != null)
                        {
                            if (lstAgendaItems.Count > 0)
                            {
                                objMeeting.agendaItems = lstAgendaItems;
                            }
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAgendaDocs", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return objMeeting;
        }

        #endregion

        #region "Update Agenda Page Count"

        public bool UpdateAgendaFilesPageCount(Meeting objMeeting, Int32 companyId)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS", conn))
                    {
                        if (objMeeting.agendaItems != null)
                        {
                            if (objMeeting.agendaItems.Count > 0)
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.CommandTimeout = 0;
                                cmd.Parameters.Clear();
                                cmd.Parameters.Add(new SqlParameter("@MODE", "UPDATE_PAGE_COUNT"));
                                cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                                cmd.Parameters.Add(new SqlParameter("@XML_DATA", CreateXmlAgendaPageCounts(objMeeting, companyId)));
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    conn.Close();
                }
                return true;
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "UpdateAgendaFilesPageCount", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return false;
        }

        #endregion

        #region "Get Annuxer Documents"

        public Meeting GetAnnuxerDocForReplace(Int32 companyId, String meetingId, String agendaId, String annuxerId)
        {
            try
            {
                Meeting objMeeting = new Meeting();
                objMeeting.ID = Convert.ToInt32(meetingId);
                objMeeting.agendaItems = new List<AgendaItems>
                {
                    new AgendaItems
                    {
                        ID = Convert.ToInt32(agendaId)
                    }
                };

                List<AgendaAnnuxers> lstAgendaAnnuxers = null;
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS_ANNUXERS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@MODE", "GET_ANNUXER_DOCUMENTS_FOR_PDF_CONVERSION"));
                        cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(new SqlParameter("@AGENDA_ID", agendaId));
                        cmd.Parameters.Add(new SqlParameter("@ANNUXERS_ID_CSV", annuxerId));
                        SqlDataReader rdr = cmd.ExecuteReader();
                        if (rdr.HasRows)
                        {
                            lstAgendaAnnuxers = new List<AgendaAnnuxers>();
                            while (rdr.Read())
                            {
                                AgendaAnnuxers obj = new AgendaAnnuxers();
                                obj.ID = Convert.ToInt32(rdr["ANNUXERS_ID"]);
                                obj.annuxerDoc = (!String.IsNullOrEmpty(Convert.ToString(rdr["ANNUXERS_DOC"]))) ? Convert.ToString(rdr["ANNUXERS_DOC"]) : String.Empty;
                                lstAgendaAnnuxers.Add(obj);
                            }
                            if (lstAgendaAnnuxers != null)
                            {
                                if (lstAgendaAnnuxers.Count > 0)
                                {
                                    objMeeting.agendaItems[0].agendaAnnuxers = lstAgendaAnnuxers;
                                }
                            }

                        }
                        rdr.Close();
                    }
                    conn.Close();
                    return objMeeting;
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAnnuxerDocForReplace", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return null;
        }

        #endregion

        #region "Update Annuxer Page Count"

        public bool UpdateAnnuxerFilesPageCount(Meeting objMeeting, Int32 companyId)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS_ANNUXERS", conn))
                    {
                        if (objMeeting.agendaItems != null)
                        {
                            if (objMeeting.agendaItems.Count > 0)
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.CommandTimeout = 0;
                                cmd.Parameters.Clear();
                                cmd.Parameters.Add(new SqlParameter("@MODE", "UPDATE_PAGE_COUNT"));
                                cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                                cmd.Parameters.Add(new SqlParameter("@XML_DATA", CreateXmlAnnuxerPageCounts(objMeeting, companyId)));
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    conn.Close();
                }
                return true;
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "UpdateAnnuxerFilesPageCount", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return false;
        }

        #endregion

        #region "Get Withdrawn Agenda Item Documents"

        public Meeting GetWithdrawnAgendaDocs(Int32 companyId, String meetingId, String agendaId)
        {
            Meeting objMeeting = new Meeting();
            List<AgendaItems> lstAgendaItems = null;
            try
            {
                objMeeting.ID = Convert.ToInt32(meetingId);
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@MODE", "GET_WITHDRAWN_ITEM_DOCUMENTS"));
                        cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(new SqlParameter("@MEETING_ID", meetingId));
                        cmd.Parameters.Add(new SqlParameter("@ID", agendaId));
                        SqlDataReader rdr = cmd.ExecuteReader();
                        if (rdr.HasRows)
                        {
                            lstAgendaItems = new List<AgendaItems>();
                            while (rdr.Read())
                            {
                                AgendaItems obj = new AgendaItems();
                                obj.ID = Convert.ToInt32(rdr["AGENDA_ID"]);
                                obj.agendaDoc = (!String.IsNullOrEmpty(Convert.ToString(rdr["AGENDA_DOC"]))) ? Convert.ToString(rdr["AGENDA_DOC"]) : String.Empty;
                                lstAgendaItems.Add(obj);
                            }
                        }
                        rdr.Close();
                        if (lstAgendaItems != null)
                        {
                            if (lstAgendaItems.Count > 0)
                            {
                                objMeeting.agendaItems = lstAgendaItems;
                            }
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAgendaDocs", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return objMeeting;
        }

        #endregion

        #region "Get Withdrawn Annuxer Documents"

        public Meeting GetWithdrawnAnnuxerDocs(Int32 companyId, String meetingId, String agendaId, String annuxerId)
        {
            try
            {
                Meeting objMeeting = new Meeting();
                objMeeting.ID = Convert.ToInt32(meetingId);
                objMeeting.agendaItems = new List<AgendaItems>
                {
                    new AgendaItems
                    {
                        ID = Convert.ToInt32(agendaId)
                    }
                };
                List<AgendaAnnuxers> lstAgendaAnnuxers = null;
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS_ANNUXERS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@MODE", "GET_WITHDRAWN_ANNUXER_DOCUMENTS"));
                        cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(new SqlParameter("@MEETING_ID", meetingId));
                        cmd.Parameters.Add(new SqlParameter("@AGENDA_ID", agendaId));
                        //parameters[4] = new SqlParameter("@ID", annuxerId);                        
                        SqlDataReader rdr = cmd.ExecuteReader();
                        if (rdr.HasRows)
                        {
                            lstAgendaAnnuxers = new List<AgendaAnnuxers>();
                            while (rdr.Read())
                            {
                                AgendaAnnuxers obj = new AgendaAnnuxers();
                                obj.ID = Convert.ToInt32(rdr["ANNUXERS_ID"]);
                                obj.annuxerDoc = (!String.IsNullOrEmpty(Convert.ToString(rdr["ANNUXERS_DOC"]))) ? Convert.ToString(rdr["ANNUXERS_DOC"]) : String.Empty;
                                obj.pageFrom = Convert.ToInt32(rdr["PAGE_FROM"]);
                                obj.pageTo = Convert.ToInt32(rdr["PAGE_TO"]);
                                lstAgendaAnnuxers.Add(obj);
                            }
                            if (lstAgendaAnnuxers != null)
                            {
                                if (lstAgendaAnnuxers.Count > 0)
                                {
                                    objMeeting.agendaItems[0].agendaAnnuxers = lstAgendaAnnuxers;
                                }
                            }
                        }
                        rdr.Close();
                    }
                    conn.Close();
                    return objMeeting;
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAnnuxerDocForReplace", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return null;
        }

        #endregion

        #region "Create Object XML to Update Agenda Page Count"

        private String CreateXmlAgendaPageCounts(Meeting objMeeting, Int32 companyId)
        {
            String xmlStr = String.Empty;
            try
            {
                xmlStr += "<MEETING>";
                foreach (AgendaItems objAgendaItems in objMeeting.agendaItems)
                {
                    xmlStr += "<AGENDA>";
                    xmlStr += "<MEETINGID>" + objMeeting.ID + "</MEETINGID>";
                    xmlStr += "<AGENDAID>" + objAgendaItems.ID + "</AGENDAID>";
                    xmlStr += "<AGENDAPAGEFROM>" + objAgendaItems.pageFrom + "</AGENDAPAGEFROM>";
                    xmlStr += "<AGENDAPAGETO>" + objAgendaItems.pageTo + "</AGENDAPAGETO>";
                    xmlStr += "</AGENDA>";
                }
                xmlStr += "</MEETING>";
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "CreateXmlAgendaPageCounts", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return xmlStr;
        }

        #endregion

        #region "Create Object XML to Update Annuxer Page Count"

        private String CreateXmlAnnuxerPageCounts(Meeting objMeeting, Int32 companyId)
        {
            String xmlStr = String.Empty;
            try
            {
                xmlStr += "<AGENDA>";
                foreach (AgendaAnnuxers objAgendaAnnuxers in objMeeting.agendaItems[0].agendaAnnuxers)
                {
                    xmlStr += "<ANNUXER>";
                    xmlStr += "<AGENDAID>" + objMeeting.agendaItems[0].ID + "</AGENDAID>";
                    xmlStr += "<ANNUXERID>" + objAgendaAnnuxers.ID + "</ANNUXERID>";
                    xmlStr += "<ANNUXERPAGEFROM>" + objAgendaAnnuxers.pageFrom + "</ANNUXERPAGEFROM>";
                    xmlStr += "<ANNUXERPAGETO>" + objAgendaAnnuxers.pageTo + "</ANNUXERPAGETO>";
                    xmlStr += "</ANNUXER>";
                }
                xmlStr += "</AGENDA>";
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "CreateXmlAnnuxerPageCounts", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return xmlStr;
        }

        #endregion

        #region "Get Agenda Supporting Document"

        public Meeting GetAgendaSupportingDoc(Int32 companyId, String meetingId, String agendaId, String documentId)
        {
            Meeting objMeeting = new Meeting();
            objMeeting.ID = Convert.ToInt32(meetingId);
            List<AgendaItems> lstAgendaItems = new List<AgendaItems>();
            AgendaItems objAgendaItems = new AgendaItems();
            objAgendaItems.ID = Convert.ToInt32(agendaId);
            List<AgendaItemSupportingDocument> listDoc = new List<AgendaItemSupportingDocument>();
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@MODE", "GET_ITEM_SUPPORTING_DOC_TO_CONVERT"));
                        cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(new SqlParameter("@COMPANY_ID", companyId));
                        cmd.Parameters.Add(new SqlParameter("@MEETING_ID", meetingId));
                        cmd.Parameters.Add(new SqlParameter("@ID", agendaId));
                        cmd.Parameters.Add(new SqlParameter("@SUPPORTING_DOC_ID", documentId));
                        SqlDataReader rdr = cmd.ExecuteReader();
                        if (rdr.HasRows)
                        {
                            while (rdr.Read())
                            {
                                AgendaItemSupportingDocument obj = new AgendaItemSupportingDocument();
                                obj.ID = Convert.ToInt32(rdr["ID"]);
                                obj.documentName = (!String.IsNullOrEmpty(Convert.ToString(rdr["PDF_DOC_NAME"]))) ? Convert.ToString(rdr["PDF_DOC_NAME"]) : String.Empty;
                                obj.originalDocumentName = (!String.IsNullOrEmpty(Convert.ToString(rdr["ORIGINAL_DOC_NAME"]))) ? Convert.ToString(rdr["ORIGINAL_DOC_NAME"]) : String.Empty;
                                objAgendaItems.agendaDoc = (!String.IsNullOrEmpty(Convert.ToString(rdr["AGENDA_DOC"]))) ? Convert.ToString(rdr["AGENDA_DOC"]) : String.Empty;
                                objAgendaItems.isItemMeredTosupportingDocument = (!String.IsNullOrEmpty(Convert.ToString(rdr["IS_MERGE_ITEM_TO_SUPPORTING_DOC"]))) ? (Convert.ToString(rdr["IS_MERGE_ITEM_TO_SUPPORTING_DOC"]) == "1" ? true : false) : false;
                                listDoc.Add(obj);
                            }
                            objAgendaItems.listSupportingDocument = listDoc;
                        }
                        rdr.Close();

                        lstAgendaItems.Add(objAgendaItems);
                        objMeeting.agendaItems = lstAgendaItems;
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAgendaSupportingDoc", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return objMeeting;
        }

        #endregion

        #region "Update Item Merged Or Not"

        public bool UpdateItemMergedOrNot(Int32 companyId, Int32 meetingId, Int32 agendaId)
        {
            Meeting objMeeting = new Meeting();
            objMeeting.ID = Convert.ToInt32(meetingId);
            List<AgendaItems> lstAgendaItems = new List<AgendaItems>();
            AgendaItems objAgendaItems = new AgendaItems();
            objAgendaItems.ID = Convert.ToInt32(agendaId);
            List<AgendaItemSupportingDocument> listDoc = new List<AgendaItemSupportingDocument>();
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@MODE", "UPDATE_ITEM_TO_BE_MERGED_OR_NOT"));
                        cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(new SqlParameter("@MEETING_ID", meetingId));
                        cmd.Parameters.Add(new SqlParameter("@ID", agendaId));
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
                return true;
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "UpdateItemMergedOrNot", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return false;
        }

        #endregion

        #region "Get Agenda Item Documents To Read FileText"

        public Meeting GetAgendaDocsToReadFileText(Int32 companyId, String meetingId, String agendaId)
        {
            Meeting objMeeting = new Meeting();
            List<AgendaItems> lstAgendaItems = null;
            try
            {
                objMeeting.ID = Convert.ToInt32(meetingId);
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@MODE", "GET_ITEM_DOCUMENTS_TO_READ_FILETEXT"));
                        cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add(new SqlParameter("@MEETING_ID", meetingId));
                        cmd.Parameters.Add(new SqlParameter("@AGENDA_ID_CSV", agendaId));
                        SqlDataReader rdr = cmd.ExecuteReader();
                        if (rdr.HasRows)
                        {
                            lstAgendaItems = new List<AgendaItems>();
                            while (rdr.Read())
                            {
                                AgendaItems obj = new AgendaItems();
                                obj.ID = Convert.ToInt32(rdr["AGENDA_ID"]);
                                obj.agendaDoc = (!String.IsNullOrEmpty(Convert.ToString(rdr["AGENDA_DOC"]))) ? Convert.ToString(rdr["AGENDA_DOC"]) : String.Empty;
                                lstAgendaItems.Add(obj);
                            }
                        }
                        rdr.Close();

                        if (lstAgendaItems != null)
                        {
                            if (lstAgendaItems.Count > 0)
                            {
                                objMeeting.agendaItems = lstAgendaItems;
                            }
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAgendaDocsToReadFileText", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return objMeeting;
        }

        #endregion

        #region "Insert Publish Items Text IntoSQL"

        public bool InsertPublishItemsTextIntoSQL(Meeting objMeeting)
        {
            bool status = false;
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.ChangeDatabase(dbName);
                    using (SqlCommand cmd = new SqlCommand("SP_PROCS_BMS_MEETING_ITEMS", conn))
                    {                        
                        if (objMeeting.agendaItems != null)
                        {
                            if (objMeeting.agendaItems.Count > 0)
                            {
                                foreach (AgendaItems objAgenda in objMeeting.agendaItems)
                                {
                                    cmd.CommandType = CommandType.StoredProcedure;
                                    cmd.CommandTimeout = 0;
                                    cmd.Parameters.Clear();
                                    cmd.Parameters.Add(new SqlParameter("@MODE", "UPDATE_PUBLISH_ITEMS_FILE_CONTENT"));
                                    cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                                    cmd.Parameters.Add(new SqlParameter("@MEETING_ID", objMeeting.ID));
                                    cmd.Parameters.Add(new SqlParameter("@ID", objAgenda.ID));
                                    cmd.Parameters.Add(new SqlParameter("@CONTENT", objAgenda.fileContent));
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                    conn.Close();
                }
                status = true;
            }
            catch (Exception ex)
            {
                status = false;
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "MeetingRepository", "InsertPublishItemsTextIntoSQL", "FilesToPDFConvertor Scheduler", 1, objMeeting.companyId);
            }
            return status;
        }

        #endregion
    }
}