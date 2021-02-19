using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FilesToPDFConvertor
{
    public class Meeting : FilesToPDFConvertor_BaseEntity
    {
        public Int32 ID {get;set;}
        public String meetingTitle { get; set; }
        public Int32 meetingNumber { get; set; }
        public String meetingDate { get; set; }
        public String timeFrom { get; set; }
        public String timeTo { get; set; }
        public String meetingStatus { get; set; }       
        public String boardBook { get; set; }               
        public Int32 membersCount { get; set; }
        public List<AgendaItems> agendaItems { get; set; }
        public Int32 uploadKey { get; set; }
        public bool isUpload { get; set; }
        public String encryptId { get; set; }
        public String databaseName { set; get; }
        public Int32 companyId { get; set; }
        public String createdBy { get; set; }
        public String createdOn { get; set; }
        public String modifiedBy { get; set; }
        public String modifiedOn { get; set; }
        public String mtyear { get; set; }
        public override void Validate()
        {
            base.Validate();
        }
    }
}