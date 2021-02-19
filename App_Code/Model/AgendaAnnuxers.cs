using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FilesToPDFConvertor
{
    public class AgendaAnnuxers : FilesToPDFConvertor_BaseEntity
    {
        public Int32 ID { get; set; }
        public String annuxerTitle { get; set; }
        public Int32 annuxerOrder { get; set; }              
        public String annuxerDoc { get; set; }
        public String status { get; set; }        
        public Int32 pageFrom { get; set; }
        public Int32 pageTo { get; set; }
        public String originalDocumentName { get; set; }
        //public User user { get; set; }
        public String createdBy { get; set; }
        public String createdOn { get; set; }
        public String modifiedBy { get; set; }
        public String modifiedOn { get; set; }

        public override void Validate()
        {
            base.Validate();
        }
    }
}