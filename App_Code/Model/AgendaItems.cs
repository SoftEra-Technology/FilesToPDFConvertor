using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FilesToPDFConvertor
{
    public class AgendaItems : FilesToPDFConvertor_BaseEntity
    {
        public Int32 ID { get; set; }
        public String agendaTitle { get; set; }      
        public String description { get; set; }
        public Int32 agendaOrder { get; set; }
        public String agendaDoc { get; set; }
        public String status { get; set; }
        public String publishedOn { get; set; }
        public Int32 pageFrom { get; set; }
        public Int32 pageTo { get; set; }        
        public List<AgendaAnnuxers> agendaAnnuxers { get; set; }
        public String originalDocumentName { get; set; }
        public bool isOnlyAddAnnuxers { get; set; }
        public bool isOnlyDeleteAnnuxers { get; set; }
        //public User user { get; set; }
        public String createdBy { get; set; }
        public String createdOn { get; set; }
        public String modifiedBy { get; set; }
        public String modifiedOn { get; set; }
        public List<AgendaItemSupportingDocument> listSupportingDocument { get; set; }
        public bool isItemMeredTosupportingDocument { get; set; }
        public String fileContent { get; set; }
        public override void Validate()
        {
            base.Validate();
        }
    }
}