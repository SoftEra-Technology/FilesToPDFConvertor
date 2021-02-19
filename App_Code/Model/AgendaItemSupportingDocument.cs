using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FilesToPDFConvertor
{
    public class AgendaItemSupportingDocument : FilesToPDFConvertor_BaseEntity
    {
        public Int32 ID { get; set; }
        public String documentName { get; set; }      
        public String originalDocumentName { get; set; }     
        public override void Validate()
        {
            base.Validate();
        }
    }
}