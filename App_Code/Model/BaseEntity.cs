using System;
using System.Collections.Generic;
using System.Web;

namespace FilesToPDFConvertor
{
    public class FilesToPDFConvertor_BaseEntity
    {
        private List<String> _brokenRules = new List<String>();
        public String moduleDatabase { get; set; }
        public virtual void Validate() { }
        public void ClearRules()
        {
            _brokenRules.Clear();
        }
        public void AddRule(String rule)
        {
            _brokenRules.Add(rule);
        }
        public List<String> GetRules()
        {
            return _brokenRules;
        }
    }
}