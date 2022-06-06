using System.Collections.Generic;

namespace DOC
{
    public class DOCS
    {
        public int? av_id { get; set; }
        public string recipient { get; set; }
        public string sender { get; set; }
        public string contractor { get; set; }
        public string agent_okpo { get; set; }
        public string doc_nom { get; set; }
        public string work_porgram { get; set; }
        public short work_program_id { get; set; }
        public List<DOC_SPEC> ds_list { get; set; }
    }
}
