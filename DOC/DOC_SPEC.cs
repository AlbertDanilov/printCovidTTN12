using System;

namespace DOC
{
    public class DOC_SPEC
    {
        public long? avs_id { get; set; }
        public string ts_temp_regim { get; set; }
        public string ts_ed_shortname { get; set; }
        public string ts_shifr { get; set; }
        public string ts_seria { get; set; }
        public DateTime? ts_sgod { get; set; }
        public string ts_sert { get; set; }
        public DateTime? ts_sert_date_po { get; set; }
        public DateTime? ts_sert_date_s { get; set; }
        public string ts_okp { get; set; }
        public string ts_p_tn { get; set; }
        public string ts_p_fv_doz { get; set; }
        public string ts_p_proizv { get; set; }
        public int? ts_sgtin_cnt { get; set; }
        public decimal? pvs_psum_bnds { get; set; }
        public decimal? pvs_rsum_nds { get; set; }
        public decimal? pvs_psum_nds { get; set; }
        public decimal? pvs_kol_tov { get; set; }
        public decimal? ts_pcena_bnds { get; set; }
        public decimal? ts_pcena_nds { get; set; }
        public decimal? ts_ocena_nds { get; set; }
        public decimal? ts_osum_nds { get; set; }
        public decimal? ts_nds_i_val { get; set; }
    }
}
