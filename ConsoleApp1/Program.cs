using DOC;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace ConsoleApp1
{
    public class Program
    {
        private static string ConnectionSting = "Data Source=192.168.0.35;Initial Catalog=rsklad;User ID=sa;Password=r12sql141007";

        static void Main(string[] args)
        {
            DataTable dt = new DataTable("T");
            List<DOCS> docItems = new List<DOCS>();

            try
            {
                using (var con = new SqlConnection(ConnectionSting))
                {
                    using (var cmd = new SqlCommand("", con))
                    {
                        using (var da = new SqlDataAdapter(cmd))
                        {
                            da.SelectCommand.CommandType = CommandType.StoredProcedure;
                            da.SelectCommand.CommandText = "DOCS_DLO_APT_covid_lpu";
                            da.SelectCommand.Parameters.Clear();
                            da.Fill(dt);
                        }
                    }
                }

                docItems = dt.AsEnumerable()
                           .GroupBy(d => new
                           {
                               av_id = d.Field<int>("av_id"),
                               recipient = d.Field<string>("recipient"),
                               sender = d.Field<string>("sender"),
                               contractor = d.Field<string>("contractor"),
                               agent_okpo = d.Field<string>("agent_okpo"),
                               doc_nom = d.Field<string>("doc_nom"),
                               work_porgram = d.Field<string>("work_porgram"),
                               work_program_id = d.Field<short>("work_program_id")
                           }).Select(ds => new DOCS
                           {
                               av_id = ds.Key.av_id,
                               recipient = ds.Key.recipient,
                               sender = ds.Key.sender,
                               contractor = ds.Key.contractor,
                               agent_okpo = ds.Key.agent_okpo,
                               doc_nom = ds.Key.doc_nom,
                               work_porgram = ds.Key.work_porgram,
                               work_program_id = ds.Key.work_program_id,
                               ds_list = ds.GroupBy(dss => new
                               {
                                   avs_id = dss.Field<long>("avs_id"),
                                   ts_temp_regim = dss.Field<string>("ts_temp_regim"),
                                   ts_ed_shortname = dss.Field<string>("ts_ed_shortname"),
                                   ts_shifr = dss.Field<string>("ts_shifr"),
                                   ts_seria = dss.Field<string>("ts_seria"),
                                   ts_sgod = dss.Field<DateTime>("ts_sgod"),
                                   ts_sert = dss.Field<string>("ts_sert"),
                                   ts_sert_date_po = dss.Field<DateTime>("ts_sert_date_po"),
                                   ts_sert_date_s = dss.Field<DateTime>("ts_sert_date_s"),
                                   ts_okp = dss.Field<string>("ts_okp"),
                                   ts_p_tn = dss.Field<string>("ts_p_tn"),
                                   ts_p_fv_doz = dss.Field<string>("ts_p_fv_doz"),
                                   ts_p_proizv = dss.Field<string>("ts_p_proizv"),
                                   ts_sgtin_cnt = dss.Field<int>("ts_sgtin_cnt"),
                                   pvs_psum_bnds = dss.Field<decimal>("pvs_psum_bnds"),
                                   pvs_rsum_nds = dss.Field<decimal>("pvs_rsum_nds"),
                                   pvs_psum_nds = dss.Field<decimal>("pvs_psum_nds"),
                                   pvs_kol_tov = dss.Field<decimal>("pvs_kol_tov"),
                                   ts_pcena_bnds = dss.Field<decimal>("ts_pcena_bnds"),
                                   ts_pcena_nds = dss.Field<decimal>("ts_pcena_nds"),
                                   ts_ocena_nds = dss.Field<decimal>("ts_ocena_nds"),
                                   ts_osum_nds = dss.Field<decimal>("ts_osum_nds"),
                                   ts_nds_i_val = dss.Field<decimal>("ts_nds_i_val")
                               }).Select(dss => new DOC_SPEC
                               {
                                   avs_id = dss.Key.avs_id,
                                   ts_temp_regim = dss.Key.ts_temp_regim,
                                   ts_ed_shortname = dss.Key.ts_ed_shortname,
                                   ts_shifr = dss.Key.ts_shifr,
                                   ts_seria = dss.Key.ts_seria,
                                   ts_sgod = dss.Key.ts_sgod,
                                   ts_sert = dss.Key.ts_sert,
                                   ts_sert_date_po = dss.Key.ts_sert_date_po,
                                   ts_sert_date_s = dss.Key.ts_sert_date_s,
                                   ts_okp = dss.Key.ts_okp,
                                   ts_p_tn = dss.Key.ts_p_tn,
                                   ts_p_fv_doz = dss.Key.ts_p_fv_doz,
                                   ts_p_proizv = dss.Key.ts_p_proizv,
                                   ts_sgtin_cnt = dss.Key.ts_sgtin_cnt,
                                   pvs_psum_bnds = dss.Key.pvs_psum_bnds,
                                   pvs_rsum_nds = dss.Key.pvs_rsum_nds,
                                   pvs_psum_nds = dss.Key.pvs_psum_nds,
                                   pvs_kol_tov = dss.Key.pvs_kol_tov,
                                   ts_pcena_bnds = dss.Key.ts_pcena_bnds,
                                   ts_pcena_nds = dss.Key.ts_pcena_nds,
                                   ts_ocena_nds = dss.Key.ts_ocena_nds,
                                   ts_osum_nds = dss.Key.ts_osum_nds,
                                   ts_nds_i_val = dss.Key.ts_nds_i_val
                               }).ToList()
                           }).ToList();

                //string jsonDocItems = Utf8Json.JsonSerializer.ToJsonString(docItems);

                foreach (DOCS doc in docItems)
                {
                    Print.Print.PrintExcel(doc);
                    //if (doc.av_id.Value == 57497)
                    //{
                    //    Print.Print.PrintExcel(doc);
                    //}
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}