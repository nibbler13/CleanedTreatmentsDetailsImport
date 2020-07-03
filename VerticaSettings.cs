using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CleanedTreatmentsDetailsImport {
	public class VerticaSettings {
		public static string host = Properties.Settings.Default.VerticaHost;
		public static string database = Properties.Settings.Default.VerticaDatabase;
		public static string user = Properties.Settings.Default.VerticaUser;
		public static string password = Properties.Settings.Default.VerticaPassword;

		public static string sqlGetDataTreatmentsDetails = 
			"select " +
			"	mfo.ordtid, " +
			"	mfo.treatcode, " +
			"	mfo.treat_nctrdate, " +
			"	mfo.schid, " +
			"	mfo.tooth_name, " +
			"	cl.fullname as client, " +
			"	mfoc.etl_insert_time, " +
			"	mfoc.file_info, " +
			"	mfoc.loadingUserName, " +
			"   mfo.schcount, " +
			"   mfo.schamount " +
			"from mis_fact_orderdet mfo " +
			"left join mis_dim_clients cl on cl.pcode = mfo.pcode " +
			"left join mis_fact_orderdet_cleaned mfoc on mfoc.ordtid = mfo.ordtid " +
			"where mfo.jpagreement_agrid in (@jids) " +
			"and (treat_nctrdate between @dateBegin and @dateEnd) " +
			"and (doctype = 5 or doctype = 11)";

		public static string sqlInsertTreatmentsDetails =
			"insert into public.mis_fact_orderdet_cleaned ( " +
			"    ordtid, " +
			"    contract, " +
			"    program, " +
			"    policy, " +
			"    patient_full_name, " +
			"    patient_birthday, " +
			"    patient_age, " +
			"    patient_histnum, " +
			"    mkb_diagnoses, " +
			"    toothnum, " +
			"    treat_date, " +
			"    kodoper, " +
			"    service_name, " +
			"    service_count, " +
			"    service_cost, " +
			"    amount_total, " +
			"    amount_total_with_discount, " +
			"    filial, " +
			"    department, " +
			"    doctor_full_name, " +
			"    referral_filial, " +
			"    referral_department, " +
			"    referral_doctor_full_name, " +
			"    agreement, " +
			"    service_restrict, " +
			"    accepted_arragement, " +
			"    expert_comment, " +
			"    average_check_exclude_rule, " +
			"    amount_total_left_without_corrections, " +
			"    expert_name, " +
			"    treatcode, " +
			"    contract_type, " +
			"    program_type, " +
			"    schid, " +
			"    dcode, " +
			"    profid, " +
			"    doctor_position_catalog, " +
			"    doctor_position, " +
			"    referral_dcode, " +
			"    referral_profid, " +
			"    referral_doctor_position_catalog, " +
			"    referral_doctor_position, " +
			"    garant_letter, " +
			"    category, " +
			"    etl_pipeline_id, " +
			"    file_info, " +
			"    loadingUserName " +
			") " +
			"values ( " +
			"    @ordtid, " +
			"    @contract, " +
			"    @program, " +
			"    @policy, " +
			"    @patient_full_name, " +
			"    @patient_birthday, " +
			"    @patient_age, " +
			"    @patient_histnum, " +
			"    @mkb_diagnoses, " +
			"    @toothnum, " +
			"    @treat_date, " +
			"    @kodoper, " +
			"    @service_name, " +
			"    @service_count, " +
			"    @service_cost, " +
			"    @amount_total, " +
			"    @amount_total_with_discount, " +
			"    @filial, " +
			"    @department, " +
			"    @doctor_full_name, " +
			"    @referral_filial, " +
			"    @referral_department, " +
			"    @referral_doctor_full_name, " +
			"    @agreement, " +
			"    @service_restrict, " +
			"    @accepted_arragement, " +
			"    @expert_comment, " +
			"    @average_check_exclude_rule, " +
			"    @amount_total_left_without_corrections, " +
			"    @expert_name, " +
			"    @treatcode, " +
			"    @contract_type, " +
			"    @program_type, " +
			"    @schid, " +
			"    @dcode, " +
			"    @profid, " +
			"    @doctor_position_catalog, " +
			"    @doctor_position, " +
			"    @referral_dcode, " +
			"    @referral_profid, " +
			"    @referral_doctor_position_catalog, " +
			"    @referral_doctor_position, " +
			"    @garant_letter, " +
			"    @category, " +
			"    @etl_pipeline_id, " +
			"    @file_info, " +
			"    @loadingUserName " +
			")";

		public static string sqlInsertProfitAndLoss =
			"insert into public.profit_and_loss ( " +
			"    object_name, " +
			"    period_year, " +
			"    period_type, " +
			"    group_name_level_1, " +
			"    group_name_level_2, " +
			"    group_name_level_3, " +
			"    value, " +
			"    group_sorting_order, " +
			"    object_sorting_order, " +
			"    quarter, " +
			"    has_data, " +
			"    etl_pipeline_id, " +
			"    file_info, " +
			"    loading_user_name " +
			") " +
			"values ( " +
			"    @object_name, " +
			"    @period_year, " +
			"    @period_type, " +
			"    @group_name_level_1, " +
			"    @group_name_level_2, " +
			"    @group_name_level_3, " +
			"    @value, " +
			"    @group_sorting_order, " +
			"    @object_sorting_order, " +
			"    @quarter, " +
			"    @has_data, " +
			"    @etl_pipeline_id, " +
			"    @file_info, " +
			"    @loading_user_name " +
			")";

		public static string sqlRefreshOrderdet = "select refresh_columns('mis_fact_orderdet', 'schcount_cleaned, schamount_cleaned', 'rebuild');";
	}
}
