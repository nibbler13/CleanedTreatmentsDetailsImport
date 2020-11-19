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
			"and treat_nctrdate between @dateBegin and @dateEnd " +
			"and doctype = 11";

		public static string sqlGetDataTreatmentsDetailsOtherIc =
			"select  " +
			"	mfo.ordtid,  " +
			"	mfo.treatcode,  " +
			"	mfo.treat_nctrdate,  " +
			"	mfo.schid,  " +
			"	mfo.tooth_name,  " +
			"	cl.fullname as client,  " +
			"	mfoc.etl_insert_time,  " +
			"	mfoc.file_info,  " +
			"	mfoc.loadingUserName,  " +
			"   mfo.schcount,  " +
			"   mfo.schamount  " +
			"from mis_fact_orderdet mfo  " +
			"	join mis_dim_clients cl on cl.pcode = mfo.pcode " +
			"	join mis_dim_wschema mdw on mdw.schid = mfo.schid   " +
			"	left join mis_fact_orderdet_cleaned mfoc on mfoc.ordtid = mfo.ordtid " +
			"where mfo.treat_nctrdate between @dateBegin and @dateEnd " +
			"	and (mfo.jpagreement_agrid not in ( " +
			"		991515382, 991517927, 991523042, 100005,     991519409, 991519865, " +
			"    	991523030, 991520911, 991514852, 991511535,  991519440, 991521374, " +
			"    	991511568, 991520964, 991526106, 991523453,  991517214, 991520348,  " +
			"    	991522348, 991522924, 991525955, 991522442,  991522386, 991517912, " +
			"    	991523451, 991519436, 991523280, 991523170,  991518370, 991523038, " +
			"    	991526075, 991519595, 991511705, 1990097479, 991516698, 991521960,  " +
			"    	991524638, 991520913, 991518470, 991519761,  991523028, 991516556, " +
			"    	991523215, 991519361, 991525970, 991515797,  991520427, 991519358, " +
			"    	991520387, 991520427, 991520911, 991520913,  991520499, 991521272) " +
			"	and mfo.filid in (1, 5, 6, 7, 12)) " +
			"	and coalesce(mfo.jplists_sectid, 0) not in (4365, 991139394) " +
			"	and coalesce(mfo.jpagreement_agrtype , '') not in ('ОМС','Франшиза','Профосмотр','Предрейсовый осмотр','Предстраховое обследование','Вакцинация','СМП') " +
			"	and mdw.kodoper not like '500%' " +
			"	and mdw.kodoper not like '400%' " +
			"	and mdw.kodoper not in ( " +
			"		'101898',  '1002285', '1002286', '1002287', '2110635', " +
			"		'2110636', '2110637', '326217',  '1002285', '1002286', " +
			"		'1002287', '2110635', '2110636', '2110637', '101899', " +
			"		'212066',  '326217',  '101939',  '101940',  '290500', " +
			"		'290501',  '290502',  '290503',  '290504',  '290505', " +
			"		'290506',  '290507',  '290508',  '290509',  '290510', " +
			"		'290511',  '290512',  '290513',  '290514',  '290515', " +
			"		'290516',  '290517',  '290518',  '290519',  '290520', " +
			"		'290521') " +
			"	and coalesce(mfo.schamount_a, 0) <> 0 " +
			"	and mfo.doctype = 11 " +
			"union all " +
			"select  " +
			"	mfo.ordtid,  " +
			"	mfo.treatcode,  " +
			"	mfo.treat_nctrdate,  " +
			"	mfo.schid,  " +
			"	mfo.tooth_name,  " +
			"	cl.fullname as client,  " +
			"	mfoc.etl_insert_time,  " +
			"	mfoc.file_info,  " +
			"	mfoc.loadingUserName,  " +
			"   mfo.schcount,  " +
			"   mfo.schamount  " +
			"from mis_fact_orderdet mfo  " +
			"	join mis_dim_clients cl on cl.pcode = mfo.pcode " +
			"	join mis_dim_wschema mdw on mdw.schid = mfo.schid   " +
			"	left join mis_fact_orderdet_cleaned mfoc on mfoc.ordtid = mfo.ordtid " +
			"where  " +
			"	mfo.treat_nctrdate between @dateBegin and @dateEnd " +
			"	and ( " +
			"		mfo.jplists_lid in ( " +
			"			991525444, 991525808, 991527454, 991525444, 991523055, 991527454, " +
			"			991523054, 991518567, 991520526, 991517931, 991520527, 991524157, " +
			"			991517234, 991516661, 991513769, 991527312, 991525939, 991525750, " +
			"			991519141, 991525424, 991522797) " +
			"		or mfo.jpagreement_agrid in  (991526592,991526634)) " +
			"	and mfo.filid in (1, 5, 6, 7, 12) " +
			"	and coalesce(mfo.jplists_sectid, 0) not in (4365) " +
			"	and coalesce(mfo.jpagreement_agrtype, '') not in ('ОМС','Профосмотр','Предрейсовый осмотр','Предстраховое обследование','Вакцинация','СМП') " +
			"	and mdw.kodoper not like '500%' " +
			"	and mdw.kodoper not like '400%' " +
			"	and mdw.kodoper not in ( " +
			"		'101898',  '1002285', '1002286', '1002287', '2110635', " +
			"		'2110636', '2110637', '326217',  '1002285', '1002286', " +
			"		'1002287', '2110635', '2110636', '2110637', '101899', " +
			"		'212066',  '326217',  '101939',  '101940',  '290500', " +
			"		'290501',  '290502',  '290503',  '290504',  '290505', " +
			"		'290506',  '290507',  '290508',  '290509',  '290510', " +
			"		'290511',  '290512',  '290513',  '290514',  '290515', " +
			"		'290516',  '290517',  '290518',  '290519',  '290520', " +
			"		'290521') " +
			"	and coalesce(mfo.schamount_a, 0) <> 0 " +
			"	and mfo.doctype = 11 ";

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
			"    loadingUserName, " +
			"    average_discount, " +
			"    amount_total_with_average_discount " +
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
			"    @loadingUserName, " +
			"    @average_discount, " +
			"    @amount_total_with_average_discount " +
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
