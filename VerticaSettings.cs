using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CleanedTreatmentsDetailsImport {
	class VerticaSettings {
		public static string host = Properties.Settings.Default.VerticaHost;
		public static string database = Properties.Settings.Default.VerticaDatabase;
		public static string user = Properties.Settings.Default.VerticaUser;
		public static string password = Properties.Settings.Default.VerticaPassword;

		public static string sqlInsert = 
			"insert into public.treatments_details_cleaned ( " +
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
			"    file_info " +
			") " +
			"values ( " +
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
			"    @file_info " +
			")";
	}
}
