using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Vertica.Data.VerticaClient;
using System.Data;
using System.Runtime.Remoting.Messaging;
using System.ComponentModel;
using System.Diagnostics;

namespace CleanedTreatmentsDetailsImport {
	public class VerticaClient {
		private readonly VerticaConnection connection;
		private readonly BackgroundWorker bw;

		public VerticaClient(string host, string database, string user, string password, BackgroundWorker bw = null) {
			VerticaConnectionStringBuilder builder = new VerticaConnectionStringBuilder {
				Host = host,
				Database = database,
				User = user,
				Password = password
			};

			this.bw = bw;
			connection = new VerticaConnection(builder.ToString());
			IsConnectionOpened();
		}

		private bool IsConnectionOpened() {
			if (connection.State != ConnectionState.Open) {
				try {
					connection.Open();
				} catch (Exception e) {
					string subject = "Ошибка подключения к БД";
					string body = e.Message + Environment.NewLine + e.StackTrace;
					SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
					Logging.ToLog(subject + " " + body);

					if (bw != null)
						bw.ReportProgress(0, subject + " " + body);
				}
			}

			return connection.State == ConnectionState.Open;
		}

		public DataTable GetDataTable(string query, Dictionary<string, object> parameters = null) {
			DataTable dataTable = new DataTable();

			if (!IsConnectionOpened())
				return dataTable;

			try {
				using (VerticaCommand command = new VerticaCommand(query, connection)) {
					if (parameters != null && parameters.Count > 0)
						foreach (KeyValuePair<string, object> parameter in parameters)
							command.Parameters.Add(new VerticaParameter(parameter.Key, parameter.Value));

					using (VerticaDataAdapter fbDataAdapter = new VerticaDataAdapter(command))
						fbDataAdapter.Fill(dataTable);
				}
			} catch (Exception e) {
				string subject = "Ошибка выполнения запроса к БД";
				string body = e.Message + Environment.NewLine + e.StackTrace;
				SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
				Logging.ToLog(subject + " " + body);
				connection.Close();

				if (bw != null)
					bw.ReportProgress(0, subject + " " + body);
			}

			return dataTable;
		}

		public bool ExecuteUpdateQuery(
						string query, 
						bool isTreatmentsDetails,
						string fileInfo) {
			bool updatedCorrected = true;

			if (!IsConnectionOpened())
				return updatedCorrected;

			string now = DateTime.Now.ToString("yyyyMMddHHmmss");

			using (VerticaTransaction transaction = connection.BeginTransaction()) {
				using (VerticaCommand update = new VerticaCommand(query, connection)) {
					if (isTreatmentsDetails) {
						DataTable dataTable = Program.FileContentTreatmentsDetails;

						if (dataTable == null)
							return false;

						for (int i = 0; i < dataTable.Rows.Count; i++) {
							try {
								update.Parameters.Clear();
								bool skip = false;

								foreach (Program.Header header in Program.headers) {
									if (Debugger.IsAttached)
										if (header.DbField == "ordtid" && string.IsNullOrEmpty(dataTable.Rows[i][header.DbField].ToString())) {
											skip = true;
											break;
										}

									update.Parameters.Add(new VerticaParameter(header.DbField, dataTable.Rows[i][header.DbField]));
								}

								if (skip)
									continue;

								update.Parameters.Add(new VerticaParameter("@etl_pipeline_id", "CleanedTreatmentsDetailsImport" + "_" + now));
								update.Parameters.Add(new VerticaParameter("@file_info", fileInfo));
								update.Parameters.Add(new VerticaParameter("@loadingUserName", Environment.UserName + "@" + Environment.MachineName));

								if (update.ExecuteNonQuery() == 0)
									updatedCorrected = false;
							} catch (Exception e) {
								string subject = "Ошибка выполнения запроса к БД";
								string body = e.Message + Environment.NewLine + e.StackTrace;
								SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
								Logging.ToLog(subject + " " + body);

								if (bw != null)
									bw.ReportProgress(0, subject + " " + body);

								Logging.ToLog("---Исходные данные:");
								foreach (Program.Header header in Program.headers)
									Logging.ToLog(header.DbField + " | " + (dataTable.Rows[i][header.DbField] == null ? "null" : dataTable.Rows[i][header.DbField].ToString()));

								transaction.Rollback();
								connection.Close();

								return false;
							}
						}
					} else {
						foreach (ItemProfitAndLoss item in Program.FileContentProfitAndLoss) {
							try {
								update.Parameters.Clear();
								update.Parameters.Add(new VerticaParameter("@object_name", item.ObjectName));
								update.Parameters.Add(new VerticaParameter("@period_year", item.PeriodYear));
								update.Parameters.Add(new VerticaParameter("@period_type", item.PeriodType));
								update.Parameters.Add(new VerticaParameter("@group_name_level_1", item.GroupNameLevel1));
								update.Parameters.Add(new VerticaParameter("@group_name_level_2", item.GroupNameLevel2));
								update.Parameters.Add(new VerticaParameter("@group_name_level_3", item.GroupNameLevel3));
								update.Parameters.Add(new VerticaParameter("@value", item.Value));
								update.Parameters.Add(new VerticaParameter("@group_sorting_order", item.GroupSortingOrder));
								update.Parameters.Add(new VerticaParameter("@object_sorting_order", item.ObjectSrotingOrder));
								update.Parameters.Add(new VerticaParameter("@quarter", item.Quarter));
								update.Parameters.Add(new VerticaParameter("@has_data", item.HasData));
								update.Parameters.Add(new VerticaParameter("@etl_pipeline_id", "CleanedTreatmentsDetailsImport" + "_" + now));
								update.Parameters.Add(new VerticaParameter("@file_info", fileInfo));
								update.Parameters.Add(new VerticaParameter("@loadingUserName", Environment.UserName + "@" + Environment.MachineName));

								if (update.ExecuteNonQuery() == 0)
									updatedCorrected = false;
							} catch (Exception e) {
								string subject = "Ошибка выполнения запроса к БД";
								string body = e.Message + Environment.NewLine + e.StackTrace;
								if (e.InnerException != null)
									body += Environment.NewLine + e.InnerException.Message + Environment.NewLine + e.InnerException.StackTrace;

								SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
								Logging.ToLog(subject + " " + body);

								if (bw != null)
									bw.ReportProgress(0, subject + " " + body);

								Logging.ToLog("---Исходные данные:");
								Logging.ToLog(item.ToString());

								transaction.Rollback();
								connection.Close();

								return false;
							}
						}
					}
				}

				transaction.Commit();
			}

			return updatedCorrected;
		}
	}
}
