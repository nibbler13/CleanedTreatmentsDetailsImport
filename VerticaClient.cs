using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Vertica.Data.VerticaClient;
using System.Data;

namespace CleanedTreatmentsDetailsImport {
	class VerticaClient {
		private VerticaConnection connection;

		public VerticaClient(string host, string database, string user, string password) {
			VerticaConnectionStringBuilder builder = new VerticaConnectionStringBuilder {
				Host = host,
				Database = database,
				User = user,
				Password = password
			};

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
			}

			return dataTable;
		}

		public bool ExecuteUpdateQuery(
						string query, Dictionary<Tuple<string, string, string, Program.FieldType>, 
						List<object>> headersAndData, 
						string fileInfo) {
			if (headersAndData == null)
				return false;

			bool updatedCorrected = true;

			if (!IsConnectionOpened())
				return updatedCorrected;

			string now = DateTime.Now.ToString("yyyyMMddHHmmss");

			using (VerticaTransaction transaction = connection.BeginTransaction()) {
				using (VerticaCommand update = new VerticaCommand(query, connection)) {
					for (int i = 0; i < headersAndData.First().Value.Count; i++) {
						try {
							update.Parameters.Clear();

							foreach (Tuple<string, string, string, Program.FieldType> header in headersAndData.Keys)
								update.Parameters.Add(new VerticaParameter(header.Item3, headersAndData[header][i]));

							update.Parameters.Add(new VerticaParameter("@etl_pipeline_id", "CleanedTreatmentsDetailsImport" + "_" + now));
							update.Parameters.Add(new VerticaParameter("@file_info", fileInfo));

							if (update.ExecuteNonQuery() == 0)
								updatedCorrected = false;
						} catch (Exception e) {
							string subject = "Ошибка выполнения запроса к БД";
							string body = e.Message + Environment.NewLine + e.StackTrace;
							SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
							Logging.ToLog(subject + " " + body);

							Logging.ToLog("---Исходные данные:");
							foreach (Tuple<string, string, string, Program.FieldType> header in headersAndData.Keys)
								Logging.ToLog(header.Item3 + " | " + (headersAndData[header][i] == null ? "null" : headersAndData[header][i].ToString()));

							transaction.Rollback();
							connection.Close();
							return false;
						}
					}
				}

				transaction.Commit();
			}

			return updatedCorrected;
		}
	}
}
