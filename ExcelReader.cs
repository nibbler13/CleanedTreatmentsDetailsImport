using CleanedTreatmentsDetailsImport;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CleanedTreatmentsDetailsImport {
	public class ExcelReader {
		static public DataTable ReadExcelFile(string fileName, string sheetName) {
			Logging.ToLog("Считывание книги: " + fileName + ", лист: " + sheetName);
			DataTable dataTable = new DataTable();

			if (!File.Exists(fileName))
				return dataTable;

			try {
				using (OleDbConnection conn = new OleDbConnection()) {
					conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Mode=Read;" +
						"Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1'";

					using (OleDbCommand comm = new OleDbCommand()) {
						if (string.IsNullOrEmpty(sheetName)) {
							conn.Open();
							DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
								new object[] { null, null, null, "TABLE" });
							sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
							conn.Close();
						} else
							sheetName += "$";

#pragma warning disable CA2100 // Review SQL queries for security vulnerabilities
						comm.CommandText = "Select * from [" + sheetName + "]";
#pragma warning restore CA2100 // Review SQL queries for security vulnerabilities
						comm.Connection = conn;

						using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter()) {
							oleDbDataAdapter.SelectCommand = comm;
							oleDbDataAdapter.Fill(dataTable);
						}
					}
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			return dataTable;
		}


		public static List<string> ReadSheetNames(string file) {
			List<string> sheetNames = new List<string>();

			using (OleDbConnection conn = new OleDbConnection()) {
				conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Mode=Read;" +
					"Extended Properties='Excel 12.0 Xml;HDR=NO;'";

				using (OleDbCommand comm = new OleDbCommand()) {
					conn.Open();
					DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
						new object[] { null, null, null, "TABLE" });
					foreach (DataRow row in dtSchema.Rows) {
						string name = row.Field<string>("TABLE_NAME");
						if (name.Contains("FilterDatabase"))
							continue;

						sheetNames.Add(name.Replace("$", "").TrimStart('\'').TrimEnd('\''));
					}
				}
			}

			return sheetNames;
		}
	}
}
