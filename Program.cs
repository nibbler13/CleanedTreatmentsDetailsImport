using ActOfProvidedServices;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace CleanedTreatmentsDetailsImport {
	class Program {
		private static readonly string subjectError = "!!! Ошибка в работе CleanedTreatmentsDetailsImport";
		private static readonly string receiverError = Properties.Settings.Default.MailCopy;
		public enum FieldType { String, Long, Float, Date };

		//Tuple<column, header, field, type>, List<Value>
		private static readonly Dictionary<Tuple<string, string, string, FieldType>, List<object>> headersAndData = new Dictionary<Tuple<string, string, string, FieldType>, List<object>> {
			{ new Tuple<string, string, string, FieldType>("F1", "договор",                                             "contract",								FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F2", "программа",                                           "program",								FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F3", "полис",                                               "policy",								FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F4", "фио пациента",                                        "patient_full_name",					FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F5", "дата рождения",                                       "patient_birthday",						FieldType.Date),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F6", "возраст",                                             "patient_age",							FieldType.Long),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F7", "история болезни",                                     "patient_histnum",						FieldType.Long),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F8", "диагноз по мкб",                                      "mkb_diagnoses",						FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F9", "№ зуба",                                              "toothnum",								FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F10", "дата приема",                                        "treat_date",							FieldType.Date),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F11", "код услуги",                                         "kodoper",								FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F12", "наименование услуги",                                "service_name",							FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F13", "количество услуг",                                   "service_count",						FieldType.Long),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F14", "стоимость услуги",                                   "service_cost",							FieldType.Float),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F15", "сумма всего",                                        "amount_total",							FieldType.Float),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F16", "сумма всего с учетом скидки",                        "amount_total_with_discount",			FieldType.Float),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F17", "филиал",                                             "filial",								FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F18", "отделение",                                          "department",							FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F19", "фио врача",                                          "doctor_full_name",						FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F20", "направление: филиал",                                "referral_filial",						FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F21", "направление: отделение",                             "referral_department",					FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F22", "направление: врач",                                  "referral_doctor_full_name",			FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F23", "согласование",                                       "agreement",							FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F24", "ограничение по услуге",                              "service_restrict",						FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F25", "примечание эксперта",                                "accepted_arragement",					FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F26", "принятые меры",                                      "expert_comment",						FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F27", "не входит с сч",                                     "average_check_exclude_rule",			FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F28", "оставлено без исправлений (сумма)",                  "amount_total_left_without_corrections",FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F29", "эксперт",                                            "expert_name",							FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F30", "id лечения",                                         "treatcode",							FieldType.Long),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F31", "тип договора",                                       "contract_type",						FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F32", "тип программы",                                      "program_type",							FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F33", "id услуги (schid)",                                  "schid",								FieldType.Long),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F34", "id врача (dcode)",                                   "dcode",								FieldType.Long),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F35", "id должность по справочнику (profid)",				"profid",								FieldType.Long),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F36", "должность по справочнику",							"doctor_position_catalog",				FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F37", "должность",											"doctor_position",						FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F38", "направление: id врача (dcode)",						"referral_dcode",						FieldType.Long),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F39", "направление: id должность по справочнику (profid)",	"referral_profid",						FieldType.Long),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F40", "направление: должность по справочнику",				"referral_doctor_position_catalog",		FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F41", "направление: должность",								"referral_doctor_position",				FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F42", "№ гарантийного письма",								"garant_letter",						FieldType.String),	new List<object>() },
			{ new Tuple<string, string, string, FieldType>("F43", "категория",											"category",								FieldType.String),	new List<object>() },
		};

		static void Main(string[] args) {
			Logging.ToLog("====== Запуск");
			string pathToImport = Properties.Settings.Default.PathToImport;

			if (!Directory.Exists(pathToImport)) {
				string message = "Не удается получить доступ (не существует) к папке: " + pathToImport;
				SystemMail.SendMail(subjectError, message, receiverError);
				Logging.ToLog(message);
				return;
			}

			//file, result
			Dictionary<string, string> filesToImport = new Dictionary<string, string>();
			try {
				Logging.ToLog("Считывание содержимого папки для импорта: " + pathToImport);
				foreach (string file in Directory.GetFiles(pathToImport, "*.xlsx"))
					filesToImport.Add(file, string.Empty);
			} catch (Exception e) {
				string message = "Не удалось получить содержимое папки: " + e.Message + Environment.NewLine + e.StackTrace;
				SystemMail.SendMail(subjectError, message, receiverError);
				Logging.ToLog(message);
				return;
			}

			if (filesToImport.Count == 0) {
				Logging.ToLog("В папке для импорта нет файлов *.xlsx, завершение");
				return;
			}

			string sheetName = Properties.Settings.Default.SheetName;
			Logging.ToLog("Считывание содержимого файлов");
			string[] files = filesToImport.Keys.ToArray();
			foreach (string file in files) {
				Logging.ToLog("Очистка данных в хэше");
				Tuple<string, string, string, FieldType>[] headers = headersAndData.Keys.ToArray();
				foreach (Tuple<string, string, string, FieldType> header in headers)
					headersAndData[header].Clear();

				Logging.ToLog("Файл: " + file + ", лист: " + sheetName);
				try {
					DataTable dataTable = ExcelReader.ReadExcelFile(file, sheetName);

					if (dataTable.Columns.Count < 43) {
						string message = "Кол-во столбцов в файле меньше 43, пропуск обработки";
						filesToImport[file] = message;
						Logging.ToLog(file);
						continue;
					}

					int firstDataRow = GetFirstDataRow(dataTable, out string columnError);
					if (firstDataRow == -1) {
						string message = "Формат файла не соответствует заданному шаблону / не удалось найти строку с заголовками";
						if (!string.IsNullOrEmpty(columnError))
							message += "; Не удалось найти столбец: '" + columnError + "'";

						filesToImport[file] = message;
						Logging.ToLog(message);
						continue;
					}

					string fileInfo = string.Empty;
					Logging.ToLog("Считывание свойств файла");
					try {
						ShellPropertyCollection properties = new ShellPropertyCollection(file);
						string lastAuthor = properties["System.Document.LastAuthor"].ValueAsObject.ToString();
						string dateSaved = properties["System.Document.DateSaved"].ValueAsObject.ToString();
						string computerName = properties["System.ComputerName"].ValueAsObject.ToString();
						fileInfo = Path.GetFileName(file) + "@" + lastAuthor + "@" + dateSaved + "@" + computerName;
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}

					Logging.ToLog("Строка, содержащая заголовки: " + (firstDataRow + 1));
					Logging.ToLog("Считывание строк");
					int errorCounter = 0;
					for (int i = firstDataRow + 1; i < dataTable.Rows.Count; i++) {
						try {
							DataRow dataRow = dataTable.Rows[i];
							bool isOk = ParseRow(dataRow);
							if (!isOk) {
								Logging.ToLog("Ошибки обработки в строке: " + (i + 1));
								errorCounter++;
							}
						} catch (Exception) {
							errorCounter++;
						}
					}

					string fileResult = "Считано строк: " + headersAndData.First().Value.Count + ", ошибок: " + errorCounter;
					Logging.ToLog(fileResult);

					Logging.ToLog("Запись в БД");
					VerticaClient verticaClient = new VerticaClient(
						VerticaSettings.host,
						VerticaSettings.database,
						VerticaSettings.user,
						VerticaSettings.password);

					bool isUpdatedCorrected = verticaClient.ExecuteUpdateQuery(VerticaSettings.sqlInsert, headersAndData, fileInfo);
					fileResult += "; Загрузка в БД: " + (isUpdatedCorrected ? "Успешно" : "С ошибками");
					filesToImport[file] = fileResult;
				} catch (Exception e) {
					string message = e.Message + Environment.NewLine + e.StackTrace;
					filesToImport[file] = message;
					Logging.ToLog(message);
				}
			}

			string pathArchive = Path.Combine(pathToImport, "Archive");
			if (!Directory.Exists(pathArchive)) {
				try {
					Directory.CreateDirectory(pathArchive);
				} catch (Exception e) {
					string message = "Не удалось создать папку для архива: " + pathArchive + Environment.NewLine +
						e.Message + Environment.NewLine + e.StackTrace;
					Logging.ToLog(message);
					SystemMail.SendMail(subjectError, message, receiverError);
					return;
				}
			}

			Logging.ToLog("Перемещение файлов в архив: ");
			string now = DateTime.Now.ToString("yyyyMMddHHmmss");
			string resultMessage = "<table border=\"1\"><tr><th>FileName</th><th>Result</th></tr>";
			foreach (KeyValuePair<string, string> fileInfo in filesToImport) {
				string fileName = Path.GetFileName(fileInfo.Key);
				string fileResult = fileInfo.Value;

				try {
					File.Move(fileInfo.Key, Path.Combine(pathArchive, fileName));
					File.WriteAllText(Path.Combine(pathArchive, fileName + "_" + now + ".log"), fileResult);
					fileResult += "; Перемещено в архив";
				} catch (Exception e) {
					string message = "Не удалось переместить файл в архив: " + fileName + Environment.NewLine +
						e.Message + Environment.NewLine + e.StackTrace;
					Logging.ToLog(message);
					SystemMail.SendMail(subjectError, message, receiverError);
					fileResult += "; Не удалось переместить в архив";
				}

				string td = "<td>";
				if (!fileResult.Contains("ошибок: 0"))
					td = "<td bgcolor=\"#FFF100\">";
				else if (fileResult.Contains("С ошибками"))
					td = "<td bgcolor=\"#FFA500\">";
				else if (fileResult.Contains("Не удалось"))
					td = "<td bgcolor=\"#00B4FF\">";

				resultMessage += "<tr><td>" + fileName + "</td>" + td + fileResult + "</td></tr>";
			}

			resultMessage += "</table>";
			SystemMail.SendMail("Результат выполнения CleanedTreatmentsDetailsImport", resultMessage, Properties.Settings.Default.MailTo);

			Logging.ToLog("------ Завершение");
		}

		private static int GetFirstDataRow(DataTable dataTable, out string columnError) {
			columnError = string.Empty;

			for (int i = 0; i < dataTable.Rows.Count; i++) {
				DataRow dataRow = dataTable.Rows[i];
				string firstColumn = dataRow[0].ToString().ToLower();
				if (!firstColumn.Equals(headersAndData.Keys.First().Item2))
					continue;

				foreach (KeyValuePair<Tuple<string, string, string, FieldType>, List<object>> header in headersAndData) {
					string columnValue = dataRow[header.Key.Item1].ToString().ToLower();
					if (!columnValue.Equals(header.Key.Item2)) {
						columnError = header.Key.Item2;
						return -1;
					}
				}

				return i;
			}

			return -1;
		}

		private static bool ParseRow(DataRow dataRow) {
			Tuple<string, string, string, FieldType>[] headers = headersAndData.Keys.ToArray();

			foreach (Tuple<string, string, string, FieldType> header in headers) {
				string column = header.Item1;
				string value = dataRow[column].ToString();

				if (column.Equals("F1") && (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))) 
					return false;

				object parsed = null;
				FieldType fieldType = header.Item4;
				switch (fieldType) {
					case FieldType.String:
						if (!string.IsNullOrEmpty(value))
							parsed = value;
						break;
					case FieldType.Long:
						if (long.TryParse(value, out long resultLong))
							parsed = resultLong;						
						break;
					case FieldType.Float:
						if (float.TryParse(value, out float resultFloat))
							parsed = resultFloat;
						break;
					case FieldType.Date:
						if (DateTime.TryParse(value, out DateTime resultDateTime))
							parsed = resultDateTime;
						break;
					default:
						break;
				}

				headersAndData[header].Add(parsed);
			}

			return true;
		}
	}
}
