using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace CleanedTreatmentsDetailsImport {
	public class Program {
		private static readonly string subjectError = "!!! Ошибка в работе CleanedTreatmentsDetailsImport";
		private static readonly string receiverError = Properties.Settings.Default.MailCopy;
		private static readonly string pathToImport = Properties.Settings.Default.PathToImport;

		//Tuple<column, header, field, type>, List<Value>
		public static readonly List<Header> headers = new List<Header> {
			new Header("F1",  "договор",                                              "contract",                                 typeof(string),     false),
			new Header("F2",  "программа",                                            "program",                                  typeof(string),     false),
			new Header("F3",  "полис",                                                "policy",                                   typeof(string),     false),
			new Header("F4",  "фио пациента",                                         "patient_full_name",                        typeof(string),     true), 
			new Header("F5",  "дата рождения",                                        "patient_birthday",                         typeof(DateTime),   false),
			new Header("F6",  "возраст",                                              "patient_age",                              typeof(long),       false),
			new Header("F7",  "история болезни",                                      "patient_histnum",                          typeof(long),       false),
			new Header("F8",  "диагноз по мкб",                                       "mkb_diagnoses",                            typeof(string),     false),
			new Header("F9",  "№ зуба",                                               "toothnum",                                 typeof(string),     false),
			new Header("F10", "дата приема",                                          "treat_date",                               typeof(DateTime),   true), 
			new Header("F11", "код услуги",                                           "kodoper",                                  typeof(string),     false),
			new Header("F12", "наименование услуги",                                  "service_name",                             typeof(string),     false),
			new Header("F13", "количество услуг",                                     "service_count",                            typeof(long),       true), 
			new Header("F14", "стоимость услуги",                                     "service_cost",                             typeof(float),      false),
			new Header("F15", "сумма всего",                                          "amount_total",                             typeof(float),      false),
			new Header("F16", "сумма всего с учетом скидки",                          "amount_total_with_discount",               typeof(float),      true), 
			new Header("F17", "филиал",                                               "filial",                                   typeof(string),     false),
			new Header("F18", "отделение",                                            "department",                               typeof(string),     false),
			new Header("F19", "фио врача",                                            "doctor_full_name",                         typeof(string),     false),
			new Header("F20", "направление: филиал",                                  "referral_filial",                          typeof(string),     false),
			new Header("F21", "направление: отделение",                               "referral_department",                      typeof(string),     false),
			new Header("F22", "направление: врач",                                    "referral_doctor_full_name",                typeof(string),     false),
			new Header("F23", "согласование",                                         "agreement",                                typeof(string),     false),
			new Header("F24", "ограничение по услуге",                                "service_restrict",                         typeof(string),     false),
			new Header("F25", "примечание эксперта",                                  "accepted_arragement",                      typeof(string),     false),
			new Header("F26", "принятые меры",                                        "expert_comment",                           typeof(string),     false),
			new Header("F27", "не входит с сч",                                       "average_check_exclude_rule",               typeof(string),     false),
			new Header("F28", "оставлено без исправлений (сумма)",                    "amount_total_left_without_corrections",    typeof(string),     false),
			new Header("F29", "эксперт",                                              "expert_name",                              typeof(string),     false),
			new Header("F30", "id лечения",                                           "treatcode",                                typeof(long),       true), 
			new Header("F31", "тип договора",                                         "contract_type",                            typeof(string),     false),
			new Header("F32", "тип программы",                                        "program_type",                             typeof(string),     false),
			new Header("F33", "id услуги (schid)",                                    "schid",                                    typeof(long),       true), 
			new Header("F34", "id врача (dcode)",                                     "dcode",                                    typeof(long),       false),
			new Header("F35", "id должность по справочнику (profid)",                 "profid",                                   typeof(long),       false),
			new Header("F36", "должность по справочнику",                             "doctor_position_catalog",                  typeof(string),     false),
			new Header("F37", "должность",                                            "doctor_position",                          typeof(string),     false),
			new Header("F38", "направление: id врача (dcode)",                        "referral_dcode",                           typeof(long),       false),
			new Header("F39", "направление: id должность по справочнику (profid)",    "referral_profid",                          typeof(long),       false),
			new Header("F40", "направление: должность по справочнику",                "referral_doctor_position_catalog",         typeof(string),     false),
			new Header("F41", "направление: должность",                               "referral_doctor_position",                 typeof(string),     false),
			new Header("F42", "№ гарантийного письма",                                "garant_letter",                            typeof(string),     false),
			new Header("F43", "категория",											  "category",                                 typeof(string),     false),
			new Header("F44", "идентификатор строки (ordtid)",						  "ordtid",                                   typeof(long),       false)
		};

		private static readonly Dictionary<int, string> months = new Dictionary<int, string> {
			{ 1, "01 Январь" },
			{ 2, "02 Февраль" },
			{ 3, "03 Март" },
			{ 4, "04 Апрель" },
			{ 5, "05 Май" },
			{ 6, "06 Июнь" },
			{ 7, "07 Июль" },
			{ 8, "08 Август" },
			{ 9, "09 Сентябрь" },
			{ 10, "10 Октябрь" },
			{ 11, "11 Ноябрь" },
			{ 12, "12 Декабрь" }
		};

		public static DataTable FileContentTreatmentsDetails = new DataTable();
		public static List<ItemProfitAndLoss> FileContentProfitAndLoss = new List<ItemProfitAndLoss>();

		public class Header {
			public string ColumnKey { get; }
			public string ColumnName { get; }
			public string DbField { get; }
			public Type FieldType { get; }
			public bool NotNull { get; }

			public Header (string columnKey, string columnName, string dbField, Type fieldType, bool notNull) {
				ColumnKey = columnKey;
				ColumnName = columnName;
				DbField = dbField;
				FieldType = fieldType;
				NotNull = notNull;
			}
		}

		private readonly static Dictionary<string, string> filesToImport = new Dictionary<string, string>();


		static void Main(string[] args) {
			Logging.ToLog("====== Запуск");

			if (!Directory.Exists(pathToImport)) {
				string message = "Не удается получить доступ (не существует) к папке: " + pathToImport;
				SystemMail.SendMail(subjectError, message, receiverError);
				Logging.ToLog(message);
				return;
			}
			
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
				try {
					Logging.ToLog("Файл: " + file + ", лист: " + sheetName);
					string fileResult = ReadTreatmentsDetailsFileContent(file, sheetName, out string fileInfo);
					if (!fileResult.Contains("ошибок: 0"))
						continue;

					Logging.ToLog("Запись в БД");
					VerticaClient verticaClient = new VerticaClient(
						VerticaSettings.host,
						VerticaSettings.database,
						VerticaSettings.user,
						VerticaSettings.password);

					bool isUpdatedCorrected = verticaClient.ExecuteUpdateQuery(VerticaSettings.sqlInsertTreatmentsDetails, true, fileInfo);
					fileResult += "; Загрузка в БД: " + (isUpdatedCorrected ? "Успешно" : "С ошибками");
					filesToImport[file] = fileResult;
				} catch (Exception e) {
					string message = e.Message + Environment.NewLine + e.StackTrace;
					filesToImport[file] = message;
					Logging.ToLog(message);
				}
			}

			string pathArchive = GetArchivePath(out string messageErrorArchive);
			if (!string.IsNullOrEmpty(pathArchive)) {
				SystemMail.SendMail(subjectError, messageErrorArchive, receiverError);
				return;
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

		public static string GetArchivePath(out string messageError) {
			messageError = string.Empty;
			string pathArchive = Path.Combine(pathToImport, "Archive");

			if (!Directory.Exists(pathArchive)) {
				try {
					Directory.CreateDirectory(pathArchive);
				} catch (Exception e) {
					messageError = "Не удалось создать папку для архива: " + pathArchive + Environment.NewLine +
						e.Message + Environment.NewLine + e.StackTrace;
					Logging.ToLog(messageError);
				}
			}

			return pathArchive;
		}

		public static string ReadProfitAndLossLFileContent(string file, string sheetName, out string fileInfo, BackgroundWorker bw = null) {
			fileInfo = "Unavailable";

			DataTable dataTable = ExcelReader.ReadExcelFile(file, sheetName);
			
			if (dataTable == null || dataTable.Rows.Count == 0) {
				string message = "Не удалось считать файл / нет данных";
				filesToImport[file] = message;
				Logging.ToLog(message);
				return message;
			}

			string msg = "Считывание свойств файла";
			Logging.ToLog(msg);
			if (bw != null)
				bw.ReportProgress(0, msg);

			try {
				ShellPropertyCollection properties = new ShellPropertyCollection(file);
				string lastAuthor = properties["System.Document.LastAuthor"].ValueAsObject.ToString();
				string dateSaved = properties["System.Document.DateSaved"].ValueAsObject.ToString();
				string computerName = properties["System.ComputerName"].ValueAsObject.ToString();
				fileInfo = Path.GetFileName(file) + "@" + lastAuthor + "@" + dateSaved + "@" + computerName;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			msg = "Разбор содержимого строк";
			Logging.ToLog(msg);
			if (bw != null)
				bw.ReportProgress(0, msg);

			int errorCounter = 0;

			int hasDataRow = 0;
			int quarterRow = 1;
			int objectNameRow = 2;
			int dateRow = 3;
			int firstDataRow = 4;

			int firstDataColumn = 3;

			if (dataTable.Rows.Count < firstDataRow + 1) {
				msg = "Формат таблицы не соответствует ожидаемому (кол-во строк меньше необходимого - " + (firstDataRow + 1) + ")";
				Logging.ToLog(msg);
				if (bw != null)
					bw.ReportProgress(0, msg);

				return msg;
			}

			msg = "Считывание строки, содержащей даты";
			Logging.ToLog(msg);
			if (bw != null)
				bw.ReportProgress(0, msg);

			int year = -1;
			DataRow dataRowDates = dataTable.Rows[dateRow];
			List<string> dates = new List<string>();
			foreach (object item in dataRowDates.ItemArray) {
				if (item == null || string.IsNullOrEmpty(item.ToString()) || string.IsNullOrWhiteSpace(item.ToString())) {
					dates.Add(string.Empty);
					continue;
				}

				if (double.TryParse(item.ToString(), out double result)) { 
					if (result > 2000 && result < 2030) {
						dates.Add(item.ToString());
						if (year == -1)
							year = (int)result;
					} else {
						try {
							DateTime dt = DateTime.FromOADate(result);
							dates.Add(months[dt.Month]);

							if (year == -1)
								year = dt.Year;
						} catch (Exception exc) {
							dates.Add(string.Empty);
							Logging.ToLog(exc.Message + Environment.NewLine + exc.StackTrace);
						}
					}
				} else {
					if (DateTime.TryParse(item.ToString(), out DateTime dt)) {
						dates.Add(months[dt.Month]);

						if (year == -1)
							year = dt.Year;
					} else {
						dates.Add(item.ToString());
					}
				}
			}

			if (year == -1) {
				msg = "Не удалось определить год в загруженном файле, поиск по строке: " + (dateRow + 1);
				Logging.ToLog(msg);
				if (bw != null)
					bw.ReportProgress(0, msg);

				return msg;
			}

			string groupNameLevel1 = null;
			string groupNameLevel2 = null;

			for (int i = firstDataRow; i < dataTable.Rows.Count; i++) {
				msg = "Строка: " + (i + 1) + " / " + dataTable.Rows.Count;
				Logging.ToLog(msg);
				if (bw != null)
					bw.ReportProgress(0, msg);

				DataRow dataRow = dataTable.Rows[i];

				try {
					string rowGroupNameLevel1 = dataRow[0].ToString();
					string rowGroupNameLevel2 = dataRow[1].ToString();
					string rowGroupNameLevel3 = dataRow[2].ToString();

					string rowGroupNameTotal = rowGroupNameLevel1 + rowGroupNameLevel2 + rowGroupNameLevel3;
					if (string.IsNullOrEmpty(rowGroupNameTotal) ||
						string.IsNullOrWhiteSpace(rowGroupNameTotal)) {
						msg = "Отсуютствую заголовки групп, пропуск";
						Logging.ToLog(msg);
						if (bw != null)
							bw.ReportProgress(0, msg);

						groupNameLevel1 = null;
						groupNameLevel2 = null;

						continue;
					}

					if (string.IsNullOrEmpty(rowGroupNameLevel3) ||
						string.IsNullOrWhiteSpace(rowGroupNameLevel3))
						rowGroupNameLevel3 = null;

					if (!string.IsNullOrEmpty(rowGroupNameLevel1) &&
						!string.IsNullOrWhiteSpace(rowGroupNameLevel1)) {
						groupNameLevel1 = rowGroupNameLevel1;
						groupNameLevel2 = null;
					}

					if (!string.IsNullOrEmpty(rowGroupNameLevel2) ||
						!string.IsNullOrWhiteSpace(rowGroupNameLevel2))
						groupNameLevel2 = rowGroupNameLevel2;

					for (int x = firstDataColumn; x < dataRow.ItemArray.Length; x++) {
						if (dataRow[x] == null || string.IsNullOrEmpty(dataRow[x].ToString()) || string.IsNullOrWhiteSpace(dataRow[x].ToString()))
							continue;

						string objectName = dataTable.Rows[objectNameRow][x].ToString();
						if (string.IsNullOrEmpty(objectName) || string.IsNullOrWhiteSpace(objectName)) {
							if (string.IsNullOrEmpty(dates[x]))
								continue;
							else
								objectName = "Не указано";
						}

						if (double.TryParse(dataRow[x].ToString(), out double result)) {
							ItemProfitAndLoss item = new ItemProfitAndLoss {
								ObjectName = dataTable.Rows[objectNameRow][x].ToString(),
								PeriodYear = year,
								PeriodType = dates[x],
								GroupNameLevel1 = groupNameLevel1,
								GroupNameLevel2 = groupNameLevel2,
								GroupNameLevel3 = rowGroupNameLevel3,
								Value = result,
								GroupSortingOrder = i + 1,
								ObjectSrotingOrder = x + 1
							};

							string quarter = dataTable.Rows[quarterRow][x].ToString();
							if (int.TryParse(quarter, out int resultQuarter))
								if (resultQuarter != 0)
									item.Quarter = resultQuarter;

							string hasData = dataTable.Rows[hasDataRow][x].ToString();
							if (int.TryParse(hasData, out int resultHasData))
								if (!string.IsNullOrEmpty(hasData) && !string.IsNullOrWhiteSpace(hasData))
									item.HasData = resultHasData == 4;

							FileContentProfitAndLoss.Add(item);
						} else {
							msg = "Не удалось преобразовать в число значение в столбце " + (x + 1) + " - '" + dataRow[x] + "'";
							Logging.ToLog(msg);
							if (bw != null)
								bw.ReportProgress(0, msg);

							errorCounter++;
							continue;
						}
					}
				} catch (Exception e) {
					if (bw != null)
						bw.ReportProgress(0, "Строка: " + (i + 1) + ", " + e.Message);

					errorCounter++;
				}
			}

			string fileResult = "Считано объектов: " + FileContentProfitAndLoss.Count + ", ошибок: " + errorCounter;
			Logging.ToLog(fileResult);
			return fileResult;
		}

		public static string ReadTreatmentsDetailsFileContent(string file, string sheetName, out string fileInfo, BackgroundWorker bw = null) {
			fileInfo = "Unavailable";

			Logging.ToLog("Подготовка таблицы для записи считанных данных");
			FileContentTreatmentsDetails = new DataTable();
			foreach (Header header in headers)
				FileContentTreatmentsDetails.Columns.Add(header.DbField, header.FieldType);

			DataTable dataTable = ExcelReader.ReadExcelFile(file, sheetName);

			int minColumnQuantity = headers.Count;
			if (dataTable.Columns.Count < minColumnQuantity) {
				string message = "Кол-во столбцов в файле меньше " + minColumnQuantity + ", пропуск обработки";
				filesToImport[file] = message;
				Logging.ToLog(file);
				return message;
			}

			int firstDataRow = GetFirstDataRowTreatmentsDetails(dataTable, out string columnError);
			if (firstDataRow == -1) {
				string message = "Формат файла не соответствует заданному шаблону / не удалось найти строку с заголовками";
				if (!string.IsNullOrEmpty(columnError))
					message += "; Не удалось найти столбец: '" + columnError + "'";

				filesToImport[file] = message;
				Logging.ToLog(message);
				return message;
			}

			string msg = "Считывание свойств файла";
			Logging.ToLog(msg);
			if (bw != null)
				bw.ReportProgress(0, msg);

			try {
				ShellPropertyCollection properties = new ShellPropertyCollection(file);
				string lastAuthor = properties["System.Document.LastAuthor"].ValueAsObject.ToString();
				string dateSaved = properties["System.Document.DateSaved"].ValueAsObject.ToString();
				string computerName = properties["System.ComputerName"].ValueAsObject.ToString();
				fileInfo = Path.GetFileName(file) + "@" + lastAuthor + "@" + dateSaved + "@" + computerName;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			msg = "Строка, содержащая заголовки: " + (firstDataRow + 1);
			Logging.ToLog(msg);
			if (bw != null)
				bw.ReportProgress(0, msg);

			msg = "Разбор содержимого строк";
			Logging.ToLog(msg);
			if (bw != null)
				bw.ReportProgress(0, msg);

			int errorCounter = 0;
			for (int i = firstDataRow + 1; i < dataTable.Rows.Count; i++) {
				try {
					ParseRow(dataTable.Rows[i]);
				} catch (Exception e) {
					if (bw != null)
						bw.ReportProgress(0, "Строка: " + (i + 1) + ", " + e.Message);

					errorCounter++;
				}
			}

			string fileResult = "Считано строк: " + FileContentTreatmentsDetails.Rows.Count + ", ошибок: " + errorCounter;
			Logging.ToLog(fileResult);
			return fileResult;
		}

		private static int GetFirstDataRowTreatmentsDetails(DataTable dataTable, out string columnError) {
			columnError = string.Empty;

			for (int i = 0; i < dataTable.Rows.Count; i++) {
				DataRow dataRow = dataTable.Rows[i];
				string firstColumn = dataRow[0].ToString().ToLower();
				if (!firstColumn.Equals(headers[0].ColumnName))
					continue;

				foreach (Header header in headers) {
					string columnValue = null;

					try {
						if (dataTable.Columns.Contains(header.ColumnKey)) {
							columnValue = dataRow[header.ColumnKey].ToString().ToLower();
							if (!columnValue.Equals(header.ColumnName)) {
								columnError = header.ColumnName;
								return -1;
							}
						}
					} catch (Exception) {}
				}

				return i;
			}

			return -1;
		}

		private static void ParseRow(DataRow dataRow) {
			DataRow dataRowParsed = FileContentTreatmentsDetails.NewRow();

			foreach (Header header in headers) {
				string column = header.ColumnKey;
				string value = null;
				try {
					value = dataRow[column].ToString();
				} catch (Exception) { }

				if (column.Equals("F1") && (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))) 
					return;

				object parsed = null;
				if (value != null) {
					if (header.FieldType == typeof(string)) { 
						if (!string.IsNullOrEmpty(value))
							parsed = value;
					} else if (header.FieldType == typeof(long)) { 
						if (long.TryParse(value, out long resultLong))
							parsed = resultLong;
					} else if (header.FieldType == typeof(float)) { 
						if (float.TryParse(value, out float resultFloat))
							parsed = resultFloat;
					} else if (header.FieldType == typeof(DateTime)) {
						if (DateTime.TryParse(value, out DateTime resultDateTime))
							parsed = resultDateTime;
						else if (double.TryParse(value, out double resultDouble))
							parsed = DateTime.FromOADate(resultDouble);
					} else { 
						Logging.ToLog("Неизвестный тип данных: " + header.FieldType);
					}
				}

				if (header.NotNull && parsed == null)
					throw new Exception("Обязательный столбец " + header.ColumnName + " имеет пустое / ошибочное значение");

				if (parsed != null)
					dataRowParsed[header.DbField] = parsed;
			}

			FileContentTreatmentsDetails.Rows.Add(dataRowParsed);
		}

		public static bool IsCompareReadedDataToDbOk(DataTable dataTableDb, 
			out int rowsToDelete, out int rowsLoadedBefore, out int rowsWithNoOrdtid,
			BackgroundWorker bw = null) {
			rowsToDelete = 0;
			rowsLoadedBefore = 0;
			rowsWithNoOrdtid = 0;

			if (dataTableDb == null) {
				if (bw != null)
					bw.ReportProgress(0, "Таблица с данными из БД пустая, невозможно выполнить проверку");

				return false;
			}

			List<string> requiredColumns = new List<string>() {
				"ordtid",
				"treatcode",
				"treat_nctrdate",
				"schid",
				"tooth_name",
				"client",
				"etl_insert_time",
				"file_info",
				"loadingUserName",
				"schcount",
				"schamount"
			};

			foreach (string column in requiredColumns) {
				if (dataTableDb.Columns.Contains(column))
					continue;

				if (bw != null)
					bw.ReportProgress(0, "В таблице с данными из БД отсутсвует необходимое поле: " + column);

				return false;
			}

			bool isCompareOk = true;

			Dictionary<string, List<int>> contentIndexes = new Dictionary<string, List<int>>();
			Dictionary<long, List<int>> ordtidIndexes = new Dictionary<long, List<int>>();
			for(int i = 0; i < FileContentTreatmentsDetails.Rows.Count; i++) {
				string id = GetRowContentSearchBy(FileContentTreatmentsDetails.Rows[i]);
				if (!contentIndexes.ContainsKey(id))
					contentIndexes.Add(id, new List<int>());

				if (long.TryParse(FileContentTreatmentsDetails.Rows[i]["ordtid"].ToString(), out long ordtid)) {
					if (!ordtidIndexes.ContainsKey(ordtid))
						ordtidIndexes.Add(ordtid, new List<int>());

					ordtidIndexes[ordtid].Add(i);
				}

				contentIndexes[id].Add(i);
			}

			DataTable fileContentDeleted = FileContentTreatmentsDetails.Clone();

			foreach (DataRow dataRow in dataTableDb.Rows) {
				//Console.WriteLine("Compare: dataRow: " + dataTableDb.Rows.IndexOf(dataRow));
				long ordtid = (long)dataRow["ordtid"];
				string etlInsertTime = dataRow["etl_insert_time"].ToString();
				string fileInfo = dataRow["file_info"].ToString();
				string loadingUserName = dataRow["loadingUserName"].ToString();

				DataRow dataRowToLoad = null;
				List<int> dataRowsByOrdtid = new List<int>();
				if (ordtidIndexes.ContainsKey(ordtid))
					dataRowsByOrdtid = ordtidIndexes[ordtid];

				if (dataRowsByOrdtid.Count == 0) {
					string idDb = GetRowDbSearchBy(dataRow);
					List<int> dataRowsFullSearchRaw = new List<int>();
					if (contentIndexes.ContainsKey(idDb)) 
						dataRowsFullSearchRaw = contentIndexes[idDb];

					List<int> dataRowsFullSearch = new List<int>();
					foreach (int item in dataRowsFullSearchRaw) {
						DataRow rowContent = FileContentTreatmentsDetails.Rows[item];
						if (string.IsNullOrEmpty(rowContent["ordtid"].ToString()))
							dataRowsFullSearch.Add(item);
					}

					if (dataRowsFullSearch.Count > 0) {
						int rowIndex = dataRowsFullSearch[0];
						dataRowToLoad = FileContentTreatmentsDetails.Rows[rowIndex];
						FileContentTreatmentsDetails.Rows[rowIndex]["ordtid"] = ordtid;
						FileContentTreatmentsDetails.Rows[rowIndex]["amount_total"] = dataRow["schamount"];
					} else {
						DataRow rowToDelete = fileContentDeleted.NewRow();
						rowToDelete["ordtid"] = ordtid;
						rowToDelete["service_count"] = 0;
						rowToDelete["amount_total_with_discount"] = 0;
						rowToDelete["amount_total"] = dataRow["schamount"];
						fileContentDeleted.Rows.Add(rowToDelete);

						rowsToDelete++;

						if (rowsToDelete == 10) {
							string messageSkip = "Имеется более 10 подобных строк, пропуск дальнейшего отображения";
							Logging.ToLog(messageSkip);
							if (bw != null)
								bw.ReportProgress(0, messageSkip);

							continue;
						} else if (rowsToDelete > 10)
							continue;

						string message = "Услуга с ordtid " + ordtid + " не найдена в загруженном файле, отмечаем как удаленную";
						Logging.ToLog(message);
						if (bw != null)
							bw.ReportProgress(0, message);

						continue;
					}
				} else if (dataRowsByOrdtid.Count > 1) {
					string message = "Найдено больше одного соответствия для строки с идентификатором (ordtid): " + ordtid;
					Logging.ToLog(message);
					if (bw != null) {
						bw.ReportProgress(0, new string('-', 40));
						bw.ReportProgress(0, message);
					}

					isCompareOk = false;
					continue;
				} else
					dataRowToLoad = FileContentTreatmentsDetails.Rows[dataRowsByOrdtid[0]];

				if (dataRowToLoad != null && !string.IsNullOrEmpty(etlInsertTime)) {
					rowsLoadedBefore++;
					isCompareOk = false;
					Console.WriteLine("rowsLoadedBefore: " + dataRowToLoad.ItemArray[43].ToString());

					if (rowsLoadedBefore == 10) {
						string messageSkip = "Имеется более 10 подобных строк, пропуск дальнейшего отображения";
						Logging.ToLog(messageSkip);
						if (bw != null)
							bw.ReportProgress(0, messageSkip);

						continue;
					} else if (rowsLoadedBefore > 10)
						continue;

					string message = "Строка с индентификатором (ordtid) " + ordtid + " была загружена в БД ранее";
					Logging.ToLog(message);
					if (bw != null) {
						bw.ReportProgress(0, new string('-', 40));
						bw.ReportProgress(0, message);
					}

					WriteOutInfoAboutRowToLoad(dataRowToLoad, bw);

					string messageWhenLoaded =
						"Было загружено ранее пользователем: " + loadingUserName + Environment.NewLine +
						"        Время загрузки: " + etlInsertTime + Environment.NewLine +
						"        Загруженный файл: " + fileInfo;

					Logging.ToLog(messageWhenLoaded);
					if (bw != null)
						bw.ReportProgress(0, messageWhenLoaded);
				}
			}

			DataRow[] rowsToLoadWithoutOrdtid = FileContentTreatmentsDetails.Select("ordtid is null");
			if (rowsToLoadWithoutOrdtid.Length > 0) {
				string message = "Не удалось найти сопоставления загруженных строк с данными в БД";
				Logging.ToLog(message);
				if (bw != null) {
					bw.ReportProgress(0, new string('-', 40));
					bw.ReportProgress(0, message);
				}

				isCompareOk = false;

				foreach (DataRow row in rowsToLoadWithoutOrdtid) {
					rowsWithNoOrdtid++;

					if (rowsWithNoOrdtid == 10) {
						string messageSkip = "Имеется более 10 подобных строк, пропуск дальнейшего отображения";
						Logging.ToLog(messageSkip);
						if (bw != null)
							bw.ReportProgress(0, messageSkip);

						continue;
					} else if (rowsWithNoOrdtid > 10)
						continue;

					WriteOutInfoAboutRowToLoad(row, bw);
				}
			}

			FileContentTreatmentsDetails.Merge(fileContentDeleted);

			return isCompareOk;
		}

		public static void ApplyAverageDiscount(DataTable dataTableDb, BackgroundWorker bw = null) {
			if (dataTableDb == null) {
				if (bw != null)
					bw.ReportProgress(0, "Таблица с данными из БД не задана, пропуск");

				return;
			}

			if (!dataTableDb.Columns.Contains("schamount") ||
				!dataTableDb.Columns.Contains("ordtid")) {
				if (bw != null)
					bw.ReportProgress(0, "");

				return;
			}

			double amountTotalDB = 0;
            foreach (DataRow row in dataTableDb.Rows)
				amountTotalDB += double.Parse(row["schamount"].ToString());

			if (bw != null)
				bw.ReportProgress(0, "Стоимость услуг в БД: " + amountTotalDB.ToString("N", CultureInfo.CurrentCulture));

			double amountTotalWithDiscount = 0;
			foreach (DataRow row in FileContentTreatmentsDetails.Rows)
				amountTotalWithDiscount += (float)row["amount_total_with_discount"];

			if (bw != null)
				bw.ReportProgress(0, "Стоимость услуг со скидкой в загруженном файле: " + amountTotalWithDiscount.ToString("N", CultureInfo.CurrentCulture));

			double averageDiscount = (amountTotalDB - amountTotalWithDiscount) / amountTotalDB;
			if (bw != null)
				bw.ReportProgress(0, "Общая сумма снятий: " + (amountTotalDB - amountTotalWithDiscount).ToString("N", CultureInfo.CurrentCulture) + 
					", процент снятий: " + averageDiscount.ToString("P", CultureInfo.CurrentCulture));

			FileContentTreatmentsDetails.Columns.Add(new DataColumn("average_discount", typeof(float)));
			FileContentTreatmentsDetails.Columns.Add(new DataColumn("amount_total_with_average_discount", typeof(float)));

			double amountTotalWithAverageDiscount = 0;
            foreach (DataRow row in FileContentTreatmentsDetails.Rows) {
				row["average_discount"] = averageDiscount;
				row["amount_total_with_average_discount"] = (float)row["amount_total"] * (1.0d - averageDiscount);
				amountTotalWithAverageDiscount += (float)row["amount_total_with_average_discount"];
			}

			if (bw != null)
				bw.ReportProgress(0, "Сумма услуг в БД с распределенной скидкой: " + amountTotalWithAverageDiscount.ToString("N", CultureInfo.CurrentCulture));

			if (bw != null)
				bw.ReportProgress(0, "Расхождение распределенной средней скидки: " + 
					(amountTotalWithDiscount - amountTotalWithAverageDiscount).ToString("N", CultureInfo.CurrentCulture));
        }

		private static string GetRowContentSearchBy(DataRow row) {
			return row["treatcode"].ToString() +
				row["treat_date"].ToString() +
				//row["patient_full_name"].ToString() +
				row["toothnum"].ToString() +
				row["schid"].ToString();
		}

		private static string GetRowDbSearchBy(DataRow row) {
			return row["treatcode"].ToString() +
				row["treat_nctrdate"].ToString() +
				//row["client"].ToString() +
				row["tooth_name"].ToString() +
				row["schid"].ToString();
		}

		private static void WriteOutInfoAboutRowToLoad(DataRow dataRowToLoad, BackgroundWorker bw) {
			string patient = dataRowToLoad["patient_full_name"].ToString();
			string treatDate = dataRowToLoad["treat_date"].ToString();
			string treatcodeLoaded = dataRowToLoad["treatcode"].ToString();
			string serviceName = dataRowToLoad["service_name"].ToString();
			string serviceSchid = dataRowToLoad["schid"].ToString();
			string serviceCount = dataRowToLoad["service_count"].ToString();
			string toothnumLoaded = dataRowToLoad["toothnum"].ToString();
			string rowInfo = "Информация о строке в загруженном файле: " + Environment.NewLine +
				"        Пациент: " + patient + Environment.NewLine +
				"        Дата приема: " + treatDate + Environment.NewLine +
				"        Код лечения Treatcode: " + treatcodeLoaded + Environment.NewLine +
				"        Услуга: " + serviceName + Environment.NewLine +
				"        Код услуги Schid: " + serviceSchid + Environment.NewLine +
				"        Кол-во услуг: " + serviceCount + Environment.NewLine +
				"        № Зуба: '" + toothnumLoaded + "'";

			Logging.ToLog(rowInfo);
			if (bw != null)
				bw.ReportProgress(0, rowInfo);
		}
	}
}
