using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Threading;
using Microsoft.VisualBasic;

namespace FastExcelDNA.ADOUtil
{
    public static class ADOManager
    {
        // Пул подключений - все существующие подключения в текущей сессии
        private static Dictionary<string, OleDbConnection> connectionPoolTasks = new Dictionary<string, OleDbConnection>();

        #region Создаем подключения и управляем ими
        // Получаем строку подключения для файла Excel
        private static string GetConnectionString(string filePath)
        {
            string conSTR = "";
            string provider = "Microsoft.ACE.OLEDB.12.0";
            filePath = filePath.ToLower();
            if (filePath.Contains("xlsx"))
            {
                conSTR = "Provider=" + provider + ";Data Source=" + filePath + "; " + "Extended Properties=\"Excel 12.0 Xml;HDR=No;IMEX=1\";";
            }
            else if (filePath.Contains("xlsm"))
            {
                conSTR = "Provider=" + provider + ";Data Source=" + filePath + "; " + "Extended Properties=\"Excel 12.0 Macro;HDR=No;IMEX=1\";";
            }
            else if (filePath.Contains("xlsb"))
            {
                conSTR = "Provider=" + provider + ";Data Source=" + filePath + "; " + "Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\";";
            }
            else if (filePath.Contains("xls"))
            {
                conSTR = "Provider=" + provider + ";Data Source=" + filePath + "; " + "Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\";";
            }
            return conSTR;
        }
        // Открываем подключение OleDbConnection асинхронно (async - >= .Net 4.5)
        public static OleDbConnection OpenConnectionSync(string filePath)
        {
            if (connectionPoolTasks.ContainsKey(filePath))
            {
                return connectionPoolTasks[filePath];
            }
            else
            {
                // Строка подключения
                string connectionString = GetConnectionString(filePath);
                // Подключение
                var con = new OleDbConnection(connectionString);
                con.Open();
                return con;
            }
        }
        // Открываем подключение OleDbConnection асинхронно (async - >= .Net 4.5)
        public static async Task<OleDbConnection> OpenConnectionAsync(string filePath, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (connectionPoolTasks.ContainsKey(filePath))
            {
                return connectionPoolTasks[filePath];
            }
            else
            {
                // Строка подключения
                string connectionString = GetConnectionString(filePath);
                // Подключение
                var con = new OleDbConnection(connectionString);
                await con.OpenAsync(cancellationToken).ConfigureAwait(false);
                return con;
            }
        }
        #endregion

        #region Считываем данные через ADO
        // Считываем данные Excel через подключение
        public static object ADO_ReadData(OleDbConnection connection, string filePath, string sheetName, string AdoCellRef, bool isNumericOnly = false)
        {
            try
            {
                // Проверяем, что подключение открыто
                if (connection.State != ConnectionState.Open)
                    return "Can't open, connection: " + connection.State.ToString();

                // Преобразовать адрес в нужный формат А1:А1
                var newAdoRef = ExcelRefConverter.RefToADO(AdoCellRef);
                // Запрос
                string queryString = "SELECT * FROM [" + sheetName + "$" + newAdoRef + "]";
                // Таблица в которой хранятся результаты загрузки
                System.Data.DataTable resTable = new System.Data.DataTable();
                // Создаем адаптер 
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                // Указываем запрос для адаптера 
                adapter.SelectCommand = new OleDbCommand(queryString, connection);
                // Получить данные
                adapter.Fill(resTable);
                // Преобразовать в массив
                var newRes = new object[resTable.Rows.Count, resTable.Columns.Count];
                for (int i = 0; i < resTable.Rows.Count; i++)
                {
                    for (int j = 0; j < resTable.Columns.Count; j++)
                    {
                        var curData = resTable.Rows[i].ItemArray[j];
                        if (!Convert.IsDBNull(curData))
                            if (Information.IsNumeric(curData) || !isNumericOnly)
                                newRes[i, j] = curData;
                    }
                }
                return newRes;

            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        // Считываем данные Excel через подключение с учетом единственного диапазона
        public static async Task<object> ADO_ReadDataFormulaAsync(OleDbConnection connection, string filePath, string sheetName, string sRngFormula, int OffsetRW, int OffsetCL, int ResizeRW, int ResizeCL, bool isNumericOnly = false, CancellationToken cancellationToken = default(CancellationToken))
        {
            try
            {
                // Распознаем в формуле все адреса типов XlA1, XlR1C1
                string[] formulas = await Task.Run(() => ExcelRefConverter.GetCellRefFromFormula(sRngFormula, ";").Split(';'), cancellationToken);

                string strToReplace = string.Empty;
                int index = 0;
                string resStr = sRngFormula;
                if (formulas.Length > 1)
                {
                    // ---- Формулы - считаем по одной ячейке ---------- ОЧЕНЬ ДОЛГО, НАДО УСКОРИТЬ!
                    object[,] resData = new object[ResizeRW, ResizeCL];
                    for (int i = 0; i < ResizeRW; i++)
                    {
                        for (int j = 0; j < ResizeCL; j++)
                        {
                            foreach (string curRef in formulas)
                            {
                                // Преобразует адрес диапазона для ADO в нужный формат A1:A1
                                string ADOAddres = await Task.Run(() => ExcelRefConverter.RefToADO(curRef, OffsetRW + i, OffsetCL + j, 1, 1), cancellationToken);

                                // READ DATA FROM ADO
                                object adoData = await Task.Run(() => ADO_ReadData(connection, filePath, sheetName, ADOAddres, isNumericOnly), cancellationToken);

                                //Заменяем значения текущей ячейки
                                strToReplace = adoData.ToString();
                                if (strToReplace.Length == 0)
                                    strToReplace = "0";
                                index = resStr.IndexOf(curRef);
                                if (index >= 0)
                                    resStr = resStr.Remove(index) + strToReplace + resStr.Substring(index + curRef.Length);
                                index = index + strToReplace.Length;
                            }

                            resStr = resStr.Replace(",", ".");
                            var eval = await Task.Run(() => AddInManager.DoEvaluate("=" + resStr), cancellationToken);
                            resData[i, j] = eval;
                        }
                    }
                    return resData;
                }
                else
                {
                    // ---- Один Дианазон - считаем сразу через один запрос -----
                    // Преобразует адрес диапазона для ADO в нужный формат A1:A1
                    string ADOAddres = await Task.Run(() => ExcelRefConverter.RefToADO(sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL), cancellationToken);
                    // READ DATA FROM ADO
                    var adoData = await Task.Run(() => ADO_ReadData(connection, filePath, sheetName, ADOAddres, isNumericOnly), cancellationToken);
                    return adoData;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        // Считываем данные Excel через подключение с учетом единственного диапазона
        public static object ADO_ReadDataFormula(OleDbConnection connection, string filePath, string sheetName, string sRngFormula, int OffsetRW, int OffsetCL, int ResizeRW, int ResizeCL, bool isNumericOnly = false)
        {
            try
            {
                // Распознаем в формуле все адреса типов XlA1, XlR1C1
                string[] formulas = ExcelRefConverter.GetCellRefFromFormula(sRngFormula, ";").Split(';');

                string strToReplace = string.Empty;
                int index = 0;
                string resStr = sRngFormula;
                if (formulas.Length > 1)
                {
                    // ---- Формулы - считаем по одной ячейке ---------- ОЧЕНЬ ДОЛГО, НАДО УСКОРИТЬ!
                    object[,] resData = new object[ResizeRW, ResizeCL];
                    for (int i = 0; i < ResizeRW; i++)
                    {
                        for (int j = 0; j < ResizeCL; j++)
                        {
                            foreach (string curRef in formulas)
                            {
                                // Преобразует адрес диапазона для ADO в нужный формат A1:A1
                                string ADOAddres = ExcelRefConverter.RefToADO(curRef, OffsetRW + i, OffsetCL + j, 1, 1);

                                // READ DATA FROM ADO
                                object adoData = ADO_ReadData(connection, filePath, sheetName, ADOAddres, isNumericOnly);

                                //Заменяем значения текущей ячейки
                                strToReplace = adoData.ToString();
                                if (strToReplace.Length == 0)
                                    strToReplace = "0";
                                index = resStr.IndexOf(curRef);
                                if (index >= 0)
                                    resStr = resStr.Remove(index) + strToReplace + resStr.Substring(index + curRef.Length);
                                index = index + strToReplace.Length;
                            }

                            resStr = resStr.Replace(",", ".");
                            var eval = AddInManager.DoEvaluate("=" + resStr);
                            resData[i, j] = eval;
                        }
                    }
                    return resData;
                }
                else
                {
                    // ---- Один Дианазон - считаем сразу через один запрос -----
                    // Преобразует адрес диапазона для ADO в нужный формат A1:A1
                    string ADOAddres = ExcelRefConverter.RefToADO(sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL);
                    // READ DATA FROM ADO
                    var adoData = ADO_ReadData(connection, filePath, sheetName, ADOAddres, isNumericOnly);
                    return adoData;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        #endregion

        #region Считываем схему книги. Получаем список листов.
        // Считываем список листов - схему Excel через подключение
        public static object ADO_ReadSheetNamesExcelDNA(string filePath)
        {
            try
            {
                // Создаем и открываем подключение асинхронно и записываем его а КЭШ, чтобы не открывать повторно
                Task<OleDbConnection> openConTask = OpenConnectionAsync(filePath);
                
                //AddConnectionTask(filePath);

                if (!openConTask.IsCompleted)
                    return "Error connection: " + openConTask.Status;
                OleDbConnection connection = openConTask.Result;
                // Проверяем, что подключение открыто
                if (connection.State != ConnectionState.Open)
                {
                    return "Can't open, connection: " + connection.State.ToString();
                }
                // Таблица со схемой книги - листы и прочие таблицы
                //System.Data.DataTable infoTable = (System.Data.DataTable)connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                System.Data.DataTable infoTable = (System.Data.DataTable)connection.GetSchema("Tables");

                // Преобразовать в массив
                var res = new object[infoTable.Rows.Count];
                for (int i = 0; i < infoTable.Rows.Count; i++)
                {
                    res[i] = ((DataRow)infoTable.Rows[i]).ItemArray[2];
                }
                // Вернуть в качестве строки
                return String.Join(";", res.Where(c => c.ToString().Contains("$")));
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }
#endregion
    }
}
