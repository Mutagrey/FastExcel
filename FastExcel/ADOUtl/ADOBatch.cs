//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Data;
//using System.Data.OleDb;
//using System.Threading;
//using Microsoft.VisualBasic;
//using ExcelDna.Integration;
//using FastExcelDNA.ADOUtil;

//namespace FastExcelDNA.ExcelDNA.RTDServer
//{
//    // Модель ADO
//    public class ADOModel
//    {
//        public ExcelReference caller;
//        public string filePath;
//        public string sheetName;
//        public string sRange;
//        public int rowID;
//        public int colID;
//        public object result = null;

//        public ADOModel(ExcelReference curCaller, string curfilePath, string curSheetName, string curRange = "A1:A1", int curRowID = 0, int curColID = 0)
//        {
//            caller = curCaller;
//            filePath = curfilePath;
//            sheetName = curSheetName;
//            sRange = curRange;
//            rowID = curRowID;
//            colID = curRowID;
//        }
//    }

//    public class ADOBatch
//    {
//        // Пул подключений - все существующие подключения в текущей сессии
//        //private static Dictionary<string, OleDbConnection> connectionPool = new Dictionary<string, OleDbConnection>();
//        // Пул входных данных - в текущей сессии
//        private static Dictionary<string, IList<ADOModel>> adoModelPool = new Dictionary<string, IList<ADOModel>>();
//        private static Dictionary<ExcelReference, object> adoResults = new Dictionary<ExcelReference, object>();

//        private static Dictionary<ExcelReference, Task<object>> adoTaskResults = new Dictionary<ExcelReference, Task<object>>();

//        //private static readonly ThreadPoolQueue queue = new ThreadPoolQueue();
//        //private static Timer timer;

//        [ExcelFunction(Category = "ADO READ DATA", Description = "Получает данные через подключение из другой книги, асинхронно - многопоточно", Name = "ADO.Batch.ReadDataAsync", IsThreadSafe = true)]
//        public static void ADOBatch_ReadData(string fileName, string sheetName, string sRange, int rowID, int colID, ExcelAsyncHandle asyncHandle)
//        {
//            if (rowID == 0)
//                rowID = 0;
//            if (colID == 0)
//                colID = 0;
//            if (sRange.Length == 0)
//                sRange = "A1:A1";

//            ExcelReference caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

//            var adoModel = new ADOModel(caller, fileName, sheetName, sRange, rowID, colID);
//            AddModelAsync(adoModel);

//            //if (!adoModelPool.ContainsKey(adoModel.filePath))
//            //    adoModelPool.Add(adoModel.filePath, new List<ADOModel>());
//            //adoModelPool[adoModel.filePath].Add(adoModel);

//            //var delay = TimeSpan.FromSeconds(1);
//            //await Task.Delay((int)delay.TotalMilliseconds);


//            var temp = adoModelPool;
//            adoModelPool = new Dictionary<string, IList<ADOModel>>();
//            if (temp.Count > 0)
//            {
//                DoRequest(temp);
//            }

//            try
//            {

//                //Task<object> task;
//                //if (adoTaskResults.TryGetValue(caller, out task))
//                //{
//                //    var res = await task;
//                //    adoTaskResults.Remove(caller);
//                //    asyncHandle.SetResult(res);
//                //}

//                object res;
//                if (adoResults.TryGetValue(caller, out res))
//                {
//                    adoResults.Remove(caller);
//                    asyncHandle.SetResult(res);
//                }

//            }
//            catch (Exception ex)
//            {
//                asyncHandle.SetException(ex);
//            }


//        }

//        private static async void AddModelAsync(ADOModel adoModel)
//        {
//            var delay = TimeSpan.FromSeconds(1);

//            if (!adoModelPool.ContainsKey(adoModel.filePath))
//                adoModelPool.Add(adoModel.filePath, new List<ADOModel>());
//            adoModelPool[adoModel.filePath].Add(adoModel);

//            await Task.Delay((int)delay.TotalMilliseconds);
//        }

//        //private static void AddModel(ADOModel adoModel)
//        //{
//        //    var delay = TimeSpan.FromSeconds(1);

//        //    if (!adoModelPool.ContainsKey(adoModel.filePath))
//        //        adoModelPool.Add(adoModel.filePath, new List<ADOModel>());
//        //    adoModelPool[adoModel.filePath].Add(adoModel);

//        //    if (timer == null)
//        //    {
//        //        timer = new Timer(_ => OnTimer(), null, (int)delay.TotalMilliseconds, Timeout.Infinite);
//        //    }

//        //}

//        //private static void OnTimer()
//        //{
//        //    if (timer != null)
//        //    {
//        //        timer.Dispose();
//        //        timer = null;
//        //    }
//        //    queue.Enqueue(SendRequest);
//        //}

//        //// I run on the queue
//        //private static void SendRequest()
//        //{
//        //    var newDic = new Dictionary<string, IList<ADOModel>>();
//        //    Dictionary<string, IList<ADOModel>> temp;

//        //    temp = adoModelPool;
//        //    adoModelPool = newDic;
//        //    if (temp.Count > 0)
//        //    {
//        //        DoRequest(temp);
//        //    }
//        //}

//        private static async void DoRequest(Dictionary<string, IList<ADOModel>> temp)
//        {
//            foreach (var kv in temp)
//            {
//                var filePath = kv.Key;
//                // Открываем подключения асинхронно
//                OleDbConnection connection = await ADOManager.OpenConnectionAsync(filePath);
//                using (connection)
//                {
//                    foreach (var adoModel in kv.Value)
//                    {
//                        var sheetName = adoModel.sheetName;
//                        var AdoCellRef = adoModel.sRange;
//                        var rowID = adoModel.rowID;
//                        var colID = adoModel.colID;
//                        var isNumericOnly = false;// Convert.ToBoolean(kv.Value[3]);

//                        if (!adoResults.ContainsKey(adoModel.caller))
//                        {
//                            var res = ADOManager.ADO_ReadData(connection, filePath, sheetName, AdoCellRef, isNumericOnly);
//                            adoResults.Add(adoModel.caller, res);
//                        }

//                        //if (!adoTaskResults.ContainsKey(adoModel.caller))
//                        //{
//                        //    var task = Task.Run(() => ADOManager.ADO_ReadData(connection, filePath, sheetName, AdoCellRef, isNumericOnly));
//                        //    adoTaskResults.Add(adoModel.caller, task);
//                        //}
                        
//                    }
//                }
//            }

//        }
//    }
//}
