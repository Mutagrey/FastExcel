using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using Timer = System.Threading.Timer;
using System.Data.OleDb;
using FastExcelDNA.ADOUtil;


namespace FastExcelDNA.ExcelDNA.RTDServer
{
    [ProgId(ServerName), ComVisible(true)]
    // [ComVisible(true)] should be set for this class if your project is set as [ComVisible(false)] - the default under Visual Studio.
    public class ADORtdServer : ExcelRtdServer
    {
        public const string ServerName = "ADORtdServer0";

        private readonly object sync = new object();
        private readonly ThreadPoolQueue queue = new ThreadPoolQueue();
        private Timer timer;
        private Dictionary<Topic, IList<string>> topicToInfo0 = new Dictionary<Topic, IList<string>>();

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            Add(topic, topicInfo);
            newValues = true;
            return "Loading...";
        }

        //[ExcelCommand(MenuName = "Remote", MenuText = "Update New")]
        //public static void UpdateNew()
        //{
        //    foreach (var kv in activeServers)
        //    {
        //        kv.Key.Refresh(false);
        //    }
        //}

        //[ExcelCommand(MenuName = "Remote", MenuText = "Update All")]
        //public static void UpdateAll()
        //{
        //    foreach (var kv in activeServers)
        //    {
        //        kv.Key.Refresh(true);
        //    }
        //}

        //private static KeyValuePair<TKey, TValue> GetPair<TKey, TValue>(TKey key, TValue value)
        //{
        //    return new KeyValuePair<TKey, TValue>(key, value);
        //}

        //private void Refresh(bool wantAll)
        //{
        //    queue.Enqueue(() =>
        //    {
        //        List<KeyValuePair<Topic, IList<string>>> topics;
        //        lock (sync)
        //        {
        //            if (wantAll)
        //            {
        //                topics = topicToInfo0.ToList();
        //            }
        //            else
        //            {
        //                topics = newTopics.Keys.Select(k => GetPair(k, topicToInfo0[k])).ToList();
        //            }
        //        }

        //        // TODO: Probably no need to make a separate dictionary
        //        // Can zip over the topics and the results and assign each result directly
        //        var topicToData = GetData(topics);
        //        foreach (var kv in topicToData)
        //        {
        //            kv.Key.UpdateValue(kv.Value);
        //        }

        //        lock (sync)
        //        {
        //            foreach (var kv in topicToData)
        //            {
        //                newTopics.Remove(kv.Key);
        //            }
        //        }
        //    });
        //}

        private void Add(Topic topic, IList<string> topicInfo)
        {
            var delay = TimeSpan.FromSeconds(1);
            lock (sync)
            {
                topicToInfo0.Add(topic, topicInfo);
                if (timer == null)
                {
                    timer = new Timer(_ => OnTimer(), null, (int)delay.TotalMilliseconds, Timeout.Infinite);
                }
            }
        }

        private void OnTimer()
        {
            lock (sync)
            {
                if (timer != null)
                {
                    timer.Dispose();
                    timer = null;
                }
            }
            queue.Enqueue(SendRequest);
        }

        private int batchCounter = 0;
        private int topicCounter = 0;

        // I run on the queue
        private void SendRequest()
        {
            var newTopicToInfo = new Dictionary<Topic, IList<string>>();
            Dictionary<Topic, IList<string>> temp;
            lock (sync)
            {
                temp = topicToInfo0;
                topicToInfo0 = newTopicToInfo;
            }
            if (temp.Count > 0)
            {
                DoRequest(temp);
            }
        }

        private async void DoRequest(Dictionary<Topic, IList<string>> temp)
        {
            var batch = ++batchCounter;
            if (temp.Count > 0)
            {
                // send request... if async, make sure to process response on queue

                // Список файлов для загрузки
                Dictionary<string, IList<string>> filesToLoad = new Dictionary<string, IList<string>>();
                foreach (var p in temp.Values)
                {
                    var file = p[0];
                    if (!filesToLoad.ContainsKey(file))
                        filesToLoad.Add(file, new List<string>());
                    filesToLoad[file].Add(string.Join(",", p.ToArray()));
                }

                // Открываем подключения асинхронно
                Dictionary<string, Task<OleDbConnection>> openConTasks = new Dictionary<string, Task<OleDbConnection>>();
                foreach (var file in filesToLoad.Keys)
                {
                    openConTasks.Add(file, Task.Run(() => ADOManager.OpenConnectionAsync(file)));
                }

                // Ждем когда любое подключения откроется
                while (openConTasks.Count > 0)
                {
                    var finishedTask = await Task.WhenAny(openConTasks.Values.ToArray());
                    var curFileTask = openConTasks.First(v => v.Value.Equals(finishedTask));
                    var curTopic = temp.First(t => t.Value[0].Equals(curFileTask.Key));
                    try
                    {
                        // Загружаем все данные из текущего подключения
                        using (OleDbConnection opennedConnection = finishedTask.Result)
                        {
                            // Проходим по всем буферным топикам (где имя файла совпадает с текущим) и пытаемся загрузить данные
                            foreach (var kv in temp.Where(t => t.Value[0].Equals(curFileTask.Key)))
                            {
                                var topicCount = ++topicCounter;

                                var filePath = kv.Value[0];
                                var sheetName = kv.Value[1];
                                var AdoCellRef = kv.Value[2];
                                var rowID = Convert.ToInt32(kv.Value[3]);
                                var colID = Convert.ToInt32(kv.Value[4]);
                                var isNumericOnly = false;// Convert.ToBoolean(kv.Value[3]);

                                object result = ADOManager.ADO_ReadData(opennedConnection, filePath, sheetName, AdoCellRef, isNumericOnly);

                                if (result.GetType().IsArray)
                                {
                                    object[,] newRes = (object[,])result;
                                    kv.Key.UpdateValue(newRes[rowID, colID]);
                                }
                                else
                                {
                                    kv.Key.UpdateValue(result);
                                }
                            }
                        }
                        
                        openConTasks.Remove(curFileTask.Key);
                    }
                    catch (AggregateException ae)
                    {
                        curTopic.Key.UpdateValue(ae.Message);
                    }

                }
            }
        }

        protected override void DisconnectData(Topic topic) { }
    }
}
