using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using System.Data;
using System.Data.OleDb;
using System.Threading;

namespace FastExcelDNA
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("ExcelDNA.ADOCOM")]
    public class ADOCOM
    {
        public string ComLibraryHello()
        {
            return "Hello from FastExcelDNA.ADOCOM";
        }

        public double Add(double x, double y)
        {
            return x + y;
        }


        public object CalcADOAsync(string[] files, string[] sheets, string[] ranges)
        {
            Dictionary<string,IList<string>> filesToLoad = new Dictionary<string,IList<string>>();

            for (int i = 0; i < files.Count(); i++ )
            {
                var file = files[i];
                var sheet = sheets[i];
                var range = ranges[i];

                if (!filesToLoad.ContainsKey(file))
                    filesToLoad.Add(file, new List<string>());
                filesToLoad[file].Add(i + "|" + sheet + "|" + range);
            }


            //Parallel.ForEach(filesToLoad, item =>
            //{

            //});
            foreach (var item in filesToLoad)
            {
                var file = item.Key;
                foreach (var curItem in item.Value)
                {
                    var split = curItem.Split('|');
                    var id = Convert.ToInt32(split[0]);
                    var sheet = split[1];
                    var range = split[2];
                }
            }

            return "NO DATA";

        }

        private async void CreateTasks(string file, IList<string> itemsToLoad)
        {
            OleDbConnection connection = await ADOUtil.ADOManager.OpenConnectionAsync(file);

            Dictionary<string, IList<string>> filesToLoad = new Dictionary<string, IList<string>>();

            using (connection)
            {
                foreach (var curItem in itemsToLoad)
                {
                    var split = curItem.Split('|');
                    var id = Convert.ToInt32(split[0]);
                    var sheet = split[1];
                    var range = split[2];

                    Task<object> task = Task.Run(() => ADOUtil.ADOManager.ADO_ReadData(connection, file, sheet, range, false));

                    var result = await task;

                }
            }

        }

    }
}
