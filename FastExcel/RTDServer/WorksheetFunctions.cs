using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace FastExcelDNA.ExcelDNA.RTDServer
{
    public static class WorksheetFunctions
    {
        [ExcelFunction(Category = "ADO READ DATA", Description = "Получает данные через подключение из другой книги, асинхронно - многопоточно", Name = "RTD.ADO")]
        public static object Func1(string fileName, string sheetName, string sRange, string rowID, string colID)
        {
            return "ПОКА НЕ РАБОТАЕТ";// XlCall.RTD(ADORtdServer.ServerName, null, fileName, sheetName, sRange, rowID, colID);
        }

    }
}
