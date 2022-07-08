using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;
using System.Data;
using System.Data.OleDb;
using ExcelDna.Integration;
using FastExcelDNA.ADOUtil;
using FastExcelDNA.ExcelDNA.Helpers;


namespace FastExcelDNA.ExcelDNA.ADOUtil
{
    public static class ADOWorksheetFunctions
    {
        private const string categoryFunction = "ADO READ DATA";

        [ExcelFunction(Category = categoryFunction, Description = "Получает данные через подключение из другой книги, асинхронно - многопоточно", Name = "ADO.ReadDataAsyncAwait")]
        public static object ADO_ReadDataAsyncAwait(
            [ExcelArgument(Name = "filePath", Description = "путь к фалу excel")] string filePath,
            [ExcelArgument(Name = "sheetName", Description = "имя листа книги")] string sheetName,
            [ExcelArgument(Name = "sRngFormula", Description = "формула-адрес (A1:B20)")] string sRngFormula,
            [ExcelArgument(Name = "OffsetRW", Description = "смещение по строкам")] int OffsetRW,
            [ExcelArgument(Name = "OffsetCL", Description = "смещение по столбцам")] int OffsetCL,
            [ExcelArgument(Name = "ResizeRW", Description = "размер диапазона ячеек по строкам")] int ResizeRW,
            [ExcelArgument(Name = "ResizeCL", Description = "размер диапазона ячеек по столбцам")] int ResizeCL,
            [ExcelArgument(Name = "isNumericOnly", Description = "загружаем только числа (текстовые значения заменяются на 0)")] bool isNumericOnly)
        {
            object result = ExcelTaskUtil.Run("ADO_ReadDataAsyncAwait", new object[] { filePath, sheetName, sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL, isNumericOnly },
                async token =>
                {
                    var con = await ADOManager.OpenConnectionAsync(filePath, AddInManager.Cancellation.Token);
                    return await ADOManager.ADO_ReadDataFormulaAsync(con, filePath, sheetName, sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL, isNumericOnly, AddInManager.Cancellation.Token);
                });
            // Check the asyncResult to see if we're still busy
            if (result.Equals(ExcelError.ExcelErrorNA))
                return "Loading...";
            return result;
        }

        [ExcelFunction(Category = categoryFunction, Description = "Получает данные через подключение из другой книги, асинхронно - многопоточно", Name = "ADO.ReadDataAsync")]
        public static object ADO_ReadDataAsync(
            [ExcelArgument(Name = "filePath", Description = "путь к фалу excel")] string filePath,
            [ExcelArgument(Name = "sheetName", Description = "имя листа книги")] string sheetName,
            [ExcelArgument(Name = "sRngFormula", Description = "формула-адрес (A1:B20)")] string sRngFormula,
            [ExcelArgument(Name = "OffsetRW", Description = "смещение по строкам")] int OffsetRW,
            [ExcelArgument(Name = "OffsetCL", Description = "смещение по столбцам")] int OffsetCL,
            [ExcelArgument(Name = "ResizeRW", Description = "размер диапазона ячеек по строкам")] int ResizeRW,
            [ExcelArgument(Name = "ResizeCL", Description = "размер диапазона ячеек по столбцам")] int ResizeCL,
            [ExcelArgument(Name = "isNumericOnly", Description = "загружаем только числа (текстовые значения заменяются на 0)")] bool isNumericOnly)
        {
            object result = ExcelAsyncUtil.Run("ADO_ReadDataAsync", new object[] { filePath, sheetName, sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL, isNumericOnly },
                delegate
                {
                    var conTask = ADOManager.OpenConnectionAsync(filePath, AddInManager.Cancellation.Token);
                    var con = conTask.Result;
                    var resTask = ADOManager.ADO_ReadDataFormulaAsync(con, filePath, sheetName, sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL, isNumericOnly, AddInManager.Cancellation.Token);
                    return resTask.Result;
                });
            // Check the asyncResult to see if we're still busy
            if (result.Equals(ExcelError.ExcelErrorNA))
                return "Loading...";
            return result;
        }
        [ExcelFunction(Category = categoryFunction, Description = "Получает данные через подключение из другой книги, асинхронно - многопоточно", Name = "ADO.ReadDataSync")]
        public static object ADO_ReadDataSync(
            [ExcelArgument(Name = "filePath", Description = "путь к фалу excel")] string filePath,
            [ExcelArgument(Name = "sheetName", Description = "имя листа книги")] string sheetName,
            [ExcelArgument(Name = "sRngFormula", Description = "формула-адрес (A1:B20)")] string sRngFormula,
            [ExcelArgument(Name = "OffsetRW", Description = "смещение по строкам")] int OffsetRW,
            [ExcelArgument(Name = "OffsetCL", Description = "смещение по столбцам")] int OffsetCL,
            [ExcelArgument(Name = "ResizeRW", Description = "размер диапазона ячеек по строкам")] int ResizeRW,
            [ExcelArgument(Name = "ResizeCL", Description = "размер диапазона ячеек по столбцам")] int ResizeCL,
            [ExcelArgument(Name = "isNumericOnly", Description = "загружаем только числа (текстовые значения заменяются на 0)")] bool isNumericOnly)
        {
            var con = ADOManager.OpenConnectionSync(filePath);
            return ADOManager.ADO_ReadDataFormula(con, filePath, sheetName, sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL, isNumericOnly);
        }

        [ExcelFunction(IsThreadSafe = true, Category = categoryFunction, Description = "Получает данные через подключение из другой книги, асинхронно - многопоточно", Name = "ADO.ReadDataAsyncThreadSafe")]
        public static async void ADO_ReadDataAsyncThreadSafe(
            [ExcelArgument(Name = "filePath", Description = "путь к фалу excel")] string filePath,
            [ExcelArgument(Name = "sheetName", Description = "имя листа книги")] string sheetName,
            [ExcelArgument(Name = "sRngFormula", Description = "формула-адрес (A1:B20)")] string sRngFormula,
            [ExcelArgument(Name = "OffsetRW", Description = "смещение по строкам")] int OffsetRW,
            [ExcelArgument(Name = "OffsetCL", Description = "смещение по столбцам")] int OffsetCL,
            [ExcelArgument(Name = "ResizeRW", Description = "размер диапазона ячеек по строкам")] int ResizeRW,
            [ExcelArgument(Name = "ResizeCL", Description = "размер диапазона ячеек по столбцам")] int ResizeCL,
            [ExcelArgument(Name = "isNumericOnly", Description = "загружаем только числа (текстовые значения заменяются на 0)")] bool isNumericOnly, ExcelAsyncHandle asyncHandle)
        {
            try
            {
                var con = await ADOManager.OpenConnectionAsync(filePath, AddInManager.Cancellation.Token);
                var result = await ADOManager.ADO_ReadDataFormulaAsync(con, filePath, sheetName, sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL, isNumericOnly, AddInManager.Cancellation.Token);
                asyncHandle.SetResult(result);
            }
            catch (Exception ex)
            {
                asyncHandle.SetException(ex);
            }
        }

        [ExcelFunction(Category = categoryFunction, Description = "Возвращает список листов книги - асинхронно")]
        public static object ADO_ReadSheetNamesAsync([ExcelArgument(Name = "filePath", Description = "путь к фалу excel")] string filePath)
        {
            object result = ExcelAsyncUtil.Run("ADO_ReadSheetNamesAsync", new object[] { filePath },
                delegate
                {
                    return ADOManager.ADO_ReadSheetNamesExcelDNA(filePath);
                });
            if (result.Equals(ExcelError.ExcelErrorNA))
                return "GettingSchema...";
            return result;
        }
    }
}
