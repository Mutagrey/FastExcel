using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Emit;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using Microsoft.Office.Interop.Excel;
using ExcelDna.IntelliSense;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;

namespace FastExcelDNA
{
    [ComVisible(false)]
    public class AddInManager : IExcelAddIn
    {
        public static Application xlApp = (Application)ExcelDnaUtil.Application;
        //public static dynamic xlApp = ExcelDnaUtil.Application;

        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
            IntelliSenseServer.Install();
            ExcelIntegration.RegisterUnhandledExceptionHandler(delegate(object ex) { return "!!! CRIT ERROR: " + ex.ToString(); });
            ExcelAsyncUtil.CalculationCanceled += CalculationCanceled;
            ExcelAsyncUtil.CalculationEnded += CalculationEnded;
            xlApp.AfterCalculate += XlApp_AfterCalculate;
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
            IntelliSenseServer.Uninstall();
            ExcelAsyncUtil.CalculationCanceled -= CalculationCanceled;
            ExcelAsyncUtil.CalculationEnded -= CalculationEnded;
        }

        #region Cancellation support
        // We keep a CancellationTokenSource around, and set to a new one whenever a calculation has finished.
        public static CancellationTokenSource Cancellation = new CancellationTokenSource();
        private void XlApp_AfterCalculate()
        {
            if (Cancellation.IsCancellationRequested)
                Cancellation = new CancellationTokenSource();
        }
        public static void CalculationCanceled()
        {
            Cancellation.Cancel();
        }

        public static void CalculationEnded()
        {
            // Maybe we only need to set a new one when it was actually used...(when IsCanceled is true)?
            //Cancellation = new CancellationTokenSource();
        }
        #endregion

        // Делает расчет формулы через COM - Microsoft.Office.Interop.Excel. Предварительно сохраняем COM - Application, чтобы не нагружать процесс и делать расчеты асинхронно!!!
        public static object DoEvaluate(string formula)
        {
            try
            {
                return xlApp.Evaluate(formula);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
