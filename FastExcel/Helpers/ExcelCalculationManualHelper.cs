using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastExcelDNA.ExcelDNA.Helpers
{
    public class ExcelCalculationManualHelper : XlCall, IDisposable
    {
        object oldCalculationMode;

        public ExcelCalculationManualHelper()
        {
            oldCalculationMode = Excel(xlfGetDocument, 14);
            Excel(xlcOptionsCalculation, 3);
        }

        public void Dispose()
        {
            Excel(xlcOptionsCalculation, oldCalculationMode);
        }
    }
}
