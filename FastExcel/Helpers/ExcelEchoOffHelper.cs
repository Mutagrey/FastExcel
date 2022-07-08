using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastExcelDNA.ExcelDNA.Helpers
{
    // RIIA-style helpers to deal with Excel selections    
    // Don't use if you agree with Eric Lippert here: http://stackoverflow.com/a/1757344/44264
    public class ExcelEchoOffHelper : XlCall, IDisposable
    {
        object oldEcho;

        public ExcelEchoOffHelper()
        {
            oldEcho = Excel(xlfGetWorkspace, 40);
            Excel(xlcEcho, false);
        }

        public void Dispose()
        {
            Excel(xlcEcho, oldEcho);
        }
    }
}
