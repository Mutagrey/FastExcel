using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Text.RegularExpressions;

namespace FastExcelDNA
{
    public static class ExcelRefConverter
    {
        private const string ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        // Преобразует индекс столбца в букву: 1 -> A, 2 -> B, 27 -> AA...
        private static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        // Преобразует буквы столбца индекс: XFD -> 16384
        private static int GetExcelColumnIndex(string columnLetter)
        {
            columnLetter = columnLetter.ToUpper();
            int sum = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                sum *= 26;
                sum += (columnLetter[i] - 'A' + 1);
            }
            return sum;
        }

        [ExcelFunction(Category = "Reference Converter", Description = "Преобразует в адрес типа A1:B2")]
        public static string ToExcelCoordinates(int RowFirst, int ColumnFirst, int RowLast, int ColumnLast)
        {
            try
            {
                if (RowFirst <= 0)
                    RowFirst = 1;
                if (RowLast <= 0)
                    RowLast = RowFirst;
                if (ColumnFirst <= 0)
                    ColumnFirst = 1;
                if (ColumnLast <= 0)
                    ColumnLast = ColumnFirst;
                if (RowLast < RowFirst)
                {
                    int temp = RowFirst;
                    RowFirst = RowLast;
                    RowLast = temp;
                }
                if (ColumnLast < ColumnFirst)
                {
                    int temp = ColumnFirst;
                    ColumnFirst = ColumnLast;
                    ColumnLast = temp;
                }
                // Преобразовать индекс в букву
                var ColNameFirst = GetExcelColumnName(ColumnFirst);
                var ColNameLast = GetExcelColumnName(ColumnLast);
                // Итоговый адрес А1:А1
                return ColNameFirst + RowFirst + ":" + ColNameLast + RowLast;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        [ExcelFunction(Category = "Reference Converter", Description = "Преобразует из адреса в номера строк/столбцов [RowFirst, ColumnFirst, RowLast, ColumnLast]")]
        public static object ToNumericCoordinates(string textRef)
        {
            try
            {
                object[] result = new object[4];
                // A1:B2
                string[] refArr = textRef.Split(':');

                string curCell = string.Empty;
                for (int index = 0; index < refArr.Length; index++)
                {
                    string ColumnStr = string.Empty;
                    string RowStr = string.Empty;
                    if (index == 0)
                        curCell = refArr[0];
                    else
                        curCell = refArr[1];

                    bool isRefR1C1 = (bool)ExcelRefConverter.isR1C1ReferenceType(curCell);
                    CharEnumerator ce;
                    if (isRefR1C1)
                    {
                        char[] charsRow = { 'r', 'R' }; 
                        char[] charsCol = { 'c', 'C', 'с', 'С' };
                        RowStr  = curCell.Substring(1, curCell.IndexOfAny(charsCol) - 1);
                        ColumnStr = curCell.Substring(curCell.IndexOfAny(charsCol)+1);
                        result[index * 2] = int.Parse(RowStr);
                        result[index * 2 + 1] = int.Parse(ColumnStr);
                    }
                    else
                    {
                        ce = curCell.GetEnumerator();
                        while (ce.MoveNext())
                            if (char.IsLetter(ce.Current))
                                ColumnStr += ce.Current;
                            else
                                if (char.IsNumber(ce.Current))
                                    RowStr += ce.Current;
                        int i = 0;
                        ce = ColumnStr.GetEnumerator();
                        while (ce.MoveNext())
                            i = (26 * i) + ALPHABET.IndexOf(ce.Current) + 1;

                        result[index * 2] = int.Parse(RowStr);
                        result[index * 2 + 1] = i;
                    }

                    if (refArr.Length < 2)
                    {
                        result[2] = result[0];
                        result[3] = result[1];
                    }

                }

                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            
        }

        // Проверяет соответствует ли тип ссылки формату R1C1 или A1
        private static object isR1C1ReferenceType(string textRef)
        {
            try
            {
                //string r1c1Str = @"(R(\[-?\d+\])C(\[-?\d+\])" +
                //                @"|R(\[-?\d+\])C" +
                //                @"|RC(\[-?\d+\])" +
                //                @"|R\d+C\d+" +
                //                @"|R\d+C" +
                //                @"|RC\d+" +
                //                @"|R{1,1}C{1,1})";
                string r1c1Str = @"R\d+C\d+";
                r1c1Str = r1c1Str + "(:" + r1c1Str + ")?";
                //string A1Str = @"|(\$?[a-zA-Z]{1,3}\$?\d{1,7}(:\$?[a-zA-Z]{1,3}\$?\d{1,7})?)";
                string pattern = r1c1Str;// +A1Str;

                string res = string.Empty;
                if (Regex.IsMatch(textRef, pattern, RegexOptions.IgnoreCase))
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        // Преобразует адрес диапазона для ADO в нужный формат A1:A1 использует ExcelRefConverter
        public static string RefToADO(string textRef, int OffsetRW = 0, int OffsetCL = 0, int ResizeRW = 0, int ResizeCL = 0)
        {
            try
            {
                // Ссылка на диапазон
                object[] refObject = (object[])ExcelRefConverter.ToNumericCoordinates(textRef);

                // Количество строк и столбцов с учетом изменения размера диапазона Range
                int Rows;
                int Columns;
                if (ResizeRW <= 0)
                    Rows = (int)refObject[2] - (int)refObject[0] + 1;
                else
                    Rows = ResizeRW;
                if (ResizeCL <= 0)
                    Columns = (int)refObject[3] - (int)refObject[1] + 1;
                else
                    Columns = ResizeCL;

                int RowFirst = (int)refObject[0] + OffsetRW;
                int RowLast = Rows + RowFirst - 1;
                int ColumnFirst = (int)refObject[1] + OffsetCL;
                int ColumnLast = Columns + ColumnFirst - 1;
                // Новая ссылка с учетом смещения и изменения размера диапазона
                string newAddress = ExcelRefConverter.ToExcelCoordinates(RowFirst, ColumnFirst, RowLast, ColumnLast);

                return newAddress;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        // RegExp - Получает все адреса из формулы
        [ExcelFunction(IsHidden = true, Category = "Reference Converter", Description = "Преобразует из адреса в номера строк/столбцов [RowFirst, ColumnFirst, RowLast, ColumnLast]")]
        public static string GetCellRefFromFormula(string sFormula, string separator = ";")
        {
            try
            {
                if (separator.Length == 0)
                    separator = ";";

                string r1c1Str = @"(R(\[-?\d+\])C(\[-?\d+\])" +
                                @"|R(\[-?\d+\])C" +
                                @"|RC(\[-?\d+\])" +
                                @"|R\d+C\d+" +
                                @"|R\d+C" +
                                @"|RC\d+)";
                r1c1Str = r1c1Str + "(:" + r1c1Str + ")?";
                string A1Str = @"|(\$?[a-zA-Z]{1,3}\$?\d{1,7}(:\$?[a-zA-Z]{1,3}\$?\d{1,7})?)";
                string pattern = r1c1Str + A1Str;

                string res = "";
                foreach (Match match in Regex.Matches(sFormula, pattern, RegexOptions.IgnoreCase))
                    res = res + match.Value + separator;
                return res.Remove(res.Length - 1);
            }
            catch
            {
                return sFormula;
            }
        }
    }
}
