using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace Konduktor_Reader
{
    internal static class Helper
    {
        public class Current_Position
        {
            private int col = 1;
            private int row = 1;
            public int Col
            {
                get => col;
                set
                {
                    if (value < 1)
                    {
                        Program.error_logger.New_Custom_Error("Błąd w programie, próba czytania komórki w kolumnie mniejszej niż 1");
                        throw new ArgumentOutOfRangeException(nameof(Col), "Błąd w programie, próba czytania komórki w kolumnie mniejszej niż 1");
                    }
                    col = value;
                }
            }
            public int Row
            {
                get => row;
                set
                {
                    if (value < 1)
                    {
                        Program.error_logger.New_Custom_Error("Błąd w programie, próba czytania komórki w rzędzie mniejszym niż 1");
                        throw new ArgumentOutOfRangeException(nameof(Col), "Błąd w programie, próba czytania komórki w rzędzie mniejszym niż 1");
                    }
                    row = value;
                }
            }
        }

        public static bool Try_Get_Type_From_String<T>(string? value, ref T result)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }
            try
            {
                if (typeof(T) == typeof(TimeSpan) && TimeSpan.TryParse(value, out var Time_Span_Value))
                {
                    result = (T)(object)Time_Span_Value;
                    return true;
                }
                if (typeof(T) == typeof(DateTime) && DateTime.TryParse(value, out var Date_Time_Value))
                {
                    result = (T)(object)Date_Time_Value;
                    return true;
                }

                if (typeof(T) == typeof(List<TimeSpan>) && TimeSpan.TryParse(value, out var timeSpanValueNew))
                {
                    if (result is List<TimeSpan> list)
                    {
                        list.Add(timeSpanValueNew);
                        return true;
                    }
                }

                var Converted_Value = Convert.ChangeType(value, typeof(T));
                if (Converted_Value != null)
                {
                    result = (T)Converted_Value;
                    return true;
                }
                

            }
            catch
            {
            }
            return false;
        }
        public static bool Try_Get_Type_From_String<T>(string? value, Action<T> method)
        {
            T refvalue = default!;
            if (Try_Get_Type_From_String<T>(value, ref refvalue))
            {
                method(refvalue);
                return true;
            }
            return false;
        }
        public static string Read_Value_Diffrent_Possible_Cells_In_Row(int Rows, int StartRow, int StartCol, IXLWorksheet Zakladka)
        {
            for (int i = 0; i < Rows; i++)
            {
                string Result = Zakladka.Cell(StartRow, StartCol).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(Result)){
                    return Result;
                }
            }
            return string.Empty;
        }
        public static List<Current_Position> Find_Staring_Points_Tabele_Stawek(IXLWorksheet Zakladka, string Key_Word)
        {
            List<Current_Position> starty = [];
            int Limiter = 1000;
            int counter = 0;
            foreach (IXLCell? cell in Zakladka.CellsUsed())
            {
                try
                {
                    if (cell.HasFormula && !cell.Address.ToString()!.Equals(cell.FormulaA1))
                    {
                        counter++;
                        if (counter > Limiter)
                        {
                            break;
                        }
                        continue;
                    }
                    if (cell.Value.ToString().Contains(Key_Word))
                    {
                        starty.Add(new Current_Position()
                        {
                            Row = cell.Address.RowNumber,
                            Col = cell.Address.ColumnNumber
                        });
                    }
                }
                catch
                {
                    continue;
                }
            }
            return starty;
        }
        public static string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }
            return value.Length > maxLength ? value.Substring(0, maxLength) : value;
        }
        public static bool Valid_SQLConnection_String(string Connection_String)
        {
            try
            {
                using (var connection = new SqlConnection(Connection_String))
                {
                    connection.Open();
                    connection.Close();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
