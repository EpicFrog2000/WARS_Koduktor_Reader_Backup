using System.Collections.Concurrent;
using ClosedXML.Excel;

namespace Excel_Data_Importer_WARS
{
    internal static class Helper
    {
        public class Current_Position
        {
            private int col = 1;
            private int row = 1;

            public Current_Position()
            {
            }
            public Current_Position(int new_col, int new_row)
            {
                Col = new_col;
                Row = new_row;
            }

            public int Col
            {
                get => col;
                set
                {
                    if (value < 1)
                    {
                        //Program.error_logger.New_Custom_Error("Błąd w programie, próba czytania komórki w kolumnie mniejszej niż 1");
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
                        ArgumentOutOfRangeException argumentOutOfRangeException = new(nameof(Row), $"Błąd w programie, próba czytania komórki w rzędzie mniejszym niż 1. Kolumna: {Col}");
                        throw argumentOutOfRangeException;
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
                if (!string.IsNullOrEmpty(Result))
                {
                    return Result;
                }
            }
            return string.Empty;
        }

        public static List<Current_Position> Find_Starting_Points(IXLWorksheet worksheet, string keyWord, bool CompareMode=true)
        {
            const int limit = 1000;
            ConcurrentBag<Current_Position> positions = [];
            IXLCell[] cells = [.. worksheet.CellsUsed()];

            Parallel.ForEach(cells, (cell, state) =>
            {
                if (positions.Count >= limit)
                {
                    state.Stop();
                    return;
                }

                if (cell.HasFormula && !cell.FormulaA1.Equals(cell.Address.ToString()))
                    return;

                string formattedValue = cell.GetFormattedString();
                if (CompareMode)
                {
                    if (formattedValue.Contains(keyWord, StringComparison.OrdinalIgnoreCase))
                    {
                        positions.Add(new Current_Position
                        {
                            Row = cell.Address.RowNumber,
                            Col = cell.Address.ColumnNumber
                        });
                    }
                }
                else
                {
                    if (formattedValue == keyWord)
                    {
                        positions.Add(new Current_Position
                        {
                            Row = cell.Address.RowNumber,
                            Col = cell.Address.ColumnNumber
                        });
                    }
                }
                
            });

            return positions.Take(limit).ToList();
        }

        public static string Truncate(string? value, int maxLength) =>
            string.IsNullOrEmpty(value) ? string.Empty : value.Length > maxLength ? value[..maxLength] : value;
    }
}