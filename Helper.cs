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
        public enum Strefa
        {
            undefined = 1,
            Czas_Pracy_Podstawowy = 2,
            Czas_Przestoju_Płatny_60 = 3,
            Czas_Przerwy = 4,
            Czas_Pracy_W_Akordzie = 5,
            Czas_Przestoju_Płatny_100 = 6,
            Czas_Pracy_Wykonywanej_Zdalnie = 7,
            Czas_Przestoju_Płatny_50 = 8,
            Czas_Pracy_Wykonywanej_Zdalnie_Okazjonalnie = 9,
            Czas_Pracy_W_Delegacji = 10,
            Czas_Pracy_Obsługi_Relacji = 11,
            Czas_Pracy_Poza_Relacją = 12,
            Czas_Odpoczynku_Nie_Wliczany_Do_CP = 13,
            Odpoczynek_Czas_Odpoczynku_Wliczany_Do_CP = 14,
            Czas_Wyjścia_Prywatnego = 15
        }
        public enum Odb_Nadg
        {
            DEFAULT = 1,
            O_BM = 2,
            O_NM = 3,
            W_PŁ = 4,
            W_NP = 5
        }
        public enum Typ_Zakladki
        {
            Nierozopznana = -1,
            Tabela_Stawek = 0,
            Karta_Ewidencji_Konduktora = 1,
            Karta_Ewidencji_Pracownika = 2,
            Grafik_Pracy_Pracownika = 3,
            Harmonogram_Pracy_Konduktora = 4
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
                        positions.Add(new Current_Position(cell.Address.ColumnNumber, cell.Address.RowNumber));
                    }
                }
                else
                {
                    if (formattedValue == keyWord)
                    {
                        positions.Add(new Current_Position(cell.Address.ColumnNumber, cell.Address.RowNumber));
                    }
                }
                
            });

            return positions.Take(limit).ToList();
        }
        public static string Truncate(string? value, int maxLength) =>
            string.IsNullOrEmpty(value) ? string.Empty : value.Length > maxLength ? value[..maxLength] : value;
    }
}