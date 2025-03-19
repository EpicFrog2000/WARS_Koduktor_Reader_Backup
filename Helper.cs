using System.Collections.Concurrent;
using System.Diagnostics;
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
        public enum Typ_Insert_Obecnosc
        {
            Zerowka = 1, // To znaczy że nie ma godzin pracy czyli insert godziny od 00:00 do 00:00
            Normalna = 2,
            Nadgodziny = 3,
            Nieinsertuj = 4
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
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            const int limit = 1000;
            ConcurrentBag<Current_Position> positions = [];
            IXLCell[] cells = [.. worksheet.CellsUsed()];

            Parallel.ForEach(cells, (cell, state) =>
            {
                if (positions.Count >= limit)
                {
                    state.Stop();
                    Helper.Pomiar.Avg_Find_Starting_Points = PomiaryStopWatch.Elapsed;
                    return;
                }

                if (cell.HasFormula && !cell.FormulaA1.Equals(cell.Address.ToString()))
                {
                    Helper.Pomiar.Avg_Find_Starting_Points = PomiaryStopWatch.Elapsed;
                    return;
                }
                    

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
            Helper.Pomiar.Avg_Find_Starting_Points = PomiaryStopWatch.Elapsed;
            return positions.Take(limit).ToList();
        }
        public static string Truncate(string? value, int maxLength) =>
            string.IsNullOrEmpty(value) ? string.Empty : value.Length > maxLength ? value[..maxLength] : value;
        public static Typ_Insert_Obecnosc Get_Typ_Insert_Obecnosc(int Rok, int Miesiac, int Dzien, List<TimeSpan> Godziny_Pracy_Od,
            List<TimeSpan> Godziny_Pracy_Do,
            decimal Liczba_Godzin_Nadliczbowych_50 = 0,
            decimal Liczba_Godzin_Nadliczbowych_100 = 0,
            decimal Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 = 0,
            decimal Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100 = 0)
        {
            DateTime startDate = new(Rok, Miesiac, 1);
            DateTime endDate = new(Rok, Miesiac, DateTime.DaysInMonth(Rok, Miesiac));

            bool found = false;
            for (DateTime dzien = startDate; dzien <= endDate; dzien = dzien.AddDays(1))
            {
                if(dzien.Day == Dzien)
                {
                    found = true;
                }
            }
            if (!found)
            {
                return Typ_Insert_Obecnosc.Zerowka;
            }

            if (Godziny_Pracy_Od.Count >= 1)
            {
                return Typ_Insert_Obecnosc.Normalna;
            }
            else
            {
                if (Liczba_Godzin_Nadliczbowych_50 > 0 || Liczba_Godzin_Nadliczbowych_100 > 0 || Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 > 0 || Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100 > 0)
                {
                    return Typ_Insert_Obecnosc.Nadgodziny;
                }
                else
                {
                    return Typ_Insert_Obecnosc.Zerowka;
                }
            }
        }

        public static Typ_Insert_Obecnosc Get_Typ_Insert_Plan(int Rok, int Miesiac, int Dzien, TimeSpan Godzina_Pracy_Od, TimeSpan Godzina_Pracy_Do)
        {
            DateTime startDate = new(Rok, Miesiac, 1);
            DateTime endDate = new(Rok, Miesiac, DateTime.DaysInMonth(Rok, Miesiac));

            bool found = false;
            for (DateTime dzien = startDate; dzien <= endDate; dzien = dzien.AddDays(1))
            {
                if (dzien.Day == Dzien)
                {
                    found = true;
                }
            }
            if (!found)
            {
                return Typ_Insert_Obecnosc.Zerowka;
            }

            if(Godzina_Pracy_Od == Godzina_Pracy_Do)
            {
                return Typ_Insert_Obecnosc.Nieinsertuj;
            }
            return Typ_Insert_Obecnosc.Normalna;

        }

        public static class Pomiar
        {
            private static TimeSpan avg_Get_Metadane_Pliku = TimeSpan.Zero;
            private static TimeSpan avg_Process_Files = TimeSpan.Zero;
            private static TimeSpan avg_Process_1_Zakladka = TimeSpan.Zero;
            private static TimeSpan avg_MoveFile = TimeSpan.Zero;
            private static TimeSpan avg_Copy_Bad_Sheet_To_Files_Folder = TimeSpan.Zero;
            private static TimeSpan avg_Open_Workbook = TimeSpan.Zero;
            private static TimeSpan avg_Get_Typ_Zakladki = TimeSpan.Zero;
            private static TimeSpan avg_Usun_Ukryte_Karty = TimeSpan.Zero;
            private static TimeSpan avg_Find_Starting_Points = TimeSpan.Zero;
            private static TimeSpan avg_Insert_Obecnosc_Command = TimeSpan.Zero;
            private static TimeSpan avg_Get_Dane_Z_Pliku = TimeSpan.Zero;
            private static TimeSpan avg_Insert_Atrybuty_Do_Optimy = TimeSpan.Zero;
            private static TimeSpan avg_Create_Transaction = TimeSpan.Zero;
            private static TimeSpan avg_Dodawanie_Do_Bazy = TimeSpan.Zero;

            public static TimeSpan Avg_Get_Metadane_Pliku
            {
                get => avg_Get_Metadane_Pliku;
                set
                {
                    if (avg_Get_Metadane_Pliku == TimeSpan.Zero)
                    {
                        avg_Get_Metadane_Pliku = value;
                        return;
                    }
                    avg_Get_Metadane_Pliku = (value + avg_Get_Metadane_Pliku) / 2;
                }
            }
            public static TimeSpan Avg_Process_Files
            {
                get => avg_Process_Files;
                set
                {
                    if (avg_Process_Files == TimeSpan.Zero)
                    {
                        avg_Process_Files = value;
                        return;
                    }
                    avg_Process_Files = (value + avg_Process_Files) / 2;
                }
            }
            public static TimeSpan Avg_MoveFile
            {
                get => avg_MoveFile;
                set
                {
                    if (avg_MoveFile == TimeSpan.Zero)
                    {
                        avg_MoveFile = value;
                        return;
                    }
                    avg_MoveFile = (value + avg_MoveFile) / 2;
                }
            }
            public static TimeSpan Avg_Copy_Bad_Sheet_To_Files_Folder
            {
                get => avg_Copy_Bad_Sheet_To_Files_Folder;
                set
                {
                    if (avg_Copy_Bad_Sheet_To_Files_Folder == TimeSpan.Zero)
                    {
                        avg_Copy_Bad_Sheet_To_Files_Folder = value;
                        return;
                    }
                    avg_Copy_Bad_Sheet_To_Files_Folder = (value + avg_Copy_Bad_Sheet_To_Files_Folder) / 2;
                }
            }
            public static TimeSpan Avg_Open_Workbook
            {
                get => avg_Open_Workbook;
                set
                {
                    if (avg_Open_Workbook == TimeSpan.Zero)
                    {
                        avg_Open_Workbook = value;
                        return;
                    }
                    avg_Open_Workbook = (value + avg_Open_Workbook) / 2;
                }
            }
            public static TimeSpan Avg_Get_Typ_Zakladki
            {
                get => avg_Get_Typ_Zakladki;
                set
                {
                    if (avg_Get_Typ_Zakladki == TimeSpan.Zero)
                    {
                        avg_Get_Typ_Zakladki = value;
                        return;
                    }
                    avg_Get_Typ_Zakladki = (value + avg_Get_Typ_Zakladki) / 2;
                }
            }
            public static TimeSpan Avg_Usun_Ukryte_Karty
            {
                get => avg_Usun_Ukryte_Karty;
                set
                {
                    if (avg_Usun_Ukryte_Karty == TimeSpan.Zero)
                    {
                        avg_Usun_Ukryte_Karty = value;
                        return;
                    }
                    avg_Usun_Ukryte_Karty = (value + avg_Usun_Ukryte_Karty) / 2;
                }
            }
            public static TimeSpan Avg_Find_Starting_Points
            {
                get => avg_Find_Starting_Points;
                set
                {
                    if (avg_Find_Starting_Points == TimeSpan.Zero)
                    {
                        avg_Find_Starting_Points = value;
                        return;
                    }
                    avg_Find_Starting_Points = (value + avg_Find_Starting_Points) / 2;
                }
            }
            public static TimeSpan Avg_Insert_Obecnosc_Command
            {
                get => avg_Insert_Obecnosc_Command;
                set
                {
                    if (avg_Insert_Obecnosc_Command == TimeSpan.Zero)
                    {
                        avg_Insert_Obecnosc_Command = value;
                        return;
                    }
                    avg_Insert_Obecnosc_Command = (value + avg_Insert_Obecnosc_Command) / 2;
                }
            }
            public static TimeSpan Avg_Get_Dane_Z_Pliku
            {
                get => avg_Get_Dane_Z_Pliku;
                set
                {
                    if (avg_Get_Dane_Z_Pliku == TimeSpan.Zero)
                    {
                        avg_Get_Dane_Z_Pliku = value;
                        return;
                    }
                    avg_Get_Dane_Z_Pliku = (value + avg_Get_Dane_Z_Pliku) / 2;
                }
            }
            public static TimeSpan Avg_Insert_Atrybuty_Do_Optimy
            {
                get => avg_Insert_Atrybuty_Do_Optimy;
                set
                {
                    if (avg_Insert_Atrybuty_Do_Optimy == TimeSpan.Zero)
                    {
                        avg_Insert_Atrybuty_Do_Optimy = value;
                        return;
                    }
                    avg_Insert_Atrybuty_Do_Optimy = (value + avg_Insert_Atrybuty_Do_Optimy) / 2;
                }
            }
            public static TimeSpan Avg_Create_Transaction
            {
                get => avg_Create_Transaction;
                set
                {
                    if (avg_Create_Transaction == TimeSpan.Zero)
                    {
                        avg_Create_Transaction = value;
                        return;
                    }
                    avg_Create_Transaction = (value + avg_Create_Transaction) / 2;
                }
            }
            public static TimeSpan Avg_Dodawanie_Do_Bazy
            {
                get => avg_Dodawanie_Do_Bazy;
                set
                {
                    if (avg_Dodawanie_Do_Bazy == TimeSpan.Zero)
                    {
                        avg_Dodawanie_Do_Bazy = value;
                        return;
                    }
                    avg_Dodawanie_Do_Bazy = (value + avg_Dodawanie_Do_Bazy) / 2;
                }
            }
            public static TimeSpan Avg_Process_1_Zakladka
            {
                get => avg_Process_1_Zakladka;
                set
                {
                    if (avg_Process_1_Zakladka == TimeSpan.Zero)
                    {
                        avg_Process_1_Zakladka = value;
                        return;
                    }
                    avg_Process_1_Zakladka = (value + avg_Process_1_Zakladka) / 2;
                }
            }
            // TODO dodać reszte do pomiarów
            public static void Display_Times()
            {
                Console.WriteLine($"Pomiar.Avg_Process_Files: {Avg_Process_Files}");
                Console.WriteLine($"Pomiar.Avg_Process_1_Zakladka: {Avg_Process_1_Zakladka}");
                Console.WriteLine($"Pomiar.Avg_Open_Workbook: {Avg_Open_Workbook}");
                Console.WriteLine($"Pomiar.Avg_Get_Metadane_Pliku: {Avg_Get_Metadane_Pliku}");
                Console.WriteLine($"Pomiar.Avg_Get_Typ_Zakladki: {Avg_Get_Typ_Zakladki}");
                //Console.WriteLine($"Pomiar.Avg_Usun_Ukryte_Karty: {Avg_Usun_Ukryte_Karty}");
                Console.WriteLine($"Pomiar.Avg_Find_Starting_Points: {Avg_Find_Starting_Points}");
                Console.WriteLine($"Pomiar.Avg_Get_Dane_Z_Pliku: {Avg_Get_Dane_Z_Pliku}");
                Console.WriteLine($"Pomiar.Avg_Insert_Obecnosc_Command: {Avg_Insert_Obecnosc_Command}");
                Console.WriteLine($"Pomiar.avg_Insert_Atrybuty_Do_Optimy: {avg_Insert_Atrybuty_Do_Optimy}");
                Console.WriteLine($"Pomiar.avg_Create_Transaction: {avg_Create_Transaction}");
                Console.WriteLine($"Pomiar.avg_Dodawanie_Do_Bazy: {avg_Dodawanie_Do_Bazy}");
                Console.WriteLine($"Pomiar.Avg_MoveFile: {Avg_MoveFile}");
                Console.WriteLine($"Pomiar.Avg_Copy_Bad_Sheet_To_Files_Folder: {Avg_Copy_Bad_Sheet_To_Files_Folder}");
            }
        }
    }
}