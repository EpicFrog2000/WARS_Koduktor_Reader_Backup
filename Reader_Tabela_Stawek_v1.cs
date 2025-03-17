using System.Data;
using System.Diagnostics;
using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace Excel_Data_Importer_WARS
{
    internal static class Reader_Tabela_Stawek_v1
    {
        // JA JEBE KURWA PRZECIEZ TO BĘDZIE ZMIENIANE Z 500 GORYLIONÓW RAZY WSZYSTKO PORA SIE ZAJEBAĆ
        public class System_Obsługi_Relacji
        {
            public Relacja Relacja = new();
            public Tabela_Stawek Tabela_Stawek = new();
        }

        public class Czas_Relacji
        {
            public decimal Calkowity = 0;
            public decimal Ogolem = 0;
            public decimal Podstawowe = 0;
            public decimal Godziny_Nadliczbowe_50 = 0;
            public decimal Godziny_Nadliczbowe_100 = 0;
            public decimal Godziny_Pracy_W_Nocy = 0;
            public decimal Czas_Odpoczynku = 0;
        }

        public class Wynagrodzenie
        {
            public decimal Podstawowa_Stawka_Godzinowa = -1;
            public decimal Podstawowe = -1;
            public decimal Wynagrodzenie_Za_Godz_Nadliczbowe = -1;
            public decimal Dodatek_Za_Pracę_W_Nocy = -1;
            public decimal Całkowite = -1;
            public decimal Dodatek_Wyjazdowy = -1;
        }

        public class Tabela_Stawek
        {
            public Czas_Relacji Czas_Relacji = new();
            public Wynagrodzenie Wynagrodzenie = new();
        }
        private static Error_Logger Internal_Error_Logger = new(true);
        public static async Task Process_Zakladka(IXLWorksheet Zakladka, Error_Logger Error_Logger)
        {
            Internal_Error_Logger = Error_Logger;
            List<Helper.Current_Position> Pozcje_Tabeli_Stawek_W_Zakladce = Helper.Find_Starting_Points(Zakladka, "Tabela Stawek");
            List<Relacja> Relacje = [];

            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();

            foreach (Helper.Current_Position pozycja in Pozcje_Tabeli_Stawek_W_Zakladce)
            {
                Relacja Relacja = new();
                Get_Dane(ref Relacja, pozycja, Zakladka);
                Relacje.Add(Relacja);
                
            }
            Helper.Pomiar.Avg_Get_Dane_Z_Pliku = PomiaryStopWatch.Elapsed;

            PomiaryStopWatch.Restart();
            await Insert_Dane_Stawek_Do_Optimy(Relacje);
            Helper.Pomiar.Avg_Dodawanie_Do_Bazy = PomiaryStopWatch.Elapsed;

        }
        private static void Get_Dane(ref Relacja Relacja, Helper.Current_Position pozycja, IXLWorksheet Zakladka)
        {
            pozycja.Col -= 2;
            pozycja.Row += 7;
            while (true)
            {
                // NR relacji/harmonogramu
                string dane = Zakladka.Cell(pozycja.Row, pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(dane))
                {
                    return;
                }
                Get_Relacja(ref Relacja, pozycja, Zakladka);
                if (string.IsNullOrEmpty(Relacja.Opis_Relacji_1))
                {
                    return;
                }
                pozycja.Row += 2;
                int offest = Get_Dane_Relacji(ref Relacja, pozycja, Zakladka);
                pozycja.Row += offest;
            }
        }
        private static void Get_Relacja(ref Relacja Relacja, Helper.Current_Position pozycja, IXLWorksheet Zakladka)
        {
            string dane = Zakladka.Cell(pozycja.Row, pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                return;
            }
            Relacja.Numer_Relacji = dane;

            dane = Zakladka.Cell(pozycja.Row, pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Internal_Error_Logger.New_Error(dane, "Opis Relacji", pozycja.Col + 1, pozycja.Row, "Brak opisu relacji");
            }
            Relacja.Opis_Relacji_1 = dane;

            dane = Zakladka.Cell(pozycja.Row + 1, pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Internal_Error_Logger.New_Error(dane, "Opis Relacji", pozycja.Col + 1, pozycja.Row + 1, "Brak opisu relacji");
            }
            Relacja.Opis_Relacji_2 = dane;
        }
        private static int Get_Dane_Relacji(ref Relacja Relacja, Helper.Current_Position pozycja, IXLWorksheet Zakladka)
        {
            int offset = 0;
            while (true)
            {
                // numer relacji
                string dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(dane))
                {
                    //Internal_Error_Logger.New_Error(dane, "PodNumer Relacji", pozycja.Col, pozycja.Row + offset, "Brak Numeru Relacji");
                    //throw new Exception(Internal_Error_Logger.Get_Error_String());
                    offset++;
                    break;
                }
                if (dane.Count(c => c == '.') < 3)
                {
                    break;
                }
                System_Obsługi_Relacji System_Obsługi_Relacji = new();
                System_Obsługi_Relacji.Relacja.Numer_Relacji = dane;

                // opis relacji
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(dane))
                {
                    Internal_Error_Logger.New_Error(dane, "Opis Relacji", pozycja.Col + 1, pozycja.Row + offset, "Brak Opisu Relacji");
                }
                System_Obsługi_Relacji.Relacja.Opis_Relacji_1 = dane;

                // Wynagrodzenie ryczałtowe podstawowe
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 10).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Podstawowe))
                {
                    Internal_Error_Logger.New_Error(dane, "Wynagrodzenie ryczałtowe podstawowe", pozycja.Col + 10, pozycja.Row + offset);
                }

                // Wynagrodzenie ryczałtowe za godz. nadliczbowe
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 11).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Wynagrodzenie_Za_Godz_Nadliczbowe))
                {
                    Internal_Error_Logger.New_Error(dane, "Wynagrodzenie ryczałtowe za godz. nadliczbow", pozycja.Col + 11, pozycja.Row + offset);
                }

                // Dodatek za pracę w nocy
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 12).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Za_Pracę_W_Nocy))
                {
                    Internal_Error_Logger.New_Error(dane, "Dodatek za pracę w nocy", pozycja.Col + 12, pozycja.Row + offset);
                }

                // Wynagrodzenie Calkowite
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 13).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Całkowite))
                {
                    Internal_Error_Logger.New_Error(dane, "Wynagrodzenie Calkowite", pozycja.Col + 13, pozycja.Row + offset);
                }

                // Dodatek wyjazdowy
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 14).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Wyjazdowy))
                {
                    Internal_Error_Logger.New_Error(dane, "Dodatek wyjazdowy", pozycja.Col + 14, pozycja.Row + offset);
                }

                offset++;
                Relacja.System_Obsługi_Relacji.Add(System_Obsługi_Relacji);
            }
            return offset;
        }
        private static async Task Insert_Dane_Stawek_Do_Optimy(List<Relacja> Relacje)
        {
            try
            {
                await DbManager.Transaction_Manager.Create_Transaction();
                foreach (Relacja Relacja in Relacje)
                {
                    try
                    {
                        Relacja.Insert_Relacja_Do_Optimy(Internal_Error_Logger);
                    }
                    catch (SqlException ex)
                    {
                        Internal_Error_Logger.New_Custom_Error($"Error podczas operacji w bazie (Insert_Relacja_Do_Optimy): {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        Internal_Error_Logger.New_Custom_Error($"Error: {ex.Message}");
                    }

                    foreach (System_Obsługi_Relacji System_Obsługi_Relacji in Relacja.System_Obsługi_Relacji)
                    {
                        try
                        {
                            System_Obsługi_Relacji.Relacja.Insert_Relacja_Do_Optimy(Internal_Error_Logger);
                        }
                        catch (SqlException ex)
                        {
                            Internal_Error_Logger.New_Custom_Error($"Error podczas operacji w bazie (System_Obsługi_Relacji.Relacja.Insert_Relacja_Do_Optimy): {ex.Message}");
                        }
                        catch (Exception ex)
                        {
                            Internal_Error_Logger.New_Custom_Error($"Error w programie (System_Obsługi_Relacji.Relacja.Insert_Relacja_Do_Optimy): {ex.Message}");
                        }
                    }
                    Insert_Atrybuty_Do_Optimy(Relacja);
                }
                DbManager.Transaction_Manager.Commit_Transaction();
            }
            catch (SqlException ex)
            {
                Internal_Error_Logger.New_Custom_Error($"Error podczas operacji w bazie(Insert_Dane_Stawek_Do_Optimy): {ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}", false);
                DbManager.Transaction_Manager.RollBack_Transaction();
                throw new Exception($"Error podczas operacji w bazie(Insert_Dane_Stawek_Do_Optimy): {ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
            }
            catch (Exception ex)
            {
                Internal_Error_Logger.New_Custom_Error($"Error: {ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}", false);
                DbManager.Transaction_Manager.RollBack_Transaction();
                throw new Exception($"Error: {ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
            }
        }
        private static void Insert_Atrybuty_Do_Optimy(Relacja Relacja)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();


            DateTime Data_Od = DateTime.ParseExact("2025.02.01 00:00:00", "yyyy.MM.dd HH:mm:ss", CultureInfo.InvariantCulture);
            DateTime Data_Do = DateTime.ParseExact("2025.03.01 00:00:00", "yyyy.MM.dd HH:mm:ss", CultureInfo.InvariantCulture).AddDays(-1);

            int counter = 0;
            foreach (System_Obsługi_Relacji System_Obsługi_Relacji in Relacja.System_Obsługi_Relacji)
            {
                if (System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Podstawowe != -1)
                {
                    counter += Insert_Command_Atrybuty(System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Podstawowe.ToString(), "Wynagrodzenie ryczałtowe - Podstawowe", Data_Od, Data_Do);
                }
                if (System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Wynagrodzenie_Za_Godz_Nadliczbowe != -1)
                {
                    counter += Insert_Command_Atrybuty(System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Wynagrodzenie_Za_Godz_Nadliczbowe.ToString(), "Wynagrodzenie ryczałtowe - Nadgodziny", Data_Od, Data_Do);
                }
                if (System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Za_Pracę_W_Nocy != -1)
                {
                    counter += Insert_Command_Atrybuty(System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Za_Pracę_W_Nocy.ToString(), "Wynagrodzenie ryczałtowe - Nocki", Data_Od, Data_Do);
                }
                if (System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Wyjazdowy != -1)
                {
                    counter += Insert_Command_Atrybuty(System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Wyjazdowy.ToString(), "Dodatek wyjazdowy", Data_Od, Data_Do);
                }
            }
            if (counter > 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Poprawnie dodano dane z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                Console.ForegroundColor = ConsoleColor.White;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Nie dodano żadnych danych pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                Console.ForegroundColor = ConsoleColor.White;
            }
            Helper.Pomiar.Avg_Insert_Atrybuty_Do_Optimy = PomiaryStopWatch.Elapsed;
        }
        private static int Insert_Command_Atrybuty(string wartosc, string Nazwa_Atrybutu, DateTime Data_Od, DateTime Data_Do)
        {
            try
            {
                using (SqlCommand command = new(DbManager.Insert_Atrybuty, DbManager.GetConnection(), DbManager.Transaction_Manager.CurrentTransaction))
                {
                    command.Parameters.Add("@NowaWartosc", SqlDbType.NVarChar, 101).Value = wartosc;
                    command.Parameters.Add("@NazwaAtrybutu", SqlDbType.NVarChar, 100).Value = Nazwa_Atrybutu;
                    command.Parameters.Add("@ATHDataOd", SqlDbType.DateTime).Value = Data_Od;
                    command.Parameters.Add("@ATHDataDo", SqlDbType.DateTime).Value = Data_Do;
                    command.ExecuteNonQuery();
                }
                return 1;
            }
            catch (SqlException ex)
            {
                Internal_Error_Logger.New_Custom_Error($"Error podczas operacji w bazie(Insert_Command_Atrybuty): {ex.Message}");
            }
            catch (Exception ex)
            {
                Internal_Error_Logger.New_Custom_Error($"Error: {ex.Message}");
            }
            return 0;
        }
    }
}