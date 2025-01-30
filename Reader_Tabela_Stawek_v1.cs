using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace Konduktor_Reader
{
    internal static class Reader_Tabela_Stawek_v1
    {
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
            public decimal Podstawowa_Stawka_Godzinowa = 0;
            public decimal Podstawowe = 0;
            public decimal Wynagrodzenie_Za_Godz_Nadliczbowe = 0;
            public decimal Dodatek_Za_Pracę_W_Nocy = 0;
            public decimal Całkowite = 0;
            public decimal Dodatek_Wyjazdowy = 0;
        }
        public class Tabela_Stawek
        {
            public Czas_Relacji Czas_Relacji = new();
            public Wynagrodzenie Wynagrodzenie = new();
        }
        public static void Process_Zakladka(IXLWorksheet Zakladka)
        {
            List<Helper.Current_Position> Pozcje_Tabeli_Stawek_W_Zakladce = Helper.Find_Staring_Points_Tabele_Stawek(Zakladka, "Tabela Stawek");
            List<Relacja> Relacje = [];

            foreach (Helper.Current_Position pozycja in Pozcje_Tabeli_Stawek_W_Zakladce)
            {
                Relacja Relacja = new();
                Get_Dane(ref Relacja, pozycja, Zakladka);
                Relacje.Add(Relacja);
            }

            foreach(Relacja Relacja in Relacje)
            {
                Relacja.Insert_Relacja();
                Insert_Stawka(Relacja);
            }
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
                pozycja.Row += 2;
                int offest = Get_Dane_Relacji(ref Relacja, pozycja, Zakladka);
                pozycja.Row += offest - 1;
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
                Program.error_logger.New_Error(dane, "Opis Relacji", pozycja.Col + 1, pozycja.Row, "Brak opisu relacji");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            Relacja.Opis_Relacji_1 = dane;

            dane = Zakladka.Cell(pozycja.Row + 1, pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "Opis Relacji", pozycja.Col + 1, pozycja.Row + 1, "Brak opisu relacji");
                throw new Exception(Program.error_logger.Get_Error_String());
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
                    Program.error_logger.New_Error(dane, "Numer Relacji", pozycja.Col, pozycja.Row + offset, "Brak Numeru Relacji");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                if (dane.Count(c => c == '.') > 2)
                {
                    break;
                }
                System_Obsługi_Relacji System_Obsługi_Relacji = new();
                System_Obsługi_Relacji.Relacja.Numer_Relacji = dane;

                // opis relacji
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(dane))
                {
                    Program.error_logger.New_Error(dane, "Opis Relacji", pozycja.Col + 1, pozycja.Row + offset, "Brak Opisu Relacji");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                System_Obsługi_Relacji.Relacja.Opis_Relacji_1 = dane;

                // Czas całkowity
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Calkowity))
                {
                    Program.error_logger.New_Error(dane, "Czas Relacji Calkowity", pozycja.Col + 2, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Czas ogolem
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 3).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Ogolem))
                {
                    Program.error_logger.New_Error(dane, "Czas Relacji Ogolem", pozycja.Col + 3, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Czas podstawowy
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 4).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Podstawowe))
                {
                    Program.error_logger.New_Error(dane, "Czas Relacji Podstawowy", pozycja.Col + 4, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Godz nadl 50
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 5).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Godziny_Nadliczbowe_50))
                {
                    Program.error_logger.New_Error(dane, "Godziny Nadliczbowe 50", pozycja.Col + 5, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Godz nadl 100
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 6).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Godziny_Nadliczbowe_100))
                {
                    Program.error_logger.New_Error(dane, "Godziny Nadliczbowe 100", pozycja.Col + 6, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Godz pracy w nocy
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 7).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Godziny_Pracy_W_Nocy))
                {
                    Program.error_logger.New_Error(dane, "Godziny pracy w nocy", pozycja.Col + 7, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Czas odpoczynku
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 8).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Czas_Odpoczynku))
                {
                    Program.error_logger.New_Error(dane, "Czas odpoczynku", pozycja.Col + 8, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // podstawowa stawka godzinowa
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 9).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Podstawowa_Stawka_Godzinowa))
                {
                    Program.error_logger.New_Error(dane, "podstawowa stawka godzinowa", pozycja.Col + 9, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Wynagrodzenie ryczałtowe podstawowe
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 10).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Podstawowe))
                {
                    Program.error_logger.New_Error(dane, "Wynagrodzenie ryczałtowe podstawowe", pozycja.Col + 10, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Wynagrodzenie ryczałtowe za godz. nadliczbowe
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 11).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Wynagrodzenie_Za_Godz_Nadliczbowe))
                {
                    Program.error_logger.New_Error(dane, "Wynagrodzenie ryczałtowe za godz. nadliczbow", pozycja.Col + 11, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Dodatek za pracę w nocy
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 12).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Za_Pracę_W_Nocy))
                {
                    Program.error_logger.New_Error(dane, "Dodatek za pracę w nocy", pozycja.Col + 12, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Wynagrodzenie Calkowite
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 13).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Całkowite))
                {
                    Program.error_logger.New_Error(dane, "Wynagrodzenie Calkowite", pozycja.Col + 13, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Dodatek wyjazdowy
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 14).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Wyjazdowy))
                {
                    Program.error_logger.New_Error(dane, "Dodatek wyjazdowy", pozycja.Col + 14, pozycja.Row + offset);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                offset++;
                Relacja.System_Obsługi_Relacji.Add(System_Obsługi_Relacji);
            }
            return offset;
        }
        private static void Insert_Stawka(Relacja Relacja)
        {
            string query = @"INSERT INTO dbo.Stawki
                            (S_RId
                            ,S_Calkowity
                            ,S_Ogolem
                            ,S_Podstawowe
                            ,S_Godz_Nadliczbowe_50
                            ,S_Godz_Nadliczbowe_100
                            ,S_Czas_Odpoczynku
                            ,S_Podstawowa_Stawka_Godzinowa
                            ,S_Wynagrodzenie_Ryczaltowe_Podstawowe
                            ,S_Wynagrodzenie_Ryczaltowe_Za_Godz_Nadlicznowe
                            ,S_Wynagrodzenie_Ryczaltowe_Dodatek_Za_Prace_W_Nocy
                            ,S_Wynagrodzenie_Ryczaltowe_Calkowite
                            ,S_Dodatek_Wyjazdowy
                            ,S_Data_Mod
                            ,S_Os_Mod)
                        VALUES
                            (@RId
                            ,@Calkowity
                            ,@Ogolem
                            ,@Podstawowe
                            ,@Godz_Nadliczbowe_50
                            ,@Godz_Nadliczbowe_100
                            ,@Czas_Odpoczynku
                            ,@Podstawowa_Stawka_Godzinowa
                            ,@Wynagrodzenie_Ryczaltowe_Podstawowe
                            ,@Wynagrodzenie_Ryczaltowe_Za_Godz_Nadlicznowe
                            ,@Wynagrodzenie_Ryczaltowe_Dodatek_Za_Prace_W_Nocy
                            ,@Wynagrodzenie_Ryczaltowe_Calkowite
                            ,@Dodatek_Wyjazdowy
                            ,@Data_Mod
                            ,@Os_Mod)";
            using (SqlConnection connection = new(Program.Optima_Conection_String))
            {
                foreach (var System_Obsługi_Relacji in Relacja.System_Obsługi_Relacji)
                {
                    using (SqlCommand command = new(query, connection))
                    {
                        command.Parameters.AddWithValue("@RId", Relacja.Get_Relacja_Id());
                        command.Parameters.AddWithValue("@Calkowity", System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Calkowity);
                        command.Parameters.AddWithValue("@Ogolem", System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Ogolem);
                        command.Parameters.AddWithValue("@Podstawowe", System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Podstawowe);
                        command.Parameters.AddWithValue("@Godz_Nadliczbowe_50", System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Godziny_Nadliczbowe_50);
                        command.Parameters.AddWithValue("@Godz_Nadliczbowe_100", System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Godziny_Nadliczbowe_100);
                        command.Parameters.AddWithValue("@Czas_Odpoczynku", System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Czas_Odpoczynku);
                        command.Parameters.AddWithValue("@Podstawowa_Stawka_Godzinowa", System_Obsługi_Relacji.Tabela_Stawek.Czas_Relacji.Calkowity);
                        command.Parameters.AddWithValue("@Wynagrodzenie_Ryczaltowe_Podstawowe", System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Podstawowe);
                        command.Parameters.AddWithValue("@Wynagrodzenie_Ryczaltowe_Za_Godz_Nadlicznowe", System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Wynagrodzenie_Za_Godz_Nadliczbowe);
                        command.Parameters.AddWithValue("@Wynagrodzenie_Ryczaltowe_Dodatek_Za_Prace_W_Nocy", System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Za_Pracę_W_Nocy);
                        command.Parameters.AddWithValue("@Wynagrodzenie_Ryczaltowe_Calkowite", System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Całkowite);
                        command.Parameters.AddWithValue("@Dodatek_Wyjazdowy", System_Obsługi_Relacji.Tabela_Stawek.Wynagrodzenie.Dodatek_Wyjazdowy);
                        command.Parameters.AddWithValue("@Data_Mod", DateTime.Now);
                        command.Parameters.AddWithValue("@Os_Mod", "Norbert Tasarz");
                        connection.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

    }
}