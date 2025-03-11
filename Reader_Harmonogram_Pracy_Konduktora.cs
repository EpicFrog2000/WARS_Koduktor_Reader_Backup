using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Microsoft.Data.SqlClient;

namespace Excel_Data_Importer_WARS
{
    internal class Reader_Harmonogram_Pracy_Konduktora
    {
        private class Harmonogram_Pracy_Konduktora
        {
            public Pracownik Konduktor = new();
            public int Miesiac = 0;
            public int Rok = 0;
            public List<Dane_Harmonogramu> Dane_Harmonogramu = [];
            public void Set_Miesiac(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return;
                }
                value = value.ToLower().Trim();
                Dictionary<int, string> months = new()
                {
                    {1, "styczeń"}, {2, "luty"}, {3, "marzec"}, {4, "kwiecień"},
                    {5, "maj"}, {6, "czerwiec"}, {7, "lipiec"}, {8, "sierpień"},
                    {9, "wrzesień"}, {10, "październik"}, {11, "listopad"}, {12, "grudzień"}
                };
                Miesiac = months.FirstOrDefault(kvp => value.Contains(kvp.Value)).Key;
            }
        }

        private class Dane_Harmonogramu
        {
            public Relacja Relacja = new();
            public string Czynnosc = string.Empty;
            public Absencja Absencja = new();
            public int Dzien = 0;
            public TimeSpan Godzina_Rozpoczecia_Pracy = TimeSpan.Zero;
            public TimeSpan Godzina_Zakonczenia_Pracy = TimeSpan.Zero;
            public TimeSpan Czas_Pracy_Poza_Relacja_Od = TimeSpan.Zero;
            public TimeSpan Czas_Pracy_Poza_Relacja_Do = TimeSpan.Zero;
            public TimeSpan Czas_Odpoczynku_Wliczany_Do_CP_Od = TimeSpan.Zero;
            public TimeSpan Czas_Odpoczynku_Wliczany_Do_CP_Do = TimeSpan.Zero;
            public TimeSpan Czas_Odpoczynku_Nie_Wliczany_Do_CP_Od = TimeSpan.Zero;
            public TimeSpan Czas_Odpoczynku_Nie_Wliczany_Do_CP_Do = TimeSpan.Zero;
        }
        private static Error_Logger Internal_Error_Logger = new(true);
        public static void Process_Zakladka(IXLWorksheet Zakladka, Error_Logger error_Logger)
        {
            Internal_Error_Logger = error_Logger;

            List<Harmonogram_Pracy_Konduktora> Harmonogramy_Pracy_Konduktora = [];
            List<Helper.Current_Position> Pozycje = Helper.Find_Starting_Points(Zakladka, "Dzień miesiąca");

            foreach (Helper.Current_Position Pozycja in Pozycje)
            {
                Harmonogram_Pracy_Konduktora Harmonogram_Pracy_Konduktora = new();
                Get_Dane_Naglowka(Zakladka, Pozycja, ref Harmonogram_Pracy_Konduktora);
                Get_Dane_Harmonogramu(Zakladka, Pozycja, ref Harmonogram_Pracy_Konduktora);
                Harmonogramy_Pracy_Konduktora.Add(Harmonogram_Pracy_Konduktora);
            }

            foreach (Harmonogram_Pracy_Konduktora Harmonogram_Pracy_Konduktora in Harmonogramy_Pracy_Konduktora)
            {
                Dodaj_Dane_Do_Optimy(Harmonogram_Pracy_Konduktora);
            }
        }
        private static void Get_Dane_Naglowka(IXLWorksheet Zakladka, Helper.Current_Position pozycja, ref Harmonogram_Pracy_Konduktora Harmonogram_Pracy_Konduktora)
        {
            string dane = Zakladka.Cell(pozycja.Row - 2, pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Internal_Error_Logger.New_Error(dane, "Imie Nazwisko", pozycja.Col + 2, pozycja.Row - 1, "Brak Imienia i Nazwiska");
            }
            else
            {
                string[] Sdane = dane.Split(' ');
                if (Sdane.Length == 2)
                {
                    Harmonogram_Pracy_Konduktora.Konduktor.Imie = Sdane[0];
                    Harmonogram_Pracy_Konduktora.Konduktor.Nazwisko = Sdane[1];
                }
                else
                {
                    Internal_Error_Logger.New_Error(dane, "Imie Nazwisko", pozycja.Col + 2, pozycja.Row - 1, "Zły format w polu Imienia i Nazwiska");
                }
            }


            dane = Zakladka.Cell(pozycja.Row - 1, pozycja.Col + 4).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Internal_Error_Logger.New_Error(dane, "Data", pozycja.Col + 4, pozycja.Row - 1, "Brak Daty");
            }
            else
            {
                string[] Sdane = dane.Split(' ');
                if (Sdane.Length == 2)
                {
                    Harmonogram_Pracy_Konduktora.Set_Miesiac(Sdane[0]);
                    if (Harmonogram_Pracy_Konduktora.Miesiac == 0)
                    {
                        Internal_Error_Logger.New_Error(dane, "Data(Miesiac)", pozycja.Col + 4, pozycja.Row - 1, "Zły format w polu Data (Miesiac)");
                    }
                    if (!Helper.Try_Get_Type_From_String<int>(Sdane[1], ref Harmonogram_Pracy_Konduktora.Rok))
                    {
                        Internal_Error_Logger.New_Error(dane, "Data(Rok)", pozycja.Col + 4, pozycja.Row - 1, "Zły format w polu Data (Rok)");
                    }
                }
                else
                {
                    Internal_Error_Logger.New_Error(dane, "Data", pozycja.Col + 4, pozycja.Row - 1, "Zły format w polu Data");
                }
            }
        }
        private static void Get_Dane_Harmonogramu(IXLWorksheet Zakladka, Helper.Current_Position pozycja, ref Harmonogram_Pracy_Konduktora Harmonogram_Pracy_Konduktora)
        {
            int offset = 0;
            pozycja.Row += 4;
            string Dzien = string.Empty;
            do
            {
                Dzien = Zakladka.Cell(pozycja.Row + offset, pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(Dzien))
                {
                    break;
                }
                Dane_Harmonogramu Dane_Harmonogramu = new();
                if (!Helper.Try_Get_Type_From_String<int>(Dzien, ref Dane_Harmonogramu.Dzien))
                {
                    Internal_Error_Logger.New_Error(Dzien, "Dzien", pozycja.Col, pozycja.Row + offset, "Zły format w polu Dzien miesiaca");
                }
                string dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    string dane2 = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(Dzien))
                    {
                        Dane_Harmonogramu.Relacja.Numer_Relacji = dane2;
                        Dane_Harmonogramu.Relacja.Opis_Relacji_1 = dane;
                    }
                    else
                    {
                        Dane_Harmonogramu.Absencja.Nazwa = dane;
                        if (!Absencja.RodzajAbsencji.TryParse(Dane_Harmonogramu.Absencja.Nazwa, out Dane_Harmonogramu.Absencja.Rodzaj_Absencji))
                        {
                            Internal_Error_Logger.New_Error(dane, "Rodzaj Absencji", pozycja.Col + 1, pozycja.Row + offset, "Nierozpoznany rodzaj absencji");
                        }
                        Dane_Harmonogramu.Absencja.Dzien = Dane_Harmonogramu.Dzien;
                        Dane_Harmonogramu.Absencja.Miesiac = Harmonogram_Pracy_Konduktora.Miesiac;
                        Dane_Harmonogramu.Absencja.Rok = Harmonogram_Pracy_Konduktora.Rok;
                        Console.WriteLine(Dane_Harmonogramu.Absencja.Nazwa);
                    }
                }

                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 4).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref Dane_Harmonogramu.Godzina_Rozpoczecia_Pracy))
                    {
                        Internal_Error_Logger.New_Error(dane, "Godzina Rozpoczecia Pracy", pozycja.Col + 4, pozycja.Row + offset, "Zły format");
                    }
                }

                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 5).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref Dane_Harmonogramu.Godzina_Zakonczenia_Pracy))
                    {
                        Internal_Error_Logger.New_Error(dane, "Godzina Zakonczenia Pracy", pozycja.Col + 5, pozycja.Row + offset, "Zły format");
                    }
                }

                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 8).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref Dane_Harmonogramu.Czas_Pracy_Poza_Relacja_Od))
                    {
                        Internal_Error_Logger.New_Error(dane, "Czas Pracy Poza Relacja Od", pozycja.Col + 8, pozycja.Row + offset, "Zły format");
                    }
                }

                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 9).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref Dane_Harmonogramu.Czas_Pracy_Poza_Relacja_Do))
                    {
                        Internal_Error_Logger.New_Error(dane, "Czas Pracy Poza Relacja Do", pozycja.Col + 9, pozycja.Row + offset, "Zły format");
                    }
                }

                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 10).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref Dane_Harmonogramu.Czas_Odpoczynku_Wliczany_Do_CP_Od))
                    {
                        Internal_Error_Logger.New_Error(dane, "Czas Odpoczynku Wliczany Do CP Od", pozycja.Col + 10, pozycja.Row + offset, "Zły format");
                    }
                }

                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 11).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref Dane_Harmonogramu.Czas_Odpoczynku_Wliczany_Do_CP_Do))
                    {
                        Internal_Error_Logger.New_Error(dane, "Czas Odpoczynku Wliczany Do CP Do", pozycja.Col + 11, pozycja.Row + offset, "Zły format");
                    }
                }

                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 12).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref Dane_Harmonogramu.Czas_Odpoczynku_Nie_Wliczany_Do_CP_Od))
                    {
                        Internal_Error_Logger.New_Error(dane, "Czas Odpoczynku Nie Wliczany Do CP Od", pozycja.Col + 12, pozycja.Row + offset, "Zły format");
                    }
                }

                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 13).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref Dane_Harmonogramu.Czas_Odpoczynku_Nie_Wliczany_Do_CP_Do))
                    {
                        Internal_Error_Logger.New_Error(dane, "Czas Odpoczynku Nie Wliczany Do CP Do", pozycja.Col + 13, pozycja.Row + offset, "Zły format");
                    }
                }
                Harmonogram_Pracy_Konduktora.Dane_Harmonogramu.Add(Dane_Harmonogramu);
                offset++;
            } while (!string.IsNullOrEmpty(Dzien));
        }
        private static void Dodaj_Dane_Do_Optimy(Harmonogram_Pracy_Konduktora Harmonogram_Pracy_Konduktora)
        {
            using (SqlConnection connection = new(DbManager.Connection_String))
            {
                connection.Open();
                using (SqlTransaction transaction = connection.BeginTransaction(System.Data.IsolationLevel.ReadUncommitted))
                {

                    // insert absencje
                    List<Absencja> Absencje = [];
                    foreach (Dane_Harmonogramu Dane_Harmonogramu in Harmonogram_Pracy_Konduktora.Dane_Harmonogramu)
                    {
                        Absencje.Add(Dane_Harmonogramu.Absencja);
                    }
                    if (Absencja.Dodaj_Absencje_do_Optimy(Absencje, transaction, connection, Harmonogram_Pracy_Konduktora.Konduktor, Internal_Error_Logger) > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Poprawnie dodano absencje z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Nie dodano żadnych absencji");
                        Console.ForegroundColor = ConsoleColor.White;
                    }

                    if (Dodaj_Plany_do_Optimy(Harmonogram_Pracy_Konduktora, connection, transaction) > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Poprawnie dodano plan pracy z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Nie dodano godzin planu pracy");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                }
            }
        }
        private static int Dodaj_Plany_do_Optimy(Harmonogram_Pracy_Konduktora Harmonogram_Pracy_Konduktora, SqlConnection connection, SqlTransaction transaction)
        {
            int dodano = 0;
            HashSet<DateTime> Pasujace_Daty = [];
            foreach (Dane_Harmonogramu Dane_Harmonogramu in Harmonogram_Pracy_Konduktora.Dane_Harmonogramu)
            {
                Pasujace_Daty.Add(new DateTime(Harmonogram_Pracy_Konduktora.Rok, Harmonogram_Pracy_Konduktora.Miesiac, Dane_Harmonogramu.Dzien));
            }
            DateTime startDate = new(Harmonogram_Pracy_Konduktora.Rok, Harmonogram_Pracy_Konduktora.Miesiac, 1);
            DateTime endDate = new(Harmonogram_Pracy_Konduktora.Rok, Harmonogram_Pracy_Konduktora.Miesiac, DateTime.DaysInMonth(Harmonogram_Pracy_Konduktora.Rok, Harmonogram_Pracy_Konduktora.Miesiac));
            for (DateTime dzien = startDate; dzien <= endDate; dzien = dzien.AddDays(1))
            {
                if (!Pasujace_Daty.Contains(dzien))
                {
                    try
                    {
                        Zrob_Insert_Plan_command(connection, transaction, Harmonogram_Pracy_Konduktora.Konduktor, DateTime.ParseExact($"{Harmonogram_Pracy_Konduktora.Rok}-{Harmonogram_Pracy_Konduktora.Miesiac:D2}-{dzien.Day:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture), TimeSpan.Zero, TimeSpan.Zero, Helper.Strefa.undefined, "");
                    }
                    catch (SqlException ex)
                    {
                        transaction.Rollback();
                        Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                    }
                    catch (FormatException ex)
                    {
                        transaction.Rollback();
                        Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                    }
                }
            }

            foreach (Dane_Harmonogramu Dane_Harmonogramu in Harmonogram_Pracy_Konduktora.Dane_Harmonogramu)
            {
                try
                {
                    // TODO inne strefy do innych godzin
                    dodano += Zrob_Insert_Plan_command(connection, transaction, Harmonogram_Pracy_Konduktora.Konduktor, DateTime.ParseExact($"{Harmonogram_Pracy_Konduktora.Rok}-{Harmonogram_Pracy_Konduktora.Miesiac:D2}-{Dane_Harmonogramu.Dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture), Dane_Harmonogramu.Godzina_Rozpoczecia_Pracy, Dane_Harmonogramu.Godzina_Zakonczenia_Pracy, Helper.Strefa.Czas_Pracy_Podstawowy, Dane_Harmonogramu.Relacja.Numer_Relacji);
                    dodano += Zrob_Insert_Plan_command(connection, transaction, Harmonogram_Pracy_Konduktora.Konduktor, DateTime.ParseExact($"{Harmonogram_Pracy_Konduktora.Rok}-{Harmonogram_Pracy_Konduktora.Miesiac:D2}-{Dane_Harmonogramu.Dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture), Dane_Harmonogramu.Czas_Pracy_Poza_Relacja_Od, Dane_Harmonogramu.Czas_Pracy_Poza_Relacja_Do, Helper.Strefa.Czas_Pracy_Poza_Relacją, Dane_Harmonogramu.Relacja.Numer_Relacji);
                    dodano += Zrob_Insert_Plan_command(connection, transaction, Harmonogram_Pracy_Konduktora.Konduktor, DateTime.ParseExact($"{Harmonogram_Pracy_Konduktora.Rok}-{Harmonogram_Pracy_Konduktora.Miesiac:D2}-{Dane_Harmonogramu.Dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture), Dane_Harmonogramu.Czas_Odpoczynku_Wliczany_Do_CP_Od, Dane_Harmonogramu.Czas_Odpoczynku_Wliczany_Do_CP_Do, Helper.Strefa.Odpoczynek_Czas_Odpoczynku_Wliczany_Do_CP, Dane_Harmonogramu.Relacja.Numer_Relacji);
                    dodano += Zrob_Insert_Plan_command(connection, transaction, Harmonogram_Pracy_Konduktora.Konduktor, DateTime.ParseExact($"{Harmonogram_Pracy_Konduktora.Rok}-{Harmonogram_Pracy_Konduktora.Miesiac:D2}-{Dane_Harmonogramu.Dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture), Dane_Harmonogramu.Czas_Odpoczynku_Nie_Wliczany_Do_CP_Od, Dane_Harmonogramu.Czas_Odpoczynku_Nie_Wliczany_Do_CP_Do, Helper.Strefa.Czas_Odpoczynku_Nie_Wliczany_Do_CP, Dane_Harmonogramu.Relacja.Numer_Relacji);
                }
                catch (SqlException ex)
                {
                    transaction.Rollback();
                    Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                }
                catch (FormatException ex)
                {

                    transaction.Rollback();
                    Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                }
            }
            
            transaction.Commit();
            return dodano;
        }
        private static int Zrob_Insert_Plan_command(SqlConnection connection, SqlTransaction transaction, Pracownik pracownik, DateTime data, TimeSpan startGodz, TimeSpan endGodz, Helper.Strefa Strefa, string Numer_Relacji)
        {
            int IdPracownika = pracownik.Get_PraId(connection, transaction);

            if (startGodz != TimeSpan.Zero && endGodz != TimeSpan.Zero)
            {
                using (SqlCommand command = new(DbManager.Check_Duplicate_Plan_Pracy, connection, transaction))
                {
                    command.Parameters.AddWithValue("@DataInsert", data + TimeSpan.FromSeconds(0));
                    command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = (DateTime)(DbManager.Base_Date + startGodz);
                    command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = (DateTime)(DbManager.Base_Date + endGodz);
                    command.Parameters.AddWithValue("@PRI_PraId", IdPracownika);
                    if ((int)command.ExecuteScalar() == 1)
                    {
                        return 0;
                    }
                }
            }

            using (SqlCommand command = new(DbManager.Insert_Plan_Pracy_Z_Relacja, connection, transaction))
            {
                command.Parameters.AddWithValue("@DataInsert", data);
                command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = (DateTime)(DbManager.Base_Date + startGodz);
                command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = (DateTime)(DbManager.Base_Date + endGodz);
                command.Parameters.AddWithValue("@PRI_PraId", IdPracownika);
                command.Parameters.Add("@Strefa", SqlDbType.Int).Value = Strefa;
                command.Parameters.Add("@NumerRelacji", SqlDbType.NVarChar, 100).Value = Numer_Relacji;

                command.Parameters.AddWithValue("@ImieMod", Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20));
                command.Parameters.AddWithValue("@NazwiskoMod", Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 50));
                command.Parameters.AddWithValue("@DataMod", Internal_Error_Logger.Last_Mod_Time);
                command.ExecuteScalar();
            }
            return 1;
        }
    }
}
// TODO NA BAZIE TESTOWEJ ORAZ PROD:
//ALTER TABLE cdn.PracPlanDni
//ADD PPL_Relacja VARCHAR(100);
