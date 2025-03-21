﻿using System.Data;
using System.Diagnostics;
using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace Excel_Data_Importer_WARS
{
    internal class Reader_Grafik_Pracy_Pracownika_2025_v3
    {
        private class Grafik
        {
            public Pracownik Pracownik  = new();
            public int Miesiac  = 0;
            public int Rok  = 0;
            public List<Dane_Dnia> Dane_Dni  = [];
            public string Nazwa_Pliku = string.Empty;
            public int Nr_Zakladki = 1;
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
        private class Dane_Dnia
        {
            public int Nr_Dnia  = 0;
            public TimeSpan Godzina_Pracy_Od  = TimeSpan.Zero;
            public TimeSpan Godzina_Pracy_Do  = TimeSpan.Zero;
        }
        public static Error_Logger Internal_Error_Logger = new(true);
        public static async Task Process_Zakladka(IXLWorksheet Zakladka, Error_Logger error_Logger)
        {
            Internal_Error_Logger = error_Logger;
            List<Helper.Current_Position> Lista_Pozycji_Grafików_Z_Zakladki = Helper.Find_Starting_Points(Zakladka, "Data");
            List<Grafik> grafiki = [];

            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();

            foreach (Helper.Current_Position Startpozycja in Lista_Pozycji_Grafików_Z_Zakladki)
            {
                Helper.Current_Position pozycja = Startpozycja;
                int counter = 0;
                while (true)
                {
                    int rowOffset = -5;
                    string dane = string.Empty;
                    Grafik grafik = new();
                    try
                    {
                        dane = Zakladka.Cell(pozycja.Row + rowOffset, pozycja.Col + 3).GetFormattedString().Trim();
                    }
                    catch
                    {
                        rowOffset = -4;
                        try
                        {
                            dane = Zakladka.Cell(pozycja.Row + rowOffset, pozycja.Col + 3).GetFormattedString().Trim();
                        }
                        catch
                        {
                            rowOffset = -3;
                            try
                            {
                                dane = Zakladka.Cell(pozycja.Row + rowOffset, pozycja.Col + 3).GetFormattedString().Trim();
                            }
                            catch
                            {
                                Internal_Error_Logger.New_Error(dane, "Naglowek", pozycja.Col + 3, pozycja.Row + rowOffset, "Zły format pliku");
                            }
                        }
                    }// xddddddddddd
                    grafik.Set_Miesiac(dane);
                    if (grafik.Miesiac < 1)
                    {
                        Internal_Error_Logger.New_Error(dane, "Miesiac", pozycja.Col + 3, pozycja.Row + rowOffset, "Błędna wartość w miesiac");
                    }


                    dane = Zakladka.Cell(pozycja.Row + rowOffset, pozycja.Col + 6).GetFormattedString().Trim();
                    if (string.IsNullOrEmpty(dane))
                    {
                        Internal_Error_Logger.New_Error(dane, "Rok", pozycja.Col + 6, pozycja.Row + rowOffset, "Brak wartości w komórce");
                    }

                    if (int.TryParse(dane, out int tmprok))
                    {
                        if (tmprok < 1900)
                        {
                            Internal_Error_Logger.New_Error(dane, "Rok", pozycja.Col + 6, pozycja.Row + rowOffset, "Błędna wartość w rok");
                        }
                        grafik.Rok = tmprok;
                    }
                    else
                    {
                        Internal_Error_Logger.New_Error(dane, "Rok", pozycja.Col + 6, pozycja.Row + rowOffset, "Błędna wartość w rok");
                    }

                    grafik.Pracownik = Get_Pracownik(Zakladka, new Helper.Current_Position(Startpozycja.Col + ((counter * 3) + 1), Startpozycja.Row));
                    if (string.IsNullOrEmpty(grafik.Pracownik.Imie) && string.IsNullOrEmpty(grafik.Pracownik.Nazwisko) && string.IsNullOrEmpty(grafik.Pracownik.Akronim))
                    {
                        break;
                    }

                    List<Dane_Dnia> dane2 = Get_Dane_Dni(Zakladka, new Helper.Current_Position(Startpozycja.Col + ((counter * 3) + 1), Startpozycja.Row + 4));
                    foreach (Dane_Dnia d in dane2)
                    {
                        grafik.Dane_Dni.Add(d);
                    }
                    grafiki.Add(grafik);
                    counter++;
                }
            }
            Helper.Pomiar.Avg_Get_Dane_Z_Pliku = PomiaryStopWatch.Elapsed;


            if (grafiki.Count > 0)
            {
                PomiaryStopWatch.Restart();
                await Dodaj_Plany_do_Optimy(grafiki);
                Helper.Pomiar.Avg_Dodawanie_Do_Bazy = PomiaryStopWatch.Elapsed;
            }
            else
            {
                Internal_Error_Logger.New_Custom_Error($"Zły format pliku, nie znaleniono żadnych grafików z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
            }

        }
        private static Pracownik Get_Pracownik(IXLWorksheet worksheet, Helper.Current_Position pozycja)
        {
            Pracownik pracownik = new();
            string pole1;
            string pole2 = string.Empty;
            int offset = 0;
            while (true)
            {
                try
                {
                    pozycja.Row--;
                    if (pozycja.Row < 1)
                    {
                        return pracownik;
                    }
                    try
                    {
                        pole1 = worksheet.Cell(pozycja.Row, pozycja.Col).GetFormattedString().Trim();
                    }
                    catch
                    {
                        return new Pracownik();
                    }
                    if (pole1 != "Godziny pracy od")
                    {
                        offset++;
                        for (int i = 0; i < 3; i++)
                        {
                            pole1 = worksheet.Cell(pozycja.Row, pozycja.Col + i).GetFormattedString().Trim();
                            if (!string.IsNullOrEmpty(pole1))
                            {
                                if (offset == 1)
                                {
                                    pole2 = worksheet.Cell(pozycja.Row - 1, pozycja.Col).GetFormattedString().Trim();
                                }
                                if (offset == 2)
                                {
                                    pole2 = pole1;
                                }
                                break;
                            }
                        }
                        if (!string.IsNullOrEmpty(pole2))
                        {
                            break;
                        }
                    }
                }
                catch
                {
                    return new Pracownik();
                }
                
            }
            if (!string.IsNullOrEmpty(pole1) && int.TryParse(pole1, out _))
            {
                pracownik.Akronim = pole1;
                if (!string.IsNullOrEmpty(pole2))
                {
                    string[] parts = pole2.Split(" ");
                    if (parts.Length == 2)
                    {
                        pracownik.Nazwisko = parts[0];
                        pracownik.Imie = parts[1];
                    }
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(pole2))
                {
                    string[] parts = pole2.Split(" ");
                    if (parts.Length == 2)
                    {
                        pracownik.Nazwisko = parts[0];
                        pracownik.Imie = parts[1];
                    }
                    else if (parts.Length == 3)
                    {
                        pracownik.Nazwisko = parts[0];
                        pracownik.Imie = parts[1];
                        if (int.TryParse(parts[2], out _))
                        {
                            pracownik.Akronim = parts[2];
                        }
                    }
                    else
                    {
                        Internal_Error_Logger.New_Error(pole1, "Imie nazwisko akronim", pozycja.Col, pozycja.Row, "Błędny format danych w komórkach imie nazwisko akronim");
                    }
                }
            }
            return pracownik;
        }
        private static List<Dane_Dnia> Get_Dane_Dni(IXLWorksheet worksheet, Helper.Current_Position pozycja)
        {
            List<Dane_Dnia> Dane_Dni = [];
            for (int i = 0; i < 31; i++)
            {
                string dane;
                string danedzien = worksheet.Cell(pozycja.Row, 1).GetFormattedString().Trim(); ;
                if (string.IsNullOrEmpty(danedzien))
                {
                    break;
                }
                Dane_Dnia dane_Dnia = new()
                {
                    Nr_Dnia = i + 1
                };
                dane = worksheet.Cell(pozycja.Row, pozycja.Col).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    pozycja.Row += 1;
                    continue;
                }

                if(!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref dane_Dnia.Godzina_Pracy_Od))
                {
                    Internal_Error_Logger.New_Error(dane, "Godzina pracy od", pozycja.Col, pozycja.Row, "Błędna wartość w godzinie pracy od");
                }

                dane = worksheet.Cell(pozycja.Row, pozycja.Col + 1).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    pozycja.Row += 1;
                    continue;
                }
                if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref dane_Dnia.Godzina_Pracy_Do))
                {
                    Internal_Error_Logger.New_Error(dane, "Godzina pracy do", pozycja.Col + 1, pozycja.Row, "Błędna wartość w godzinie pracy do");
                }
                Dane_Dni.Add(dane_Dnia);
                pozycja.Row += 1;
                
            }
            return Dane_Dni;
        }
        private static async Task Dodaj_Plany_do_Optimy(List<Grafik> grafiki)
        {
            await DbManager.Transaction_Manager.Create_Transaction();
            int dodano = 0;
            try
            {
                foreach (Grafik grafik in grafiki)
                {
                    if (grafik.Dane_Dni.Count <= 0)
                    {
                        continue;
                    }
                    int liczbaDniWMiesiacu = DateTime.DaysInMonth(grafik.Rok, grafik.Miesiac);
                    for (int dzien = 1; dzien <= liczbaDniWMiesiacu; dzien++)
                    {
                        if (!DateTime.TryParse($"{grafik.Rok}-{grafik.Miesiac:D2}-{dzien:D2}", out DateTime Data_Karty))
                        {
                            continue;
                        }

                        var daneKarty = grafik.Dane_Dni.FirstOrDefault(d => d.Nr_Dnia == dzien);
                        if (daneKarty == null)
                        {
                            Zrob_Insert_Plan_command(grafik.Pracownik, Data_Karty, TimeSpan.Zero, TimeSpan.Zero);
                            continue;
                        }

                        Helper.Typ_Insert_Obecnosc typ = Helper.Get_Typ_Insert_Plan(daneKarty.Godzina_Pracy_Od, daneKarty.Godzina_Pracy_Do);
                        switch (typ)
                        {
                            case Helper.Typ_Insert_Obecnosc.Zerowka:
                                if (DateTime.TryParseExact($"{grafik.Rok}-{grafik.Miesiac:D2}-{daneKarty.Nr_Dnia:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                                {
                                    Zrob_Insert_Plan_command(grafik.Pracownik, result, TimeSpan.Zero, TimeSpan.Zero);
                                }
                                break;

                            case Helper.Typ_Insert_Obecnosc.Normalna:
                                if (DateTime.TryParseExact($"{grafik.Rok}-{grafik.Miesiac:D2}-{daneKarty.Nr_Dnia:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result2))
                                {
                                    dodano += Zrob_Insert_Plan_command(grafik.Pracownik, result2, daneKarty.Godzina_Pracy_Od, daneKarty.Godzina_Pracy_Do);
                                }
                                break;

                            case Helper.Typ_Insert_Obecnosc.Nieinsertuj:
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}", false);
                DbManager.Transaction_Manager.RollBack_Transaction();
                throw new Exception($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
            }

            DbManager.Transaction_Manager.Commit_Transaction();
            if (dodano > 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Poprawnie dodawno plan z pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}");
                Console.ForegroundColor = ConsoleColor.White;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Nie dodano żadnego planu z pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}");
                Console.ForegroundColor = ConsoleColor.White;
            }
        }
        private static int Zrob_Insert_Plan_command(Pracownik pracownik, DateTime data, TimeSpan startGodz, TimeSpan endGodz)
        {
            try
            {
                int IdPracownika = pracownik.Get_PraId();
                using (SqlCommand command = new(DbManager.Check_Duplicate_Plan_Pracy, DbManager.GetConnection(), DbManager.Transaction_Manager.CurrentTransaction))
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

                using (SqlCommand command = new(DbManager.Insert_Plan_Pracy, DbManager.GetConnection(), DbManager.Transaction_Manager.CurrentTransaction))
                {
                    command.Parameters.AddWithValue("@DataInsert", data);
                    command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = (DateTime)(DbManager.Base_Date + startGodz);
                    command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = (DateTime)(DbManager.Base_Date + endGodz);
                    command.Parameters.AddWithValue("@PRI_PraId", IdPracownika);
                    command.Parameters.Add("@Strefa", SqlDbType.Int).Value = (int)Helper.Strefa.undefined;
                    command.Parameters.AddWithValue("@ImieMod", Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20));
                    command.Parameters.AddWithValue("@NazwiskoMod", Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 50));
                    command.Parameters.AddWithValue("@DataMod", Internal_Error_Logger.Last_Mod_Time);
                    int affected = command.ExecuteNonQuery();
                    if (startGodz != TimeSpan.Zero && endGodz != TimeSpan.Zero)
                    {
                        return affected;
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return 0;
        }
    }
}