using System.Data;
using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace Excel_Data_Importer_WARS
{
    internal static class Reader_Karta_Ewidencji_Konduktora_v1
    {
        // JA JEBE KURWA PRZECIEZ TO BĘDZIE ZMIENIANE Z 500 GORYLIONÓW RAZY WSZYSTKO PORA SIE ZAJEBAĆ
        private class Karta_Ewidencji
        {
            public int Rok = 0;
            public int Miesiac = 0;
            public Pracownik Pracownik = new();
            public List<Dane_Karty> Dane_Karty = [];
            public List<Absencja> Absencje = [];
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
            public void Set_Rok(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return;
                }
                value = value.ToLower().Trim();
                string[] parts = value.Split(' ');
                if (!Helper.Try_Get_Type_From_String<int>(parts[^1], ref Rok))
                {
                    return;
                }
            }
            public void Set_Date(string value)
            {
                Set_Miesiac(value);
                Set_Rok(value);
            }
        }
        private class Dane_Karty
        {
            public Relacja Relacja = new();
            public List<Dane_Dnia> Dane_Dni_Relacji = [];
        }
        private class Dane_Dnia
        {
            public int Dzien = 0;
            public List<TimeSpan> Godziny_Pracy_Od = [];
            public List<TimeSpan> Godziny_Pracy_Do = [];
            public List<TimeSpan> Godziny_Odpoczynku_Od = [];
            public List<TimeSpan> Godziny_Odpoczynku_Do = [];
            public decimal Liczba_Godzin_Nadliczbowych_50 = 0;
            public decimal Liczba_Godzin_Nadliczbowych_100 = 0;
            public decimal Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 = 0;
            public decimal Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100 = 0;
            public string Absencja_Nazwa = string.Empty;
            public decimal Liczba_Godzin_Absencji = 0;
            public void Podziel_Nadgodziny()
            {
                if (Godziny_Pracy_Do.Count == 0) return;


                TimeSpan shiftEnd = Godziny_Pracy_Do[^1];

                TimeSpan overtime50 = Liczba_Godzin_Nadliczbowych_50 > 0
                    ? TimeSpan.FromHours((double)Liczba_Godzin_Nadliczbowych_50)
                    : TimeSpan.FromHours((double)Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50);

                TimeSpan overtime100 = Liczba_Godzin_Nadliczbowych_100 > 0
                    ? TimeSpan.FromHours((double)Liczba_Godzin_Nadliczbowych_100)
                    : TimeSpan.FromHours((double)Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100);

                TimeSpan overtimeStart = shiftEnd - (overtime50 + overtime100);
                if (overtimeStart < TimeSpan.Zero)
                {
                    overtimeStart = TimeSpan.FromHours(24) + overtimeStart;
                }

                List<TimeSpan> newGodziny_Pracy_Od = [];
                List<TimeSpan> newGodziny_Pracy_Do = [];


                for (int i = 0; i < Godziny_Pracy_Od.Count; i++)
                {
                    if (Godziny_Pracy_Do[i] == TimeSpan.Zero)
                    {
                        if (overtimeStart > Godziny_Pracy_Od[i] && overtimeStart > Godziny_Pracy_Do[i])
                        {

                            if (Godziny_Pracy_Od[i] != overtimeStart)
                            {
                                newGodziny_Pracy_Od.Add(Godziny_Pracy_Od[i]);
                                newGodziny_Pracy_Do.Add(overtimeStart);

                            }
                            if (overtimeStart + overtime50 != overtimeStart)
                            {
                                newGodziny_Pracy_Od.Add(overtimeStart);
                                newGodziny_Pracy_Do.Add(overtimeStart + overtime50);

                            }

                            newGodziny_Pracy_Od.Add(overtimeStart + overtime50);
                            newGodziny_Pracy_Do.Add(shiftEnd);
                            break;
                        }
                    }
                    else if (Godziny_Pracy_Do[i] > overtimeStart)
                    {
                        if (Godziny_Pracy_Do[i] < Godziny_Pracy_Od[i])
                        {
                            newGodziny_Pracy_Od.Add(Godziny_Pracy_Od[i]);
                            newGodziny_Pracy_Do.Add(TimeSpan.FromHours(24));
                            newGodziny_Pracy_Od.Add(TimeSpan.Zero);
                            newGodziny_Pracy_Do.Add(overtimeStart);
                        }
                        else
                        {
                            if (Godziny_Pracy_Od[i] != overtimeStart)
                            {
                                newGodziny_Pracy_Od.Add(Godziny_Pracy_Od[i]);
                                newGodziny_Pracy_Do.Add(overtimeStart);
                            }
                        }
                        if (overtimeStart + overtime50 != overtimeStart)
                        {
                            newGodziny_Pracy_Od.Add(overtimeStart);
                            newGodziny_Pracy_Do.Add(overtimeStart + overtime50);
                        }
                        newGodziny_Pracy_Od.Add(overtimeStart + overtime50);
                        newGodziny_Pracy_Do.Add(shiftEnd);
                        break;
                    }
                    newGodziny_Pracy_Od.Add(Godziny_Pracy_Od[i]);
                    newGodziny_Pracy_Do.Add(Godziny_Pracy_Do[i]);
                }
                Godziny_Pracy_Od = newGodziny_Pracy_Od;
                Godziny_Pracy_Do = newGodziny_Pracy_Do;
            }

        }
        public class Prowizje
        {
            public decimal Suma_Wartosc_Towarow = 0;
            public decimal Suma_Liczba_Napojow_Awaryjnych = 0;
            public int Miesiac = 0;
            public int Rok = 0;
            public Pracownik Pracownik = new();
        }

        private static Error_Logger Internal_Error_Logger = new(true);
        public static void Process_Zakladka(IXLWorksheet Zakladka, Error_Logger Error_Logger)
        {
            Internal_Error_Logger = Error_Logger;
            List<Karta_Ewidencji> Karty_Ewidencji = [];
            List<Helper.Current_Position> Pozycje = Helper.Find_Starting_Points(Zakladka, "Dzień miesiąca");
            List<Prowizje> Prowizje = [];
            foreach (Helper.Current_Position Pozycja in Pozycje)
            {
                Karta_Ewidencji Karta_Ewidencji = new();

                Get_Dane_Naglowka(ref Karta_Ewidencji, Zakladka, Pozycja);
                Pozycja.Row += 4;
                Get_Dane_Miesiaca(ref Karta_Ewidencji, Zakladka, Pozycja);
                Get_Absencje(ref Karta_Ewidencji, Zakladka, Pozycja);
                Karty_Ewidencji.Add(Karta_Ewidencji);
                Prowizje.Add(Get_Prowizje(Pozycja, Zakladka));
                Prowizje[^1].Pracownik = Karta_Ewidencji.Pracownik;
                Prowizje[^1].Rok = Karta_Ewidencji.Rok;
                Prowizje[^1].Miesiac = Karta_Ewidencji.Miesiac;
            }

            foreach (Karta_Ewidencji Karta_Ewidencji in Karty_Ewidencji)
            {
                Dodaj_Dane_Do_Optimy(Karta_Ewidencji, Prowizje);
            }
        }
        private static void Get_Dane_Naglowka(ref Karta_Ewidencji Karta_Ewidencji, IXLWorksheet Zakladka, Helper.Current_Position Pozycja)
        {
            string dane = Zakladka.Cell(Pozycja.Row - 3, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Internal_Error_Logger.New_Error(dane, "Naglowek", Pozycja.Col, Pozycja.Row - 3, "Program nie znalazł naglowka karty");
                throw new Exception(Internal_Error_Logger.Get_Error_String());
            }
            Karta_Ewidencji.Set_Date(dane);
            if (Karta_Ewidencji.Miesiac < 1 || Karta_Ewidencji.Rok == 0)
            {
                Internal_Error_Logger.New_Error(dane, "Naglowek", Pozycja.Col, Pozycja.Row - 3, "Zły format naglowka, nie wczytano miesiaca lub roku");
                throw new Exception(Internal_Error_Logger.Get_Error_String());
            }

            dane = Zakladka.Cell(Pozycja.Row - 2, Pozycja.Col + 22).GetFormattedString().Trim().Replace("  ", " ");
            if (!string.IsNullOrEmpty(dane))
            {
                string[] parts = dane.Trim().Split(' ');
                Karta_Ewidencji.Pracownik.Imie = parts[0].Trim();
                Karta_Ewidencji.Pracownik.Nazwisko = parts[1].Trim();

            }

            dane = Zakladka.Cell(Pozycja.Row - 1, Pozycja.Col + 22).GetFormattedString().Trim().Replace("  ", " ");
            if (!string.IsNullOrEmpty(dane))
            {
                Karta_Ewidencji.Pracownik.Akronim = dane;
            }

            if (string.IsNullOrEmpty(Karta_Ewidencji.Pracownik.Akronim) && string.IsNullOrEmpty(Karta_Ewidencji.Pracownik.Imie))
            {
                Internal_Error_Logger.New_Error(dane, "Imie Nazwisko Akronim", Pozycja.Col + 22, Pozycja.Row - 2, "Program nie znalazł imienia nazwiska i akronimu karty");
                throw new Exception(Internal_Error_Logger.Get_Error_String());
            }
        }
        private static void Get_Dane_Miesiaca(ref Karta_Ewidencji Karta_Ewidencji, IXLWorksheet Zakladka, Helper.Current_Position Pozycja)
        {
            string dzien;
            string dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Internal_Error_Logger.New_Error(dane, "Dzien", Pozycja.Col, Pozycja.Row, "Program nie znalazł danych dnia miesiaca karty");
                throw new Exception(Internal_Error_Logger.Get_Error_String());
            }
            int Row_Offset = 0;
            while (true) // skip puste pierwsze rzędy
            {
                dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    break;
                }
                Row_Offset++;
            }
            while (!string.IsNullOrEmpty(dane))
            {
                dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
                if(dane.Contains("Relacja z poprzedniego miesiąca"))
                {
                    Internal_Error_Logger.New_Error(dane, "Numer relacji", Pozycja.Col + 1, Pozycja.Row + Row_Offset, $"Proszę wpisać poprawny nr relacji oraz jej opis zamiast {dane}");
                    throw new Exception(Internal_Error_Logger.Get_Error_String());
                }
                Dane_Karty Dane_Karty = new();
                Dane_Karty.Relacja.Numer_Relacji = dane;
                dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<string>(dane, ref Dane_Karty.Relacja.Opis_Relacji_1))
                {
                    Internal_Error_Logger.New_Error(dane, "Opis Relacji", Pozycja.Col + 2, Pozycja.Row + Row_Offset, "Program nie znalazł opisu do relacji");
                    throw new Exception(Internal_Error_Logger.Get_Error_String());
                }

                do // exit na brak nr dnia
                {
                    // todo nr dnia
                    Dane_Dnia Dane_Dnia = new();
                    dzien = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (!Helper.Try_Get_Type_From_String<int>(dzien, ref Dane_Dnia.Dzien))
                    {
                        Internal_Error_Logger.New_Error(dzien, "Dzien", Pozycja.Col, Pozycja.Row + Row_Offset, "Zły format Dnia miesiaca");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }

                    //godz pracy od
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 4).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        string[] parts = dane.Trim().Split(' ');
                        foreach (string part in parts)
                        {
                            if (!Helper.Try_Get_Type_From_String<List<TimeSpan>>(part, ref Dane_Dnia.Godziny_Pracy_Od))
                            {
                                Internal_Error_Logger.New_Error(part, "Godziny Pracy Od", Pozycja.Col + 4, Pozycja.Row + Row_Offset, "Zły format Godziny Pracy Od");
                                throw new Exception(Internal_Error_Logger.Get_Error_String());
                            }
                        }
                    }

                    //godz pracy do
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 5).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        string[] parts = dane.Trim().Split(' ');
                        foreach (string part in parts)
                        {
                            if (!Helper.Try_Get_Type_From_String<List<TimeSpan>>(part, ref Dane_Dnia.Godziny_Pracy_Do))
                            {
                                Internal_Error_Logger.New_Error(part, "Godziny Pracy Do", Pozycja.Col + 5, Pozycja.Row + Row_Offset, "Zły format Godziny Pracy Do");
                                throw new Exception(Internal_Error_Logger.Get_Error_String());
                            }
                        }
                    }

                    //godz. odpoczynku od
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 6).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        string[] parts = dane.Trim().Split(' ');
                        foreach (string part in parts)
                        {
                            if (!Helper.Try_Get_Type_From_String<List<TimeSpan>>(part, ref Dane_Dnia.Godziny_Odpoczynku_Od))
                            {
                                Internal_Error_Logger.New_Error(part, "Godziny Odpoczynku Od", Pozycja.Col + 6, Pozycja.Row + Row_Offset, "Zły format Godziny Odpoczynku Od");
                                throw new Exception(Internal_Error_Logger.Get_Error_String());
                            }
                        }
                    }
                    //godz. odpoczynku do
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 7).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        string[] parts = dane.Trim().Split(' ');
                        foreach (string part in parts)
                        {
                            if (!Helper.Try_Get_Type_From_String<List<TimeSpan>>(part, ref Dane_Dnia.Godziny_Odpoczynku_Do))
                            {
                                Internal_Error_Logger.New_Error(part, "Godziny Odpoczynku Do", Pozycja.Col + 7, Pozycja.Row + Row_Offset, "Zły format Godziny Odpoczynku Do");
                                throw new Exception(Internal_Error_Logger.Get_Error_String());
                            }
                        }
                    }
                    //nadg 50
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 13).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_50))
                        {
                            Internal_Error_Logger.New_Error(dane, "Liczba Godzin Nadliczbowych 50", Pozycja.Col + 13, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych 50");
                            throw new Exception(Internal_Error_Logger.Get_Error_String());
                        }
                    }
                    //nadg 100
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 14).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_100))
                        {
                            Internal_Error_Logger.New_Error(dane, "Liczba Godzin Nadliczbowych 100", Pozycja.Col + 14, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych 100");
                            throw new Exception(Internal_Error_Logger.Get_Error_String());
                        }
                    }
                    //nadg rycz 50
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 15).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50))
                        {
                            Internal_Error_Logger.New_Error(dane, "Liczba Godzin Nadliczbowych W Ryczalcie 50", Pozycja.Col + 15, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych W Ryczalcie 50");
                            throw new Exception(Internal_Error_Logger.Get_Error_String());
                        }
                    }
                    //nadg rycz 100
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 16).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100))
                        {
                            Internal_Error_Logger.New_Error(dane, "Liczba Godzin Nadliczbowych W Ryczalcie 100", Pozycja.Col + 16, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych W Ryczalcie 100");
                            throw new Exception(Internal_Error_Logger.Get_Error_String());
                        }
                    }
                    Dane_Karty.Dane_Dni_Relacji.Add(Dane_Dnia);
                    Row_Offset++;
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
                    dzien = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (string.IsNullOrEmpty(dzien)) { break; }
                } while (string.IsNullOrEmpty(dane));

                Karta_Ewidencji.Dane_Karty.Add(Dane_Karty);
                dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");

            }
        }
        private static void Get_Absencje(ref Karta_Ewidencji Karta_Ewidencji, IXLWorksheet Zakladka, Helper.Current_Position Pozycja)
        {
            int Row_Offset = 0;
            do
            {
                string dzien = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");

                if (string.IsNullOrEmpty(dzien))
                {
                    return;
                }
                string dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 17).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(dane))
                {
                    Row_Offset++;
                    continue;
                }
                Absencja Absencja = new();
                if (!Helper.Try_Get_Type_From_String<int>(dzien, ref Absencja.Dzien))
                {
                    Internal_Error_Logger.New_Error(dzien, "Dzien Absencji", Pozycja.Col, Pozycja.Row + Row_Offset, "Zły format Dnia absencji");
                    throw new Exception(Internal_Error_Logger.Get_Error_String());
                }
                if (!Helper.Try_Get_Type_From_String<string>(dane, ref Absencja.Nazwa))
                {
                    Internal_Error_Logger.New_Error(dane, "Nazwa Absencji", Pozycja.Col + 17, Pozycja.Row + Row_Offset, "Zły format Nazwy absencji");
                    throw new Exception(Internal_Error_Logger.Get_Error_String());
                }
                if(!Absencja.RodzajAbsencji.TryParse(Absencja.Nazwa, out Absencja.Rodzaj_Absencji)){
                    Internal_Error_Logger.New_Error(dane, "Rodzaj Absencji", Pozycja.Col + 17, Pozycja.Row + Row_Offset, "Nierozpoznany rodzaj absencji");
                    throw new Exception(Internal_Error_Logger.Get_Error_String());
                }

                dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 18).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Absencja.Liczba_Godzin_Absencji))
                {
                    //Internal_Error_Logger.New_Error(dane, "Liczba godz Absencji", Pozycja.Col + 18, Pozycja.Row + Row_Offset, "Zły format lub bark Liczba godz Absencji");
                    //throw new Exception(Internal_Error_Logger.Get_Error_String());
                }
                Absencja.Rok = Karta_Ewidencji.Rok;
                Absencja.Miesiac = Karta_Ewidencji.Miesiac;
                
                Karta_Ewidencji.Absencje.Add(Absencja);
                Row_Offset++;
            } while (true);
        }
        private static int Dodaj_Obecnosci_do_Optimy(Karta_Ewidencji Karta_Ewidencji, SqlTransaction tran, SqlConnection connection)
        {
            
            HashSet<DateTime> Pasujace_Daty = [];
            foreach (Dane_Karty daneKarty in Karta_Ewidencji.Dane_Karty)
            {
                foreach (Dane_Dnia daneDnia in daneKarty.Dane_Dni_Relacji)
                {
                    Pasujace_Daty.Add(new DateTime(Karta_Ewidencji.Rok, Karta_Ewidencji.Miesiac, daneDnia.Dzien));
                }
            }
            DateTime startDate = new(Karta_Ewidencji.Rok, Karta_Ewidencji.Miesiac, 1);
            DateTime endDate = new(Karta_Ewidencji.Rok, Karta_Ewidencji.Miesiac, DateTime.DaysInMonth(Karta_Ewidencji.Rok, Karta_Ewidencji.Miesiac));
            for (DateTime dzien = startDate; dzien <= endDate; dzien = dzien.AddDays(1))
            {
                if (!Pasujace_Daty.Contains(dzien))
                {
                    Zrob_Insert_Obecnosc_Command(connection, tran, dzien, TimeSpan.Zero, TimeSpan.Zero, Karta_Ewidencji, 1, ""); // 1 - pusta strefa
                }
            }

            int ilosc_wpisow = 0;
            foreach (Dane_Karty Dane_Karty in Karta_Ewidencji.Dane_Karty)
            {
                foreach (Dane_Dnia Dane_Dnia in Dane_Karty.Dane_Dni_Relacji)
                {
                    if (DateTime.TryParse($"{Karta_Ewidencji.Rok}-{Karta_Ewidencji.Miesiac:D2}-{Dane_Dnia.Dzien:D2}", out DateTime Data_Karty))
                    {
                        if (Dane_Dnia.Godziny_Pracy_Od.Count >= 1)
                        {
                            //Dane_Dnia.Podziel_Nadgodziny();
                            for (int j = 0; j < Dane_Dnia.Godziny_Pracy_Od.Count; j++)
                            {
                                ilosc_wpisow += Zrob_Insert_Obecnosc_Command(connection, tran, Data_Karty, Dane_Dnia.Godziny_Pracy_Od[j], Dane_Dnia.Godziny_Pracy_Do[j], Karta_Ewidencji, 2, Dane_Karty.Relacja.Numer_Relacji);
                            }
                        }
                        else
                        {
                            decimal godzNadlPlatne50 = Dane_Dnia.Liczba_Godzin_Nadliczbowych_50 <= 0
                                ? (decimal)Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50
                                : (decimal)Dane_Dnia.Liczba_Godzin_Nadliczbowych_50;

                            decimal godzNadlPlatne100 = Dane_Dnia.Liczba_Godzin_Nadliczbowych_100 <= 0
                                ? (decimal)Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100
                                : (decimal)Dane_Dnia.Liczba_Godzin_Nadliczbowych_100;

                            if (godzNadlPlatne50 > 0 || godzNadlPlatne100 > 0)
                            {
                                TimeSpan baseTime = TimeSpan.FromHours(8);
                                Dane_Dnia.Godziny_Pracy_Od.Add(baseTime);
                                Dane_Dnia.Godziny_Pracy_Do.Add(baseTime + TimeSpan.FromHours((double)(godzNadlPlatne50 + godzNadlPlatne100)));
                                for (int k = 0; k < Dane_Dnia.Godziny_Pracy_Od.Count; k++)
                                {
                                    ilosc_wpisow += Zrob_Insert_Obecnosc_Command(connection, tran, Data_Karty, Dane_Dnia.Godziny_Pracy_Od[k], Dane_Dnia.Godziny_Pracy_Do[k], Karta_Ewidencji, 2, Dane_Karty.Relacja.Numer_Relacji);
                                }
                            }
                            else
                            {
                                Zrob_Insert_Obecnosc_Command(connection, tran, Data_Karty, TimeSpan.Zero, TimeSpan.Zero, Karta_Ewidencji, 1, ""); // 1 - pusta strefa
                            }
                        }
                    }
                }
            }
            return ilosc_wpisow;
        }
        private static int Zrob_Insert_Obecnosc_Command(SqlConnection connection, SqlTransaction transaction, DateTime Data_Karty, TimeSpan startPodstawowy, TimeSpan endPodstawowy, Karta_Ewidencji Karta_Ewidencji, int Typ_Pracy, string Numer_Relacji)
        {
            try
            {
                DateTime godzOdDate = DbManager.Base_Date + startPodstawowy;
                DateTime godzDoDate = DbManager.Base_Date + endPodstawowy;
                bool duplicate = false;
                int IdPracownika = -1;
                try
                {
                    IdPracownika = Karta_Ewidencji.Pracownik.Get_PraId(connection, transaction);
                }
                catch (Exception ex)
                {
                    connection.Close();
                    Internal_Error_Logger.New_Custom_Error(ex.Message + " z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}" + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                }

                using (SqlCommand cmd = new(DbManager.Check_Duplicate_Obecnosc, connection, transaction))
                {
                    cmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                    cmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                    cmd.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = Data_Karty;
                    cmd.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                    cmd.Parameters.Add("@TypPracy", SqlDbType.Int).Value = Typ_Pracy;
                    duplicate = (int)cmd.ExecuteScalar() == 1;
                }

                if (!duplicate)
                {
                    using (SqlCommand insertCmd = new(DbManager.Insert_Obecnosci, connection, transaction))
                    {
                        insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                        insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                        insertCmd.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = Data_Karty;
                        insertCmd.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                        insertCmd.Parameters.Add("@TypPracy", SqlDbType.Int).Value = Typ_Pracy;
                        insertCmd.Parameters.Add("@ImieMod", SqlDbType.NVarChar, 20).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20);
                        insertCmd.Parameters.Add("@NazwiskoMod", SqlDbType.NVarChar, 50).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 50);
                        insertCmd.Parameters.Add("@DataMod", SqlDbType.DateTime).Value = Internal_Error_Logger.Last_Mod_Time;
                        insertCmd.Parameters.Add("@Numer_Relacji", SqlDbType.NVarChar, 20).Value = Numer_Relacji;
                        insertCmd.ExecuteScalar();
                    }
                    return 1;
                }
            }
            catch (SqlException ex)
            {
                Internal_Error_Logger.New_Custom_Error("Error podczas operacji w bazie(Zrob_Insert_Obecnosc_Command): " + ex.Message);
                transaction.Rollback();
                throw;
            }
            catch (Exception ex)
            {
                Internal_Error_Logger.New_Custom_Error("Error: " + ex.Message);
                transaction.Rollback();
                throw;
            }
            return 0;
        }
        private static void Dodaj_Dane_Do_Optimy(Karta_Ewidencji Karta_Ewidencji, List<Prowizje> Prowizje)
        {
            using (SqlConnection connection = new(DbManager.Connection_String))
            {
                connection.Open();
                using (SqlTransaction tran = connection.BeginTransaction())
                {
                    if(Dodaj_Obecnosci_do_Optimy(Karta_Ewidencji, tran, connection) > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Poprawnie dodano obecnosci z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Nie dodano żadnych obesnosci");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    if (Absencja.Dodaj_Absencje_do_Optimy(Karta_Ewidencji.Absencje, tran, connection, Karta_Ewidencji.Pracownik, Internal_Error_Logger) > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Poprawnie dodano absencje z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Nie dodano żadnych absencji");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    if (Insert_Prowizje(Prowizje, tran, connection) > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Poprawnie dodano prowizje z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Nie dodano żadnych prowizji");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    tran.Commit();
                }
            }
        }
        private static Prowizje Get_Prowizje(Helper.Current_Position pozycja, IXLWorksheet Zakladka)
        {
            int offset = 0;
            string Dzien;
            decimal suma_wart_towarow = 0;
            decimal Liczba_Napojow_Awaryjnych = 0;
            Prowizje Prowizja = new();
            while (true)
            {
                Dzien = Zakladka.Cell(pozycja.Row + offset, pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(Dzien))
                {
                    break;
                }
                string dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 23).GetFormattedString().Trim().Replace("  ", " ");
                
                if (!string.IsNullOrEmpty(dane))
                {
                    decimal parsed_wartosc = 0;
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref parsed_wartosc))
                    {
                        Internal_Error_Logger.New_Error(dane, "Wartość Towarów", pozycja.Col + 23, pozycja.Row + offset);
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }
                    suma_wart_towarow += parsed_wartosc;
                }
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 24).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    decimal parsed_wartosc = 0;
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref parsed_wartosc))
                    {
                        Internal_Error_Logger.New_Error(dane, "Liczba Napojow Awaryjnych", pozycja.Col + 24, pozycja.Row + offset);
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }
                    Liczba_Napojow_Awaryjnych += parsed_wartosc;
                }
                offset++;
            }
            Prowizja.Suma_Wartosc_Towarow = suma_wart_towarow;
            Prowizja.Suma_Liczba_Napojow_Awaryjnych = Liczba_Napojow_Awaryjnych;
            return Prowizja;
        }
        private static int Insert_Prowizje(List<Prowizje> Prowizje, SqlTransaction transaction, SqlConnection connection)
        {
            int count = 0;
            try
            {
                foreach (Prowizje Prowizja in Prowizje)
                {
                    DateTime Data_Od = DateTime.ParseExact($"{Prowizja.Rok}.{Prowizja.Miesiac}.01 00:00:00", "yyyy.MM.dd HH:mm:ss", CultureInfo.InvariantCulture);
                    DateTime Data_Do = Data_Od.AddMonths(1).AddDays(-1);

                    int pracId = -1;
                    try
                    {
                        pracId = Prowizja.Pracownik.Get_PraId(connection, transaction);
                    }
                    catch (Exception ex)
                    {
                        connection.Close();
                        Internal_Error_Logger.New_Custom_Error(ex.Message + " z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                        throw new Exception(ex.Message + $" w pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}" + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                    }

                    if(Prowizja.Suma_Liczba_Napojow_Awaryjnych > 0)
                    {
                        using (SqlCommand command = new(DbManager.Insert_Prowizje, connection, transaction))
                        {
                            command.Parameters.Add("@PracID", SqlDbType.Int).Value = pracId;
                            command.Parameters.Add("@NowaWartosc", SqlDbType.Decimal).Value = Prowizja.Suma_Liczba_Napojow_Awaryjnych;
                            command.Parameters.Add("@NazwaAtrybutu", SqlDbType.NVarChar, 50).Value = "Prowizja za wydane napoje awaryjne";
                            command.Parameters.Add("@ATHDataOd", SqlDbType.DateTime).Value = Data_Od;
                            command.Parameters.Add("@ATHDataDo", SqlDbType.DateTime).Value = Data_Do;
                            count += command.ExecuteNonQuery();
                        }
                    }
                    if (Prowizja.Suma_Wartosc_Towarow > 0)
                    {
                        using (SqlCommand command = new(DbManager.Insert_Prowizje, connection, transaction))
                        {
                            command.Parameters.Add("@PracID", SqlDbType.Int).Value = pracId;
                            command.Parameters.Add("@NowaWartosc", SqlDbType.Decimal).Value = Prowizja.Suma_Wartosc_Towarow;
                            command.Parameters.Add("@NazwaAtrybutu", SqlDbType.NVarChar, 50).Value = "Prowizja za towar";
                            command.Parameters.Add("@ATHDataOd", SqlDbType.DateTime).Value = Data_Od;
                            command.Parameters.Add("@ATHDataDo", SqlDbType.DateTime).Value = Data_Do;
                            count += command.ExecuteNonQuery();

                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                Internal_Error_Logger.New_Custom_Error("Error podczas operacji w bazie(Insert_Prowizje): " + ex.Message);
                transaction.Rollback();
                throw;
            }
            catch (Exception ex)
            {
                Internal_Error_Logger.New_Custom_Error("Error: " + ex.Message);
                transaction.Rollback();
                throw;
            }
            return count;
        }
    }
}
