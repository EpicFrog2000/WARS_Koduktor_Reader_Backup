using System.Data;
using System.Globalization;
using System.Transactions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Data.SqlClient;
using static Konduktor_Reader.Reader_Tabela_Stawek_v1;

namespace Konduktor_Reader
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
                var months = new Dictionary<int, string>
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
                var parts = value.Split(' ');
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

                List<TimeSpan> newGodziny_Pracy_Od = new List<TimeSpan>();
                List<TimeSpan> newGodziny_Pracy_Do = new List<TimeSpan>();


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
        private class Absencja
        {
            public int Dzien = 0;
            public int Miesiac = 0;
            public int Rok = 0;
            public string Nazwa = string.Empty;
            public decimal Liczba_Godzin_Absencji = 8;
            public RodzajAbsencji Rodzaj_Absencji = 0;
        }
        public class Prowizje
        {
            public decimal Suma_Wartosc_Towarow = 0;
            public decimal Suma_Liczba_Napojow_Awaryjnych = 0;
            public int Miesiac = 0;
            public int Rok = 0;
            public Pracownik Pracownik = new();
        }
        private enum RodzajAbsencji
        {
            DE,     // Delegacja
            DM,     // Dodatkowy urlop macierzyński
            DR,     // Urlop rodzicielski
            IK,     // Izolacja - Koronawirus
            NB,     // Badania lekarskie - okresowe
            NN,     // Nieobecność nieusprawiedliwiona
            NR,     // Badania lekarskie - z tyt. niepełnosprawności
            NU,     // Nieobecność usprawiedliwiona
            OD,     // Oddelegowanie do prac w ZZ
            OG,     // Odbiór godzin dyżuru
            ON,     // Odbiór nadgodzin
            OO,     // Odbiór pracy w niedziele
            OP,     // Urlop opiekuńczy (niepłatny)
            OS,     // Odbiór pracujących sobót
            PP,     // Poszukiwanie pracy
            PZ,     // Praca zdalna okazjonalna
            SW,     // Urlop/zwolnienie z tyt. siły wyższej
            SZ,     // Szkolenie
            SP,     // Zwolniony z obowiązku świadcz. pracy
            U9,     // Urlop rodzicielski 9 tygodni
            UA,     // Długotrwały urlop bezpłatny
            UB,     // Urlop bezpłatny
            UC,     // Urlop ojcowski
            UD,     // Na opiekę nad dzieckiem art.K.P.188
            UJ,     // Ćwiczenia wojskowe
            UK,     // Urlop dla krwiodawcy
            UL,     // Służba wojskowa
            ULawnika, // Praca ławnika w sądzie
            UM,     // Urlop macierzyński
            UN,     // Urlop z tyt. niepełnosprawności
            UO,     // Urlop okolicznościowy
            UP,     // Dodatkowy urlop osoby represjonowanej
            UR,     // Dodatkowe dni na turnus rehabilitacyjny
            US,     // Urlop szkoleniowy
            UV,     // Urlop weterana
            UW,     // Urlop wypoczynkowy
            UY,     // Urlop wychowawczy
            UZ,     // Urlop na żądanie
            WY,     // Wypoczynek skazanego
            ZC,     // Opieka nad członkiem rodziny (ZLA)
            ZD,     // Opieka nad dzieckiem (ZUS ZLA)
            ZK,     // Opieka nad dzieckiem Koronawirus
            ZL,     // Zwolnienie lekarskie (ZUS ZLA)
            ZN,     // Zwolnienie lekarskie niepłatne (ZLA)
            ZP,     // Kwarantanna sanepid
            ZR,     // Zwolnienie na rehabilitację (ZUS ZLA)
            ZS,     // Zwolnienie szpitalne (ZUS ZLA)
            ZY,     // Zwolnienie powypadkowe (ZUS ZLA)
            ZZ      // Zwolnienie lek. (ciąża) (ZUS ZLA)
        }
        public static void Process_Zakladka(IXLWorksheet Zakladka)
        {
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
                Program.error_logger.New_Error(dane, "Naglowek", Pozycja.Col, Pozycja.Row - 3, "Program nie znalazł naglowka karty");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            Karta_Ewidencji.Set_Date(dane);
            if (Karta_Ewidencji.Miesiac < 1 || Karta_Ewidencji.Rok == 0)
            {
                Program.error_logger.New_Error(dane, "Naglowek", Pozycja.Col, Pozycja.Row - 3, "Zły format naglowka, nie wczytano miesiaca lub roku");
                throw new Exception(Program.error_logger.Get_Error_String());
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
                Program.error_logger.New_Error(dane, "Imie Nazwisko Akronim", Pozycja.Col + 22, Pozycja.Row - 2, "Program nie znalazł imienia nazwiska i akronimu karty");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
        }
        private static void Get_Dane_Miesiaca(ref Karta_Ewidencji Karta_Ewidencji, IXLWorksheet Zakladka, Helper.Current_Position Pozycja)
        {
            string dzien = "";
            string dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "Dzien", Pozycja.Col, Pozycja.Row, "Program nie znalazł danych dnia miesiaca karty");
                throw new Exception(Program.error_logger.Get_Error_String());
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
                    Program.error_logger.New_Error(dane, "Numer relacji", Pozycja.Col + 1, Pozycja.Row + Row_Offset, $"Proszę wpisać poprawny nr relacji oraz jej opis zamiast {dane}");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                Dane_Karty Dane_Karty = new();
                Dane_Karty.Relacja.Numer_Relacji = dane;
                dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<string>(dane, ref Dane_Karty.Relacja.Opis_Relacji_1))
                {
                    Program.error_logger.New_Error(dane, "Opis Relacji", Pozycja.Col + 2, Pozycja.Row + Row_Offset, "Program nie znalazł opisu do relacji");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                do // exit na brak nr dnia
                {
                    // todo nr dnia
                    Dane_Dnia Dane_Dnia = new();
                    dzien = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (!Helper.Try_Get_Type_From_String<int>(dzien, ref Dane_Dnia.Dzien))
                    {
                        Program.error_logger.New_Error(dzien, "Dzien", Pozycja.Col, Pozycja.Row + Row_Offset, "Zły format Dnia miesiaca");
                        throw new Exception(Program.error_logger.Get_Error_String());
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
                                Program.error_logger.New_Error(part, "Godziny Pracy Od", Pozycja.Col + 4, Pozycja.Row + Row_Offset, "Zły format Godziny Pracy Od");
                                throw new Exception(Program.error_logger.Get_Error_String());
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
                                Program.error_logger.New_Error(part, "Godziny Pracy Do", Pozycja.Col + 5, Pozycja.Row + Row_Offset, "Zły format Godziny Pracy Do");
                                throw new Exception(Program.error_logger.Get_Error_String());
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
                                Program.error_logger.New_Error(part, "Godziny Odpoczynku Od", Pozycja.Col + 6, Pozycja.Row + Row_Offset, "Zły format Godziny Odpoczynku Od");
                                throw new Exception(Program.error_logger.Get_Error_String());
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
                                Program.error_logger.New_Error(part, "Godziny Odpoczynku Do", Pozycja.Col + 7, Pozycja.Row + Row_Offset, "Zły format Godziny Odpoczynku Do");
                                throw new Exception(Program.error_logger.Get_Error_String());
                            }
                        }
                    }
                    //nadg 50
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 13).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_50))
                        {
                            Program.error_logger.New_Error(dane, "Liczba Godzin Nadliczbowych 50", Pozycja.Col + 13, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych 50");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                    }
                    //nadg 100
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 14).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_100))
                        {
                            Program.error_logger.New_Error(dane, "Liczba Godzin Nadliczbowych 100", Pozycja.Col + 14, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych 100");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                    }
                    //nadg rycz 50
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 15).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50))
                        {
                            Program.error_logger.New_Error(dane, "Liczba Godzin Nadliczbowych W Ryczalcie 50", Pozycja.Col + 15, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych W Ryczalcie 50");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                    }
                    //nadg rycz 100
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 16).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100))
                        {
                            Program.error_logger.New_Error(dane, "Liczba Godzin Nadliczbowych W Ryczalcie 100", Pozycja.Col + 16, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych W Ryczalcie 100");
                            throw new Exception(Program.error_logger.Get_Error_String());
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
                    Program.error_logger.New_Error(dzien, "Dzien Absencji", Pozycja.Col, Pozycja.Row + Row_Offset, "Zły format Dnia absencji");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                if (!Helper.Try_Get_Type_From_String<string>(dane, ref Absencja.Nazwa))
                {
                    Program.error_logger.New_Error(dane, "Nazwa Absencji", Pozycja.Col + 17, Pozycja.Row + Row_Offset, "Zły format Nazwy absencji");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                if(!RodzajAbsencji.TryParse(Absencja.Nazwa, out Absencja.Rodzaj_Absencji)){
                    Program.error_logger.New_Error(dane, "Rodzaj Absencji", Pozycja.Col + 17, Pozycja.Row + Row_Offset, "Nierozpoznany rodzaj absencji");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 18).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Absencja.Liczba_Godzin_Absencji))
                {
                    //Program.error_logger.New_Error(dane, "Liczba godz Absencji", Pozycja.Col + 18, Pozycja.Row + Row_Offset, "Zły format lub bark Liczba godz Absencji");
                    //throw new Exception(Program.error_logger.Get_Error_String());
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
            foreach (var daneKarty in Karta_Ewidencji.Dane_Karty)
            {
                foreach (var daneDnia in daneKarty.Dane_Dni_Relacji)
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
                            Dane_Dnia.Podziel_Nadgodziny();
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
                                var baseTime = TimeSpan.FromHours(8);
                                if (godzNadlPlatne50 > 0)
                                {
                                    Dane_Dnia.Godziny_Pracy_Od.Add(baseTime);
                                    Dane_Dnia.Godziny_Pracy_Do.Add(baseTime + TimeSpan.FromHours((double)godzNadlPlatne50));
                                }
                                if (godzNadlPlatne100 > 0)
                                {
                                    if (Dane_Dnia.Godziny_Pracy_Od.Count == 0)
                                    {
                                        Dane_Dnia.Godziny_Pracy_Od.Add(baseTime);
                                    }
                                    else
                                    {
                                        Dane_Dnia.Godziny_Pracy_Od.Add(baseTime + TimeSpan.FromHours((double)godzNadlPlatne50));
                                    }
                                    Dane_Dnia.Godziny_Pracy_Do.Add(baseTime + TimeSpan.FromHours((double)godzNadlPlatne50) + TimeSpan.FromHours((double)godzNadlPlatne100));
                                }
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
                DateTime godzOdDate = Program.baseDate + startPodstawowy;
                DateTime godzDoDate = Program.baseDate + endPodstawowy;
                bool duplicate = false;
                int IdPracownika = Karta_Ewidencji.Pracownik.Get_PraId(connection, transaction);
                using (SqlCommand cmd = new(@"
        IF EXISTS (
            SELECT 1
            FROM cdn.PracPracaDni P
            INNER JOIN CDN.PracPracaDniGodz G ON P.PPR_PprId = G.PGR_PprId
            WHERE P.PPR_PraId = @PRI_PraId 
              AND P.PPR_Data = @DataInsert
              AND G.PGR_OdGodziny = @GodzOdDate
              AND G.PGR_DoGodziny = @GodzDoDate
              AND G.PGR_Strefa = @TypPracy
        )
        BEGIN
            SELECT 1;
        END
        ELSE
        BEGIN
            SELECT 0;
        END", connection, transaction))
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
                    string sqlQueryInsertObecnościDoOptimy = @"
DECLARE @EXISTSDZIEN DATETIME = (SELECT PracPracaDni.PPR_Data FROM cdn.PracPracaDni WHERE PPR_PraId = @PRI_PraId and PPR_Data = @DataInsert)
IF @EXISTSDZIEN is null
BEGIN
    BEGIN TRY
        INSERT INTO [CDN].[PracPracaDni]
                    ([PPR_PraId]
                    ,[PPR_Data]
                    ,[PPR_TS_Zal]
                    ,[PPR_TS_Mod]
                    ,[PPR_OpeModKod]
                    ,[PPR_OpeModNazwisko]
                    ,[PPR_OpeZalKod]
                    ,[PPR_OpeZalNazwisko]
                    ,[PPR_Zrodlo]
                    ,[PPR_Relacja])
                VALUES
                    (@PRI_PraId
                    ,@DataInsert
                    ,@DataMod
                    ,@DataMod
                    ,@ImieMod
                    ,@NazwiskoMod
                    ,@ImieMod
                    ,@NazwiskoMod
                    ,0
                    ,@Numer_Relacji)
    END TRY
    BEGIN CATCH
    END CATCH
END

INSERT INTO CDN.PracPracaDniGodz
		(PGR_PprId,
		PGR_Lp,
		PGR_OdGodziny,
		PGR_DoGodziny,
		PGR_Strefa,
		PGR_DzlId,
		PGR_PrjId,
		PGR_Uwagi,
		PGR_OdbNadg)
	VALUES
		((select PPR_PprId from cdn.PracPracaDni where CAST(PPR_Data as datetime) = @DataInsert and PPR_PraId = @PRI_PraId),
		1,
		@GodzOdDate,
		@GodzDoDate,
		@TypPracy,
		1,
		1,
		'',
		1);";
                    using (SqlCommand insertCmd = new(sqlQueryInsertObecnościDoOptimy, connection, transaction))
                    {
                        insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                        insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                        insertCmd.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = Data_Karty;
                        insertCmd.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                        insertCmd.Parameters.Add("@TypPracy", SqlDbType.Int).Value = Typ_Pracy;
                        insertCmd.Parameters.Add("@ImieMod", SqlDbType.NVarChar, 20).Value = Helper.Truncate(Program.error_logger.Last_Mod_Osoba, 20);
                        insertCmd.Parameters.Add("@NazwiskoMod", SqlDbType.NVarChar, 50).Value = Helper.Truncate(Program.error_logger.Last_Mod_Osoba, 50);
                        insertCmd.Parameters.Add("@DataMod", SqlDbType.DateTime).Value = Program.error_logger.Last_Mod_Time;
                        insertCmd.Parameters.Add("@Numer_Relacji", SqlDbType.NVarChar, 20).Value = Numer_Relacji;
                        insertCmd.ExecuteScalar();
                    }
                    return 1;
                }
            }
            catch (SqlException ex)
            {
                Program.error_logger.New_Custom_Error("Error podczas operacji w bazie(Zrob_Insert_Obecnosc_Command): " + ex.Message);
                transaction.Rollback();
                throw;
            }
            catch (Exception ex)
            {
                Program.error_logger.New_Custom_Error("Error: " + ex.Message);
                transaction.Rollback();
                throw;
            }
            return 0;
        }
        private static void Dodaj_Dane_Do_Optimy(Karta_Ewidencji Karta_Ewidencji, List<Prowizje> Prowizje)
        {
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                connection.Open();
                using (SqlTransaction tran = connection.BeginTransaction())
                {
                    if(Dodaj_Obecnosci_do_Optimy(Karta_Ewidencji, tran, connection) > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Poprawnie dodano obecnosci z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Nie dodano żadnych obesnosci");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    if (Dodaj_Absencje_do_Optimy(Karta_Ewidencji.Absencje, tran, connection, Karta_Ewidencji.Pracownik) > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Poprawnie dodano absencje z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
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
                        Console.WriteLine($"Poprawnie dodano prowizje z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
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
        private static int Ile_Dni_Roboczych(List<Absencja> Absencje)
        {
            return Absencje.Count(Absencja =>
            {
                DateTime absenceDate = new(Absencja.Rok, Absencja.Miesiac, Absencja.Dzien);
                return absenceDate.DayOfWeek != DayOfWeek.Saturday && absenceDate.DayOfWeek != DayOfWeek.Sunday;
            });
        }
        private static List<List<Absencja>> Podziel_Absencje_Na_Osobne(List<Absencja> Absencje)
        {
            List<List<Absencja>> OsobneAbsencje = new();
            List<Absencja> currentGroup = new();

            foreach (Absencja Absencja in Absencje)
            {
                if (currentGroup.Count == 0 || Absencja.Dzien == currentGroup[^1].Dzien + 1)
                {
                    currentGroup.Add(Absencja);
                }
                else
                {
                    OsobneAbsencje.Add(new List<Absencja>(currentGroup));
                    currentGroup = new List<Absencja> { Absencja };
                }
            }

            if (currentGroup.Count > 0)
            {
                OsobneAbsencje.Add(currentGroup);
            }

            return OsobneAbsencje;
        }
        private static string Dopasuj_Nieobecnosc(RodzajAbsencji rodzaj)
        {
            return rodzaj switch
            {
                RodzajAbsencji.UO => "Urlop okolicznościowy",
                RodzajAbsencji.ZL => "Zwolnienie chorobowe/F",
                RodzajAbsencji.ZY => "Zwolnienie chorobowe/wyp.w drodze/F",
                RodzajAbsencji.ZS => "Zwolnienie chorobowe/wyp.przy pracy/F",
                RodzajAbsencji.ZN => "Zwolnienie chorobowe/bez prawa do zas.",
                RodzajAbsencji.ZP => "Zwolnienie chorobowe/pozbawiony prawa",
                RodzajAbsencji.UR => "Urlop rehabilitacyjny",
                RodzajAbsencji.ZR => "Urlop rehabilitacyjny/wypadek w drodze..",
                RodzajAbsencji.ZD => "Urlop rehabilitacyjny/wypadek przy pracy",
                RodzajAbsencji.UM => "Urlop macierzyński",
                RodzajAbsencji.UC => "Urlop ojcowski",
                RodzajAbsencji.OP => "Urlop opiekuńczy (zasiłek)",
                RodzajAbsencji.UY => "Urlop wychowawczy (121)",
                RodzajAbsencji.UW => "Urlop wypoczynkowy",
                RodzajAbsencji.NU => "Nieobecność usprawiedliwiona (151)",
                RodzajAbsencji.NN => "Nieobecność nieusprawiedliwiona (152)",
                RodzajAbsencji.UL => "Służba wojskowa",
                RodzajAbsencji.DR => "Urlop rodzicielski",
                RodzajAbsencji.DM => "Urlop macierzyński dodatkowy",
                RodzajAbsencji.PP => "Dni wolne na poszukiwanie pracy",
                RodzajAbsencji.UK => "Dni wolne z tyt. krwiodawstwa",
                RodzajAbsencji.IK => "Covid19",
                _ => "Nieobecność (B2B)"
            };
        }
        private static int Dopasuj_Przyczyne(RodzajAbsencji rodzaj)
        {
            return rodzaj switch
            {
                RodzajAbsencji.ZL => 1,        // Zwolnienie lekarskie
                RodzajAbsencji.DM => 2,        // Urlop macierzyński
                RodzajAbsencji.DR => 13,        // Urlop opiekuńczy
                RodzajAbsencji.NB => 1,        // Zwolnienie lekarskie
                RodzajAbsencji.NN => 5,        // Nieobecność nieusprawiedliwiona
                RodzajAbsencji.UC => 21,       // Urlop opiekuńczy
                RodzajAbsencji.UD => 21,       // Urlop opiekuńczy
                RodzajAbsencji.UJ => 10,       // Służba wojskowa
                RodzajAbsencji.UL => 10,       // Służba wojskowa
                RodzajAbsencji.UM => 2,       // Urlop macierzyński
                RodzajAbsencji.UO => 4,       // Urlop okolicznościowy
                RodzajAbsencji.UN => 3,       // Urlop rehabilitacyjny
                RodzajAbsencji.UR => 3,       // Urlop rehabilitacyjny
                RodzajAbsencji.ZC => 21,       // Urlop opiekuńczy
                RodzajAbsencji.ZD => 21,       // Urlop opiekuńczy
                RodzajAbsencji.ZK => 21,       // Urlop opiekuńczy
                RodzajAbsencji.ZN => 1,       // Zwolnienie lekarskie
                RodzajAbsencji.ZR => 3,       // Urlop rehabilitacyjny
                RodzajAbsencji.ZZ => 1,       // Zwolnienie lekarskie
                _ => 9                             // Nie dotyczy dla pozostałych przypadków
            };
        }
        private static int Dodaj_Absencje_do_Optimy(List<Absencja> Absencje, SqlTransaction tran, SqlConnection connection, Pracownik Pracownik)
        {
            int ilosc_wpisow = 0;
            List<List<Absencja>> ListyAbsencji = Podziel_Absencje_Na_Osobne(Absencje);
            foreach (var ListaAbsencji in ListyAbsencji)
            {
                DateTime Data_Absencji_Start;
                DateTime Data_Absencji_End;

                try
                {
                    Data_Absencji_Start = new DateTime(ListaAbsencji[0].Rok, ListaAbsencji[0].Miesiac, ListaAbsencji[0].Dzien);
                    Data_Absencji_End = new DateTime(ListaAbsencji[ListaAbsencji.Count - 1].Rok, ListaAbsencji[ListaAbsencji.Count - 1].Miesiac, ListaAbsencji[ListaAbsencji.Count - 1].Dzien);
                }
                catch
                {
                    continue;
                }

                int przyczyna = Dopasuj_Przyczyne(ListaAbsencji[0].Rodzaj_Absencji);
                string nazwa_nieobecnosci = Dopasuj_Nieobecnosc(ListaAbsencji[0].Rodzaj_Absencji);

                if (string.IsNullOrEmpty(nazwa_nieobecnosci))
                {
                    Program.error_logger.New_Custom_Error($"W programie brak dopasowanego kodu nieobecnosci: {ListaAbsencji[0].Rodzaj_Absencji} w dniu {new DateTime(ListaAbsencji[0].Rok, ListaAbsencji[0].Miesiac, ListaAbsencji[0].Dzien)} z pliku: {Program.error_logger.Nazwa_Pliku} z zakladki: {Program.error_logger.Nr_Zakladki}. Absencja nie dodana.");
                    var e = new Exception();
                    e.Data["Kod"] = 42069;
                    throw e;
                }
                var dni_robocze = Ile_Dni_Roboczych(ListaAbsencji);
                var dni_calosc = ListaAbsencji.Count;

                bool duplicate = false;

                using (SqlCommand cmd = new SqlCommand(@"IF EXISTS (
SELECT 1 
FROM CDN.PracNieobec
WHERE [PNB_PraId] = @PRI_PraId
    AND [PNB_TnbId] = (
        SELECT TNB_TnbId 
        FROM cdn.TypNieobec 
        WHERE TNB_Nazwa = @NazwaNieobecnosci
    )
    AND [PNB_OkresOd] = @DataOd
    AND [PNB_OkresDo] = @DataDo
    AND [PNB_RozliczData] = @BaseDate
    AND [PNB_Przyczyna] = @Przyczyna
    AND [PNB_DniPracy] = @DniPracy
    AND [PNB_DniKalend] = @DniKalendarzowe
)
BEGIN
SELECT 1
END
ELSE 
BEGIN
SELECT 0
END
"
                , connection, tran))
                {

                    cmd.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = Pracownik.Get_PraId(connection, tran);
                    cmd.Parameters.Add("@NazwaNieobecnosci", SqlDbType.NVarChar, 50).Value = nazwa_nieobecnosci;
                    cmd.Parameters.Add("@DniPracy", SqlDbType.Int).Value = dni_robocze;
                    cmd.Parameters.Add("@DniKalendarzowe", SqlDbType.Int).Value = dni_calosc;
                    cmd.Parameters.Add("@Przyczyna", SqlDbType.NVarChar, 50).Value = przyczyna;
                    cmd.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = Data_Absencji_Start;
                    cmd.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = Program.baseDate;
                    cmd.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = Data_Absencji_End;
                    if ((int)cmd.ExecuteScalar() == 1)
                    {
                        duplicate = true;
                        return 0;
                    }
                }

                if (!duplicate)
                {
                    try
                    {
                        const string sqlQueryInsertNieObecnoŚciDoOptimy = @$"
DECLARE @TNBID INT = (SELECT TNB_TnbId FROM cdn.TypNieobec WHERE TNB_Nazwa = @NazwaNieobecnosci);
    INSERT INTO [CDN].[PracNieobec]
               ([PNB_PraId]
               ,[PNB_TnbId]
               ,[PNB_TyuId]
               ,[PNB_NaPodstPoprzNB]
               ,[PNB_OkresOd]
               ,[PNB_Seria]
               ,[PNB_Numer]
               ,[PNB_OkresDo]
               ,[PNB_Opis]
               ,[PNB_Rozliczona]
               ,[PNB_RozliczData]
               ,[PNB_ZwolFPFGSP]
               ,[PNB_UrlopNaZadanie]
               ,[PNB_Przyczyna]
               ,[PNB_DniPracy]
               ,[PNB_DniKalend]
               ,[PNB_Calodzienna]
               ,[PNB_ZlecZasilekPIT]
               ,[PNB_PracaRodzic]
               ,[PNB_Dziecko]
               ,[PNB_OpeZalId]
               ,[PNB_StaZalId]
               ,[PNB_TS_Zal]
               ,[PNB_TS_Mod]
               ,[PNB_OpeModKod]
               ,[PNB_OpeModNazwisko]
               ,[PNB_OpeZalKod]
               ,[PNB_OpeZalNazwisko]
               ,[PNB_Zrodlo])
         VALUES
               (@PRI_PraId
               ,@TNBID
               ,99999
               ,0
               ,@DataOd
               ,''
               ,''
               ,@DataDo
               ,''
               ,0
               ,@BaseDate
               ,0
               ,0
               ,@Przyczyna
               ,@DniPracy
               ,@DniKalendarzowe
               ,1
               ,0
               ,0
               ,''
               ,1
               ,1
               ,@DataMod
               ,@DataMod
               ,@ImieMod
               ,@NazwiskoMod
               ,@ImieMod
               ,@NazwiskoMod
               ,0);";
                        using (SqlCommand insertCmd = new SqlCommand(sqlQueryInsertNieObecnoŚciDoOptimy, connection, tran))
                        {
                            insertCmd.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = Pracownik.Get_PraId(connection, tran);
                            insertCmd.Parameters.Add("@NazwaNieobecnosci", SqlDbType.NVarChar, 50).Value = nazwa_nieobecnosci;
                            insertCmd.Parameters.Add("@DniPracy", SqlDbType.Int).Value = dni_robocze;
                            insertCmd.Parameters.Add("@DniKalendarzowe", SqlDbType.Int).Value = dni_calosc;
                            insertCmd.Parameters.Add("@Przyczyna", SqlDbType.NVarChar, 50).Value = przyczyna;
                            insertCmd.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = Data_Absencji_Start;
                            insertCmd.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = Program.baseDate;
                            insertCmd.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = Data_Absencji_End;
                            insertCmd.Parameters.Add("@ImieMod", SqlDbType.NVarChar, 20).Value = Helper.Truncate(Program.error_logger.Last_Mod_Osoba, 20);
                            insertCmd.Parameters.Add("@NazwiskoMod", SqlDbType.NVarChar, 50).Value = Helper.Truncate(Program.error_logger.Last_Mod_Osoba, 20);
                            insertCmd.Parameters.Add("@DataMod", SqlDbType.DateTime).Value = Program.error_logger.Last_Mod_Time;
                            insertCmd.ExecuteScalar();
                        }
                    }

                    catch (FormatException ex)
                    {
                        Program.error_logger.New_Custom_Error($"{ex.Message}");

                        continue;
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                    ilosc_wpisow++;
                }
            }
            return ilosc_wpisow;
        }
        private static Prowizje Get_Prowizje(Helper.Current_Position pozycja, IXLWorksheet Zakladka)
        {
            int offset = 0;
            string Dzien = string.Empty;
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
                        Program.error_logger.New_Error(dane, "Wartość Towarów", pozycja.Col + 23, pozycja.Row + offset);
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                    suma_wart_towarow += parsed_wartosc;
                }
                dane = Zakladka.Cell(pozycja.Row + offset, pozycja.Col + 24).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    decimal parsed_wartosc = 0;
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref parsed_wartosc))
                    {
                        Program.error_logger.New_Error(dane, "Liczba Napojow Awaryjnych", pozycja.Col + 24, pozycja.Row + offset);
                        throw new Exception(Program.error_logger.Get_Error_String());
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
            string query = @"WITH CTE AS (
            SELECT OAT_OatId 
            FROM cdn.OAtrybuty 
            WHERE OAT_AtkId = (SELECT ATK_AtkId FROM cdn.OAtrybutyKlasy WHERE ATK_Nazwa = @NazwaAtrybutu) and
			OAT_PrcId = @PracID
        )

        MERGE cdn.OAtrybutyHist AS target
        USING CTE AS source
        ON target.ATH_OatId = source.OAT_OatId
           AND target.ATH_DataOd = @ATHDataOd
           AND target.ATH_DataDo = @ATHDataDo
        WHEN MATCHED THEN
            UPDATE SET ATH_Wartosc = @NowaWartosc
        WHEN NOT MATCHED THEN
            INSERT (ATH_PrcId, ATH_AtkId, ATH_OatId, ATH_Wartosc, ATH_DataOd, ATH_DataDo)
            VALUES (0, 4, source.OAT_OatId, @NowaWartosc, @ATHDataOd, @ATHDataDo);";
            try
            {
                foreach (Prowizje Prowizja in Prowizje)
                {
                    DateTime Data_Od = DateTime.ParseExact($"{Prowizja.Rok}.{Prowizja.Miesiac}.01 00:00:00", "yyyy.MM.dd HH:mm:ss", CultureInfo.InvariantCulture);
                    DateTime Data_Do = Data_Od.AddMonths(1).AddDays(-1);
                    int pracId = Prowizja.Pracownik.Get_PraId(connection, transaction);
                    if(Prowizja.Suma_Liczba_Napojow_Awaryjnych > 0)
                    {
                        using (SqlCommand command = new(query, connection, transaction))
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
                        using (SqlCommand command = new(query, connection, transaction))
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
                Program.error_logger.New_Custom_Error("Error podczas operacji w bazie(Insert_Prowizje): " + ex.Message);
                transaction.Rollback();
                throw;
            }
            catch (Exception ex)
            {
                Program.error_logger.New_Custom_Error("Error: " + ex.Message);
                transaction.Rollback();
                throw;
            }
            return count;
        }
    }
}
