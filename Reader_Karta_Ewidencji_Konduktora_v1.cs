using System.Data;
using System.Diagnostics;
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

            // Ja jebe ten WARS to jest jebany syf 
            // Te godz ponizej zamiast tych Godz_PRacy_Od i Do
            // TODO: Dodać wczytytwanie z pliku
            // TODO: Dodać zapisywanie do bazy danych -> jest kod napisany tylko przetestować
            public List<TimeSpan> Godziny_Pracy_Obsluga_relacji_Od = [];
            public List<TimeSpan> Godziny_Pracy_Obsluga_relacji_Do = [];
            public List<TimeSpan> Godziny_Pracy_Inne_Czynnosci_Od = [];
            public List<TimeSpan> Godziny_Pracy_Inne_Czynnosci_Do = [];

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
        public static async Task Process_Zakladka(IXLWorksheet Zakladka, Error_Logger Error_Logger)
        {
            Internal_Error_Logger = Error_Logger;

            List<Karta_Ewidencji> Karty_Ewidencji = [];
            List<Helper.Current_Position> Pozycje = Helper.Find_Starting_Points(Zakladka, "Dzień miesiąca");
            List<Prowizje> Prowizje = [];

            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();

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

            Helper.Pomiar.Avg_Get_Dane_Z_Pliku = PomiaryStopWatch.Elapsed;

            PomiaryStopWatch.Restart();
            await Dodaj_Dane_Do_Optimy(Karty_Ewidencji, Prowizje);
            Helper.Pomiar.Avg_Dodawanie_Do_Bazy = PomiaryStopWatch.Elapsed;
        }
        private static void Get_Dane_Naglowka(ref Karta_Ewidencji Karta_Ewidencji, IXLWorksheet Zakladka, Helper.Current_Position Pozycja)
        {
            string dane = Zakladka.Cell(Pozycja.Row - 3, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Internal_Error_Logger.New_Error(dane, "Naglowek", Pozycja.Col, Pozycja.Row - 3, "Program nie znalazł naglowka karty");
            }
            Karta_Ewidencji.Set_Date(dane);
            if (Karta_Ewidencji.Miesiac < 1 || Karta_Ewidencji.Rok == 0)
            {
                Internal_Error_Logger.New_Error(dane, "Naglowek", Pozycja.Col, Pozycja.Row - 3, "Zły format naglowka, nie wczytano miesiaca lub roku");
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
            }
        }
        private static void Get_Dane_Miesiaca(ref Karta_Ewidencji Karta_Ewidencji, IXLWorksheet Zakladka, Helper.Current_Position Pozycja)
        {
            string dzien;
            string dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Internal_Error_Logger.New_Error(dane, "Dzien", Pozycja.Col, Pozycja.Row, "Program nie znalazł danych dnia miesiaca karty");
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
                // czasem jest zamiast opisu relacji to i to za dużo roboty to obslugiwanie tego, niech po prostu dobrze wpisują
                if (dane.Contains("Relacja z poprzedniego miesiąca"))
                {
                    Internal_Error_Logger.New_Error(dane, "Numer relacji", Pozycja.Col + 1, Pozycja.Row + Row_Offset, $"Proszę wpisać poprawny nr relacji oraz jej opis zamiast {dane}");
                }
                Dane_Karty Dane_Karty = new();
                Dane_Karty.Relacja.Numer_Relacji = dane;
                dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                if (!Helper.Try_Get_Type_From_String<string>(dane, ref Dane_Karty.Relacja.Opis_Relacji_1))
                {
                    Internal_Error_Logger.New_Error(dane, "Opis Relacji", Pozycja.Col + 2, Pozycja.Row + Row_Offset, "Program nie znalazł opisu do relacji");
                }

                do // exit na brak nr dnia
                {
                    // nr dnia
                    Dane_Dnia Dane_Dnia = new();
                    dzien = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (!Helper.Try_Get_Type_From_String<int>(dzien, ref Dane_Dnia.Dzien))
                    {
                        Internal_Error_Logger.New_Error(dzien, "Dzien", Pozycja.Col, Pozycja.Row + Row_Offset, "Zły format Dnia miesiaca");
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
                            }
                        }

                        //godz pracy do
                        dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 5).GetFormattedString().Trim().Replace("  ", " ");
                        if (!string.IsNullOrEmpty(dane))
                        {
                            string[] parts2 = dane.Trim().Split(' ');
                            foreach (string part in parts2)
                            {
                                if (!Helper.Try_Get_Type_From_String<List<TimeSpan>>(part, ref Dane_Dnia.Godziny_Pracy_Do))
                                {
                                    Internal_Error_Logger.New_Error(part, "Godziny Pracy Do", Pozycja.Col + 5, Pozycja.Row + Row_Offset, "Zły format Godziny Pracy Do");
                                }
                            }
                        }
                        else
                        {
                            Internal_Error_Logger.New_Error("", "Godziny Pracy Do", Pozycja.Col + 5, Pozycja.Row + Row_Offset, "Brak Godziny Pracy Do w tym dniu");
                        }
                        if (Dane_Dnia.Godziny_Pracy_Od.Count != Dane_Dnia.Godziny_Pracy_Do.Count)
                        {
                            Internal_Error_Logger.New_Error("", "Godziny Pracy", Pozycja.Col + 5, Pozycja.Row + Row_Offset, "Nie zgadza się liczba godzin pracy w tym dniu");
                        }
                    }

                    //TODO wczytywanie rzeczy ponirzej przesunć o 1 kolimne w prawo a powyrzej zamienić na inne czasy pracy



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
                            }
                        }

                        //godz. odpoczynku do
                        dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 7).GetFormattedString().Trim().Replace("  ", " ");
                        if (!string.IsNullOrEmpty(dane))
                        {
                            string[] parts2 = dane.Trim().Split(' ');
                            foreach (string part in parts2)
                            {
                                if (!Helper.Try_Get_Type_From_String<List<TimeSpan>>(part, ref Dane_Dnia.Godziny_Odpoczynku_Do))
                                {
                                    Internal_Error_Logger.New_Error(part, "Godziny Odpoczynku Do", Pozycja.Col + 7, Pozycja.Row + Row_Offset, "Zły format Godziny Odpoczynku Do");
                                }
                            }
                        }
                        else
                        {
                            Internal_Error_Logger.New_Error("", "Godziny Odpoczynku Do", Pozycja.Col + 7, Pozycja.Row + Row_Offset, "Brak Godziny Odpoczynku Do w tym dniu");
                        }

                        if (Dane_Dnia.Godziny_Odpoczynku_Od.Count != Dane_Dnia.Godziny_Odpoczynku_Do.Count)
                        {
                            Internal_Error_Logger.New_Error("", "Godziny odpoczynku", Pozycja.Col + 7, Pozycja.Row + Row_Offset, "Nie zgadza się liczba godzin odpoczynku w tym dniu");
                        }
                    }


                    //nadg 50
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 13).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_50))
                        {
                            Internal_Error_Logger.New_Error(dane, "Liczba Godzin Nadliczbowych 50", Pozycja.Col + 13, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych 50");
                        }
                    }
                    //nadg 100
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 14).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_100))
                        {
                            Internal_Error_Logger.New_Error(dane, "Liczba Godzin Nadliczbowych 100", Pozycja.Col + 14, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych 100");
                        }
                    }
                    //nadg rycz 50
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 15).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50))
                        {
                            Internal_Error_Logger.New_Error(dane, "Liczba Godzin Nadliczbowych W Ryczalcie 50", Pozycja.Col + 15, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych W Ryczalcie 50");
                        }
                    }
                    //nadg rycz 100
                    dane = Zakladka.Cell(Pozycja.Row + Row_Offset, Pozycja.Col + 16).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100))
                        {
                            Internal_Error_Logger.New_Error(dane, "Liczba Godzin Nadliczbowych W Ryczalcie 100", Pozycja.Col + 16, Pozycja.Row + Row_Offset, "Zły format Liczba Godzin Nadliczbowych W Ryczalcie 100");
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
                }
                if (!Helper.Try_Get_Type_From_String<string>(dane.ToUpper(), ref Absencja.Nazwa))
                {
                    Internal_Error_Logger.New_Error(dane, "Nazwa Absencji", Pozycja.Col + 17, Pozycja.Row + Row_Offset, "Zły format Nazwy absencji");
                }
                if(!Absencja.RodzajAbsencji.TryParse(Absencja.Nazwa, out Absencja.Rodzaj_Absencji)){
                    Internal_Error_Logger.New_Error(dane, "Rodzaj Absencji", Pozycja.Col + 17, Pozycja.Row + Row_Offset, "Nierozpoznany rodzaj absencji");
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
        private static int Dodaj_Obecnosci_do_Optimy(Karta_Ewidencji Karta_Ewidencji)
        {
            int ilosc_wpisow = 0;
            int liczbaDniWMiesiacu = DateTime.DaysInMonth(Karta_Ewidencji.Rok, Karta_Ewidencji.Miesiac);

            for (int dzien = 1; dzien <= liczbaDniWMiesiacu; dzien++)
            {
                if (!DateTime.TryParse($"{Karta_Ewidencji.Rok}-{Karta_Ewidencji.Miesiac:D2}-{dzien:D2}", out DateTime dataKarty))
                {
                    continue;
                }

                var dane = Karta_Ewidencji.Dane_Karty
                    .SelectMany(k => k.Dane_Dni_Relacji, (karta, dzien) => new { karta.Relacja.Numer_Relacji, Dane_Dnia = dzien })
                    .FirstOrDefault(d => d.Dane_Dnia.Dzien == dzien);

                if (dane == null)
                {
                    Zrob_Insert_Obecnosc_Command(dataKarty, TimeSpan.Zero, TimeSpan.Zero, Karta_Ewidencji, Helper.Strefa.undefined, "");
                    continue;
                }
                Helper.Typ_Insert_Obecnosc typ = Helper.Get_Typ_Insert_Obecnosc(dane.Dane_Dnia.Godziny_Pracy_Od, dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_50, dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_100, dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50, dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100);
                switch (typ)
                {
                    case Helper.Typ_Insert_Obecnosc.Zerowka:
                        Zrob_Insert_Obecnosc_Command(dataKarty, TimeSpan.Zero, TimeSpan.Zero, Karta_Ewidencji, Helper.Strefa.undefined, "");
                        break;

                    case Helper.Typ_Insert_Obecnosc.Normalna:
                        for (int j = 0; j < dane.Dane_Dnia.Godziny_Pracy_Od.Count; j++)
                        {
                            ilosc_wpisow += Zrob_Insert_Obecnosc_Command(dataKarty, dane.Dane_Dnia.Godziny_Pracy_Od[j], dane.Dane_Dnia.Godziny_Pracy_Do[j], Karta_Ewidencji, Helper.Strefa.Czas_Pracy_Podstawowy, dane.Numer_Relacji);
                        }
                        break;

                    case Helper.Typ_Insert_Obecnosc.Nadgodziny:
                        decimal godzNadlPlatne50 = dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_50 <= 0
                            ? (decimal)dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50
                            : (decimal)dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_50;

                        decimal godzNadlPlatne100 = dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_100 <= 0
                            ? (decimal)dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100
                            : (decimal)dane.Dane_Dnia.Liczba_Godzin_Nadliczbowych_100;

                        if (godzNadlPlatne50 > 0 || godzNadlPlatne100 > 0)
                        {
                            TimeSpan baseTime = TimeSpan.FromHours(8);
                            dane.Dane_Dnia.Godziny_Pracy_Od.Add(baseTime);
                            dane.Dane_Dnia.Godziny_Pracy_Do.Add(baseTime + TimeSpan.FromHours((double)(godzNadlPlatne50 + godzNadlPlatne100)));
                            for (int k = 0; k < dane.Dane_Dnia.Godziny_Pracy_Od.Count; k++)
                            {
                                ilosc_wpisow += Zrob_Insert_Obecnosc_Command(dataKarty, dane.Dane_Dnia.Godziny_Pracy_Od[k], dane.Dane_Dnia.Godziny_Pracy_Do[k], Karta_Ewidencji, Helper.Strefa.Czas_Pracy_Podstawowy, dane.Numer_Relacji);
                            }
                        }
                        break;

                    case Helper.Typ_Insert_Obecnosc.Nieinsertuj:
                        break;
                }
            }
            return ilosc_wpisow;
        }
        private static int Zrob_Insert_Obecnosc_Command(DateTime Data_Karty, TimeSpan startPodstawowy, TimeSpan endPodstawowy, Karta_Ewidencji Karta_Ewidencji, Helper.Strefa Strefa, string Numer_Relacji)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            try
            {
                DateTime godzOdDate = DbManager.Base_Date + startPodstawowy;
                DateTime godzDoDate = DbManager.Base_Date + endPodstawowy;
                bool duplicate = false;
                bool duplicateDE = false;
                int IdPracownika = -1;
                try
                {
                    IdPracownika = Karta_Ewidencji.Pracownik.Get_PraId();
                }
                catch (Exception ex)
                {
                    Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                }

                using (SqlCommand command = new(DbManager.Check_Duplicate_Obecnosc, DbManager.GetConnection(), DbManager.Transaction_Manager.CurrentTransaction))
                {
                    command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                    command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                    command.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = Data_Karty;
                    command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                    command.Parameters.Add("@Strefa", SqlDbType.Int).Value = Strefa;
                    duplicate = (int)command.ExecuteScalar() == 1;
                }

                using (SqlCommand command = new(DbManager.Check_Duplicate_Obecnosc, DbManager.GetConnection(), DbManager.Transaction_Manager.CurrentTransaction))
                {
                    command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                    command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                    command.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = Data_Karty;
                    command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                    command.Parameters.Add("@Strefa", SqlDbType.Int).Value = Helper.Strefa.Czas_Pracy_W_Delegacji;
                    duplicateDE = (int)command.ExecuteScalar() == 1;
                }

                if (!duplicate && !duplicateDE)
                {
                    using (SqlCommand command = new(DbManager.Insert_Obecnosci, DbManager.GetConnection(), DbManager.Transaction_Manager.CurrentTransaction))
                    {
                        command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                        command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                        command.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = Data_Karty;
                        command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                        command.Parameters.Add("@Strefa", SqlDbType.Int).Value = Strefa;
                        command.Parameters.Add("@ImieMod", SqlDbType.NVarChar, 20).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20);
                        command.Parameters.Add("@NazwiskoMod", SqlDbType.NVarChar, 50).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 50);
                        command.Parameters.Add("@DataMod", SqlDbType.DateTime).Value = Internal_Error_Logger.Last_Mod_Time;
                        command.Parameters.Add("@Numer_Relacji", SqlDbType.NVarChar, 20).Value = Numer_Relacji;
                        command.ExecuteScalar();
                    }
                    Helper.Pomiar.Avg_Insert_Obecnosc_Command = PomiaryStopWatch.Elapsed;
                    return 1;
                }
            }
            catch (SqlException ex)
            {
                Internal_Error_Logger.New_Custom_Error($"Error podczas operacji w bazie(Zrob_Insert_Obecnosc_Command): {ex.Message}", false);
                DbManager.Transaction_Manager.RollBack_Transaction();
                Helper.Pomiar.Avg_Insert_Obecnosc_Command = PomiaryStopWatch.Elapsed;

                throw new Exception($"Error podczas operacji w bazie(Zrob_Insert_Obecnosc_Command): {ex.Message}");
            }
            catch (Exception ex)
            {
                Internal_Error_Logger.New_Custom_Error($"Error: {ex.Message}", false);
                DbManager.Transaction_Manager.RollBack_Transaction();
                Helper.Pomiar.Avg_Insert_Obecnosc_Command = PomiaryStopWatch.Elapsed;
                throw new Exception($"Error: {ex.Message}");

            }
            Helper.Pomiar.Avg_Insert_Obecnosc_Command = PomiaryStopWatch.Elapsed;
            return 0;
        }
        private static async Task Dodaj_Dane_Do_Optimy(List<Karta_Ewidencji> Karty_Ewidencji, List<Prowizje> Prowizje)
        {
            await DbManager.Transaction_Manager.Create_Transaction();
            foreach (Karta_Ewidencji Karta_Ewidencji in Karty_Ewidencji)
            {
                if (Dodaj_Obecnosci_do_Optimy(Karta_Ewidencji) > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Poprawnie dodano obecnosci z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                    Console.ForegroundColor = ConsoleColor.White;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"Nie dodano żadnych obesnosci");
                    Console.ForegroundColor = ConsoleColor.White;
                }
                if (Absencja.Dodaj_Absencje_do_Optimy(Karta_Ewidencji.Absencje, Karta_Ewidencji.Pracownik, Internal_Error_Logger) > 0)
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
                if (Insert_Prowizje(Prowizje) > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Poprawnie dodano prowizje z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                    Console.ForegroundColor = ConsoleColor.White;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"Nie dodano żadnych prowizji");
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }
            DbManager.Transaction_Manager.Commit_Transaction();
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
                    }
                    Liczba_Napojow_Awaryjnych += parsed_wartosc;
                }
                offset++;
            }
            Prowizja.Suma_Wartosc_Towarow = suma_wart_towarow;
            Prowizja.Suma_Liczba_Napojow_Awaryjnych = Liczba_Napojow_Awaryjnych;
            return Prowizja;
        }
        private static int Insert_Prowizje(List<Prowizje> Prowizje)
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
                        pracId = Prowizja.Pracownik.Get_PraId();
                    }
                    catch (Exception ex)
                    {
                        Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                    }

                    if(Prowizja.Suma_Liczba_Napojow_Awaryjnych > 0)
                    {
                        using (SqlCommand command = new(DbManager.Insert_Prowizje, DbManager.GetConnection(), DbManager.Transaction_Manager.CurrentTransaction))
                        {
                            command.Parameters.Add("@PracID", SqlDbType.Int).Value = pracId;
                            command.Parameters.Add("@NowaWartosc", SqlDbType.Decimal).Value = Prowizja.Suma_Liczba_Napojow_Awaryjnych;
                            command.Parameters.Add("@NazwaAtrybutu", SqlDbType.NVarChar, 41).Value = "Prowizja za wydane napoje awaryjne";
                            command.Parameters.Add("@ATHDataOd", SqlDbType.DateTime).Value = Data_Od;
                            command.Parameters.Add("@ATHDataDo", SqlDbType.DateTime).Value = Data_Do;
                            count += command.ExecuteNonQuery();
                        }
                    }
                    if (Prowizja.Suma_Wartosc_Towarow > 0)
                    {
                        using (SqlCommand command = new(DbManager.Insert_Prowizje, DbManager.GetConnection(), DbManager.Transaction_Manager.CurrentTransaction))
                        {
                            command.Parameters.Add("@PracID", SqlDbType.Int).Value = pracId;
                            command.Parameters.Add("@NowaWartosc", SqlDbType.Decimal).Value = Prowizja.Suma_Wartosc_Towarow;
                            command.Parameters.Add("@NazwaAtrybutu", SqlDbType.NVarChar, 41).Value = "Prowizja za towar";
                            command.Parameters.Add("@ATHDataOd", SqlDbType.DateTime).Value = Data_Od;
                            command.Parameters.Add("@ATHDataDo", SqlDbType.DateTime).Value = Data_Do;
                            count += command.ExecuteNonQuery();

                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                Internal_Error_Logger.New_Custom_Error($"Error podczas operacji w bazie(Insert_Prowizje): {ex.Message}", false);
                DbManager.Transaction_Manager.RollBack_Transaction();
                throw new Exception($"Error podczas operacji w bazie(Insert_Prowizje): {ex.Message}");
            }
            catch (Exception ex)
            {
                Internal_Error_Logger.New_Custom_Error($"Error: {ex.Message}", false);
                DbManager.Transaction_Manager.RollBack_Transaction();
                throw new Exception($"Error: {ex.Message}");
            }
            return count;
        }
    }
}
