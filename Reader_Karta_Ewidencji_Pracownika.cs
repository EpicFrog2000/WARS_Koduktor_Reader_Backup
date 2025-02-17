using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using Microsoft.Data.SqlClient;
using static Konduktor_Reader.Helper;

//Jest przekopiowane praktycznie 1 do 1 ze starego programu dlatego jest taki murzyński kod
// TODO w przyszłości zrobić to lepiej

namespace Konduktor_Reader
{
    internal class Reader_Karta_Ewidencji_Pracownika
    {
        private class Karta_Ewidencji_Pracownika
        {
            public Pracownik Pracownik = new();
            public int Miesiac = 0;
            public int Rok = 0;
            public List<Dane_Karty> Dane_Karty = [];
            public List<Absencja> Absencje = [];
            public int Set_Data(string Wartosc)
            {
                if (string.IsNullOrEmpty(Wartosc))
                {
                    return 1;
                }
                try
                {
                    DateTime data;
                    if (!DateTime.TryParse(Wartosc, out data))
                    {
                        return 1;
                    }
                    Rok = data.Year;
                    Miesiac = data.Month;
                }
                catch
                {
                    return 1;
                }
                return 0;
            }
            public void Set_Miesiac(string nazwa)
            {
                if (!string.IsNullOrEmpty(nazwa))
                {
                    if (nazwa.ToLower() == "styczeń")
                    {
                        Miesiac = 1;
                    }
                    else if (nazwa.ToLower() == "i")
                    {
                        Miesiac = 1;
                    }
                    else if (nazwa.ToLower() == "luty")
                    {
                        Miesiac = 2;
                    }
                    else if (nazwa.ToLower() == "ii")
                    {
                        Miesiac = 2;
                    }
                    else if (nazwa.ToLower() == "marzec")
                    {
                        Miesiac = 3;
                    }
                    else if (nazwa.ToLower() == "iii")
                    {
                        Miesiac = 3;
                    }
                    else if (nazwa.ToLower() == "kwiecień")
                    {
                        Miesiac = 4;
                    }
                    else if (nazwa.ToLower() == "iv")
                    {
                        Miesiac = 4;
                    }
                    else if (nazwa.ToLower() == "maj")
                    {
                        Miesiac = 5;
                    }
                    else if (nazwa.ToLower() == "v")
                    {
                        Miesiac = 5;
                    }
                    else if (nazwa.ToLower() == "czerwiec")
                    {
                        Miesiac = 6;
                    }
                    else if (nazwa.ToLower() == "vi")
                    {
                        Miesiac = 6;
                    }
                    else if (nazwa.ToLower() == "lipiec")
                    {
                        Miesiac = 7;
                    }
                    else if (nazwa.ToLower() == "vii")
                    {
                        Miesiac = 7;
                    }
                    else if (nazwa.ToLower() == "sierpień")
                    {
                        Miesiac = 8;
                    }
                    else if (nazwa.ToLower() == "viii")
                    {
                        Miesiac = 8;
                    }
                    else if (nazwa.ToLower() == "wrzesień")
                    {
                        Miesiac = 9;
                    }
                    else if (nazwa.ToLower() == "ix")
                    {
                        Miesiac = 9;
                    }
                    else if (nazwa.ToLower() == "październik")
                    {
                        Miesiac = 10;
                    }
                    else if (nazwa.ToLower() == "x")
                    {
                        Miesiac = 10;
                    }
                    else if (nazwa.ToLower() == "listopad")
                    {
                        Miesiac = 11;
                    }
                    else if (nazwa.ToLower() == "xi")
                    {
                        Miesiac = 11;
                    }
                    else if (nazwa.ToLower() == "grudzień")
                    {
                        Miesiac = 12;
                    }
                    else if (nazwa.ToLower() == "xii")
                    {
                        Miesiac = 12;
                    }
                    else
                    {
                        Miesiac = 0;
                    }
                }
            }
        }
        private class Dane_Karty
        {
            public int Dzien = 0;
            public List<TimeSpan> Godziny_Rozpoczecia_Pracy = [];
            public List<TimeSpan> Godziny_Zakonczenia_Pracy = [];
            public decimal Godziny_Nadliczbowe_Platne_Z_Dodatkiem_50 = 0;
            public decimal Godziny_Nadliczbowe_Platne_Z_Dodatkiem_100 = 0;
            public decimal Ilosc_Godzin_Do_Odbioru = 0;
            public decimal Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach = 0;
            public void Podziel_Nadgodziny()
            {
                if (Godziny_Zakonczenia_Pracy.Count == 0) return;


                TimeSpan shiftEnd = Godziny_Zakonczenia_Pracy[^1];

                TimeSpan overtime50 = TimeSpan.FromHours((double)Godziny_Nadliczbowe_Platne_Z_Dodatkiem_50);


                TimeSpan overtime100 = TimeSpan.FromHours((double)Godziny_Nadliczbowe_Platne_Z_Dodatkiem_100);

                TimeSpan overtimeStart = shiftEnd - (overtime50 + overtime100);
                if (overtimeStart < TimeSpan.Zero)
                {
                    overtimeStart = TimeSpan.FromHours(24) + overtimeStart;
                }

                List<TimeSpan> newGodziny_Pracy_Od = new List<TimeSpan>();
                List<TimeSpan> newGodziny_Pracy_Do = new List<TimeSpan>();


                for (int i = 0; i < Godziny_Rozpoczecia_Pracy.Count; i++)
                {
                    if (Godziny_Zakonczenia_Pracy[i] == TimeSpan.Zero)
                    {
                        if (overtimeStart > Godziny_Rozpoczecia_Pracy[i] && overtimeStart > Godziny_Zakonczenia_Pracy[i])
                        {

                            if (Godziny_Rozpoczecia_Pracy[i] != overtimeStart)
                            {
                                newGodziny_Pracy_Od.Add(Godziny_Rozpoczecia_Pracy[i]);
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
                    else if (Godziny_Zakonczenia_Pracy[i] > overtimeStart)
                    {
                        if (Godziny_Zakonczenia_Pracy[i] < Godziny_Rozpoczecia_Pracy[i])
                        {
                            newGodziny_Pracy_Od.Add(Godziny_Rozpoczecia_Pracy[i]);
                            newGodziny_Pracy_Do.Add(TimeSpan.FromHours(24));
                            newGodziny_Pracy_Od.Add(TimeSpan.Zero);
                            newGodziny_Pracy_Do.Add(overtimeStart);
                        }
                        else
                        {
                            if (Godziny_Rozpoczecia_Pracy[i] != overtimeStart)
                            {
                                newGodziny_Pracy_Od.Add(Godziny_Rozpoczecia_Pracy[i]);
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
                    newGodziny_Pracy_Od.Add(Godziny_Rozpoczecia_Pracy[i]);
                    newGodziny_Pracy_Do.Add(Godziny_Zakonczenia_Pracy[i]);
                }
                Godziny_Rozpoczecia_Pracy = newGodziny_Pracy_Od;
                Godziny_Zakonczenia_Pracy = newGodziny_Pracy_Do;
            }
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
        private class Absencja
        {
            public int Dzien = 0;
            public int Miesiac = 0;
            public int Rok = 0;
            public string Nazwa = string.Empty;
            public decimal Liczba_Godzin_Absencji = 0;
            public RodzajAbsencji Rodzaj_Absencji = 0;
        }
        private static Error_Logger Internal_Error_Logger = new(true);
        public static void Process_Zakladka(IXLWorksheet Zakladka, Error_Logger Error_Logger)
        {
            Internal_Error_Logger = Error_Logger;
            List<Karta_Ewidencji_Pracownika> Karty_Ewidencji_Pracownika = [];
            List<Helper.Current_Position> Pozycje = Helper.Find_Starting_Points(Zakladka, "Dzień", false);

            foreach (Helper.Current_Position Pozycja in Pozycje)
            {
                Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika = new();
                Get_Dane_Naglowka_Karty(ref Karta_Ewidencji_Pracownika, Pozycja, Zakladka);
                Get_Dane_Dni(ref Karta_Ewidencji_Pracownika, Pozycja, Zakladka);
                Karty_Ewidencji_Pracownika.Add(Karta_Ewidencji_Pracownika);
            }

            foreach (Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika in Karty_Ewidencji_Pracownika)
            {
                Dodaj_Dane_Do_Optimy(Karta_Ewidencji_Pracownika);
            }
        }
        private static void Get_Dane_Naglowka_Karty(ref Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika, Current_Position StartKarty, IXLWorksheet Zakladka)
        {
            //wczytaj date
            string dane = Zakladka.Cell(StartKarty.Row - 3, StartKarty.Col + 4).GetFormattedString().Trim().ToLower();
            for (int i = 0; i < 12; i++)
            {
                if (string.IsNullOrEmpty(dane))
                {
                    dane = Zakladka.Cell(StartKarty.Row - 3, StartKarty.Col + 4 + i).GetFormattedString().Trim().ToLower();
                }
                else
                {
                    //here try to get data i rok
                    if (dane.EndsWith("r"))
                    {
                        dane = dane.Substring(0, dane.Length - 1).Trim();
                    }
                    if (dane.EndsWith("r."))
                    {
                        dane = dane.Substring(0, dane.Length - 2).Trim();
                    }

                    string[] dateFormats = { "dd.MM.yyyy" };
                    if (DateTime.TryParseExact(dane, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedData))
                    {
                        Karta_Ewidencji_Pracownika.Miesiac = parsedData.Month;
                        Karta_Ewidencji_Pracownika.Rok = parsedData.Year;
                    }
                    else
                    {
                        if (dane.Contains("pażdziernik"))
                        {
                            dane = dane.Replace("pażdziernik", "październik");
                        }
                        if (Karta_Ewidencji_Pracownika.Set_Data(dane) == 1)
                        {
                            if (dane.Split(" ").Length == 2)
                            {
                                string[] ndata = dane.Split(" ");
                                try
                                {
                                    Karta_Ewidencji_Pracownika.Set_Miesiac(ndata[0]);
                                    if (int.TryParse(Regex.Replace(ndata[1], @"\D", ""), out int rok))
                                    {
                                        Karta_Ewidencji_Pracownika.Rok = rok;
                                    }
                                }
                                catch { }
                            }
                            else if (dane.Split(" ").Length == 3)
                            {
                                string[] ndata = dane.Split(" ");
                                try
                                {
                                    Karta_Ewidencji_Pracownika.Set_Miesiac(ndata[1]);
                                    if (int.TryParse(ndata[2], out int rok))
                                    {
                                        Karta_Ewidencji_Pracownika.Rok = rok;
                                    }
                                }
                                catch { }
                            }
                            else
                            {
                                if (dane.Split(" ").Count() > 1)
                                {
                                    //wez 2 od tylu
                                    string[] ndata = dane.Split(" ");
                                    try
                                    {
                                        Karta_Ewidencji_Pracownika.Set_Miesiac(ndata[^2]);
                                        if (int.TryParse(ndata[^1], out int rok))
                                        {
                                            Karta_Ewidencji_Pracownika.Rok = rok;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                    if (Karta_Ewidencji_Pracownika.Miesiac == 0 || Karta_Ewidencji_Pracownika.Rok == 0)
                    {
                        dane = Zakladka.Cell(StartKarty.Row - 4, StartKarty.Col + 4 + i - 1).GetFormattedString().Trim().ToLower();
                        if (!string.IsNullOrEmpty(dane) && DateTime.TryParseExact(dane, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedData2))
                        {
                            Karta_Ewidencji_Pracownika.Miesiac = parsedData2.Month;
                            Karta_Ewidencji_Pracownika.Rok = parsedData2.Year;
                        }
                    }

                    if (Karta_Ewidencji_Pracownika.Miesiac != 0 && Karta_Ewidencji_Pracownika.Rok != 0)
                    {
                        break;
                    }
                }
            }
            if (Karta_Ewidencji_Pracownika.Miesiac == 0 || Karta_Ewidencji_Pracownika.Rok == 0)
            {
                Internal_Error_Logger.New_Error(dane, "data", StartKarty.Col + 11, StartKarty.Row - 3, $"Nie wykryto daty w pliku. Oczekiwana dat między kolumna[{StartKarty.Col + 4}] rząd[{StartKarty.Row - 3}] a kolumna[{StartKarty.Col + 4 + 11}] rząd[{StartKarty.Row - 3}]");
                throw new Exception(Internal_Error_Logger.Get_Error_String());
            }

            dane = "";
            //wczytaj nazwisko i imie
            try
            {
                for (int i = 0; i < 5; i++)
                {
                    try
                    {
                        Pracownik pracownik = Get_Pracownik(i, StartKarty, Zakladka);
                        if (pracownik.Nazwisko != "" && pracownik.Imie != "")
                        {
                            Karta_Ewidencji_Pracownika.Pracownik.Nazwisko = pracownik.Nazwisko;
                            Karta_Ewidencji_Pracownika.Pracownik.Imie = pracownik.Imie;
                            goto Found;
                        }
                    }
                    catch
                    {
                    }
                }
                if (string.IsNullOrEmpty(Karta_Ewidencji_Pracownika.Pracownik.Imie) && string.IsNullOrEmpty(Karta_Ewidencji_Pracownika.Pracownik.Nazwisko))
                {
                    Internal_Error_Logger.New_Error(dane, "nazwisko i imie", StartKarty.Col, StartKarty.Row - 2, $"Nie znaleziono pola z nazwiskiem i imieniem między kolumna[{StartKarty.Col}] rząd[{StartKarty.Row - 2}] a kolumna[{StartKarty.Col + 5}] rząd[{StartKarty.Row - 2}]");
                    throw new Exception(Internal_Error_Logger.Get_Error_String());
                }

                Found:
                ;
                // znajdz akronim w prawo
                for (int i = 4; i < 9; i++)
                {
                    dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col + i).GetFormattedString().Trim().ToLower();
                    if (!string.IsNullOrEmpty(dane))
                    {
                        if (dane.Contains("akronim"))
                        {
                            dane.Replace("akronim", "").Replace(":", "");
                            if (!string.IsNullOrEmpty(dane))
                            {
                                if (int.TryParse(dane, out int parseAkr))
                                {
                                    Karta_Ewidencji_Pracownika.Pracownik.Akronim = dane;
                                }
                            }
                            else
                            {
                                dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col + i + 1).GetFormattedString().Trim().ToLower();
                                if (!string.IsNullOrEmpty(dane))
                                {
                                    if (int.TryParse(dane, out int parseAkr))
                                    {
                                        Karta_Ewidencji_Pracownika.Pracownik.Akronim = dane;
                                    }
                                }
                            }
                        }
                        else
                        {
                            dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col + i + 1).GetFormattedString().Trim().ToLower();
                            if (!string.IsNullOrEmpty(dane))
                            {
                                if (int.TryParse(dane, out int parseAkr))
                                {
                                    Karta_Ewidencji_Pracownika.Pracownik.Akronim = dane;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Internal_Error_Logger.New_Error(dane, "Imie nazwisko akronim", StartKarty.Row - 2, StartKarty.Col, "Nieznany format");
                throw new Exception($"{ex}, {Internal_Error_Logger.Get_Error_String()}");
            }

        }
        private static Pracownik Get_Pracownik(int sposob, Current_Position StartKarty, IXLWorksheet Zakladka)
        {
            string[] wordsToRemove = ["IMIĘ:", "IMIE:", "NAZWISKO:", "NAZWISKO", " IMIE", "IMIĘ", ":"];
            string dane = "";
            Pracownik pracownik = new();
            switch (sposob)
            {
                case 0: // Karta pracy: Nazwisko Imie #lub# Karta pracy: Nazwisko Imie akronim #lub# Karta pracy: akronim Nazwisko Imie
                    dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        foreach (string word in wordsToRemove)
                        {
                            string pattern = $@"\b{Regex.Escape(word)}\b";
                            dane = Regex.Replace(dane, pattern, "", RegexOptions.IgnoreCase);
                        }
                        dane = Regex.Replace(dane, @"\s+", " ").Trim();
                        if (dane.Contains("KARTA PRACY:"))
                        {
                            dane = dane.Replace("KARTA PRACY:", "").Trim();
                        }
                        if (dane.Contains("KARTA PRACY"))
                        {
                            dane = dane.Replace("KARTA PRACY", "").Trim();
                        }
                        if (!string.IsNullOrEmpty(dane))
                        {
                            string[] parts = dane.Trim().Split(' ');
                            if (parts.Length == 2)
                            {
                                pracownik.Nazwisko = dane.Trim().Split(' ')[0];
                                pracownik.Imie = dane.Trim().Split(' ')[1];
                            }
                            else if (parts.Length == 3)
                            {
                                if (int.TryParse(parts[0], out int parsedValue))
                                {
                                    pracownik.Akronim = parts[0];
                                    pracownik.Nazwisko = dane.Trim().Split(' ')[1];
                                    pracownik.Imie = dane.Trim().Split(' ')[2];
                                }
                                else if (int.TryParse(parts[2], out int parsedValue2))
                                {
                                    pracownik.Akronim = parts[2];
                                    pracownik.Nazwisko = dane.Trim().Split(' ')[0];
                                    pracownik.Imie = dane.Trim().Split(' ')[1];
                                    return pracownik;
                                }
                            }
                        }
                    }
                    return pracownik;
                case 1:  // Karta pracy: || Nazwisko Imie
                    dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        foreach (string word in wordsToRemove)
                        {
                            string pattern = $@"\b{Regex.Escape(word)}\b";
                            dane = Regex.Replace(dane, pattern, "", RegexOptions.IgnoreCase);
                        }
                        dane = Regex.Replace(dane, @"\s+", " ").Trim();
                        if (dane.Contains("KARTA PRACY:"))
                        {
                            dane = dane.Replace("KARTA PRACY:", "").Trim();
                        }
                        if (dane.Contains("KARTA PRACY"))
                        {
                            dane = dane.Replace("KARTA PRACY", "").Trim();
                        }
                        if (string.IsNullOrEmpty(dane))
                        {
                            dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                            if (!string.IsNullOrEmpty(dane))
                            {
                                string[] parts = dane.Trim().Split(' ');
                                if (parts.Length == 2)
                                {
                                    pracownik.Nazwisko = parts[0];
                                    pracownik.Imie = parts[1];
                                    return pracownik;
                                }
                            }
                        }
                    }
                    break;
                case 2:  // Karta pracy: || Nazwisko | Imie
                    dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        foreach (string word in wordsToRemove)
                        {
                            string pattern = $@"\b{Regex.Escape(word)}\b";
                            dane = Regex.Replace(dane, pattern, "", RegexOptions.IgnoreCase);
                        }
                        dane = Regex.Replace(dane, @"\s+", " ").Trim();
                        if (dane.Contains("KARTA PRACY:"))
                        {
                            dane = dane.Replace("KARTA PRACY:", "").Trim();
                        }
                        if (dane.Contains("KARTA PRACY"))
                        {
                            dane = dane.Replace("KARTA PRACY", "").Trim();
                        }
                        if (string.IsNullOrEmpty(dane))
                        {
                            dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                            if (!string.IsNullOrEmpty(dane))
                            {
                                pracownik.Nazwisko = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                                pracownik.Imie = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col + 3).GetFormattedString().Trim().Replace("  ", " ");
                                return pracownik;
                            }
                        }
                    }
                    break;
                case 3: // Nazwisko Imie
                    dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        string[] parts = dane.Trim().Split(' ');
                        if (parts.Length == 2)
                        {
                            pracownik.Nazwisko = parts[0];
                            pracownik.Imie = parts[1];
                            return pracownik;
                        }
                    }
                    break;
                case 4: // Nazwisko | Imie
                    dane = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (!string.IsNullOrEmpty(dane))
                    {
                        string[] parts = dane.Trim().Split(' ');
                        if (parts.Length == 1)
                        {
                            if (!string.IsNullOrEmpty(dane))
                            {
                                pracownik.Nazwisko = dane;
                                pracownik.Imie = Zakladka.Cell(StartKarty.Row - 2, StartKarty.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
                                return pracownik;
                            }
                        }
                    }
                    break;
                default:
                    return pracownik;
            }
            return pracownik;
        }
        private static void Get_Dane_Dni(ref Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika, Current_Position Pozycja, IXLWorksheet Zakladka)
        {
            Pozycja.Row += 3;
            while (true)
            {
                // dzien
                Dane_Karty Dane_Karty = new();
                string dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(dane))
                {
                    break;
                }
                if (!Helper.Try_Get_Type_From_String<int>(dane, ref Dane_Karty.Dzien))
                {
                    break;
                }
                // godz rozp pracy
                dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    foreach (string d in dane.Split(' '))
                    {
                        if (!Helper.Try_Get_Type_From_String<List<TimeSpan>>(d, ref Dane_Karty.Godziny_Rozpoczecia_Pracy))
                        {
                            Internal_Error_Logger.New_Error(dane, "Godzina Rozpoczęcia pracy", Pozycja.Col + 1, Pozycja.Row, "Zły format Godziny");
                            throw new Exception(Internal_Error_Logger.Get_Error_String());
                        }
                    }
                }

                // godz zak pracy
                dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    foreach (string d in dane.Split(' '))
                    {
                        if (!Helper.Try_Get_Type_From_String<List<TimeSpan>>(d, ref Dane_Karty.Godziny_Zakonczenia_Pracy))
                        {
                            Internal_Error_Logger.New_Error(dane, "Godzina Zakończenia pracy", Pozycja.Col + 2, Pozycja.Row, "Zły format Godziny");
                            throw new Exception(Internal_Error_Logger.Get_Error_String());
                        }
                    }
                }

                // absencja
                dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 3).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    Absencja Absencja = new();
                    Absencja.Dzien = Dane_Karty.Dzien;
                    Absencja.Miesiac = Karta_Ewidencji_Pracownika.Miesiac;
                    Absencja.Rok = Karta_Ewidencji_Pracownika.Rok;
                    if (!Helper.Try_Get_Type_From_String<string>(dane.ToUpper(), ref Absencja.Nazwa))
                    {
                        Internal_Error_Logger.New_Error(dane, "Nazwa Absencji", Pozycja.Col + 3, Pozycja.Row, "Zły format Nazwy absencji");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }
                    if (!RodzajAbsencji.TryParse(Absencja.Nazwa, out Absencja.Rodzaj_Absencji))
                    {
                        Internal_Error_Logger.New_Error(dane, "Rodzaj Absencji", Pozycja.Col + 3, Pozycja.Row, "Nierozpoznany rodzaj absencji");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }

                    dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Absencja.Liczba_Godzin_Absencji))
                    {
                        //Internal_Error_Logger.New_Error(dane, "Liczba godz Absencji", Pozycja.Col + 18, Pozycja.Row + Row_Offset, "Zły format lub bark Liczba godz Absencji");
                        //throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }
                    Karta_Ewidencji_Pracownika.Absencje.Add(Absencja);
                }

                //Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach
                dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 5).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Karty.Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach))
                    {
                        Internal_Error_Logger.New_Error(dane, "Liczba Godzin Do Odbioru Za Prace W Nadgodzinach", Pozycja.Col + 5, Pozycja.Row, "Zły format Liczba Godzin Do Odbioru Za Prace W Nadgodzinach");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }
                }

                // Godziny_Nadliczbowe_Platne_Z_Dodatkiem_50
                dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 9).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Karty.Godziny_Nadliczbowe_Platne_Z_Dodatkiem_50))
                    {
                        Internal_Error_Logger.New_Error(dane, "Godziny Nadliczbowe Platne Z Dodatkiem 50", Pozycja.Col + 9, Pozycja.Row, "Zły format Godziny Nadliczbowe Platne Z Dodatkiem 50");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }
                }

                // Godziny_Nadliczbowe_Platne_Z_Dodatkiem_100
                dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 10).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Karty.Godziny_Nadliczbowe_Platne_Z_Dodatkiem_100))
                    {
                        Internal_Error_Logger.New_Error(dane, "Godziny Nadliczbowe Platne Z Dodatkiem 100", Pozycja.Col + 10, Pozycja.Row, "Zły format Godziny Nadliczbowe Platne Z Dodatkiem 100");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }
                }
                Pozycja.Row++;
                Karta_Ewidencji_Pracownika.Dane_Karty.Add(Dane_Karty);
            }
        }
        private static void Dodaj_Dane_Do_Optimy(Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika)
        {
            //TODO
            using (SqlConnection connection = new SqlConnection(Program.config.Optima_Conection_String))
            {
                connection.Open();
                using (SqlTransaction transaction = connection.BeginTransaction())
                {
                    if (Dodaj_Obecnosci_do_Optimy(Karta_Ewidencji_Pracownika, transaction, connection) > 0)
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
                    if (Dodaj_Absencje_do_Optimy(Karta_Ewidencji_Pracownika.Absencje, transaction, connection, Karta_Ewidencji_Pracownika.Pracownik) > 0)
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
                    if (Dodaj_Godz_Odbior_Do_Optimy(Karta_Ewidencji_Pracownika, transaction, connection) > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Poprawnie dodano odbiory nadgodzin z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Nie dodano żadnych odbiorow nadgodzin");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    transaction.Commit();
                }
            }
        }
        private static int Dodaj_Obecnosci_do_Optimy(Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika, SqlTransaction transaction, SqlConnection connection)
        {

            HashSet<DateTime> Pasujace_Daty = [];
            foreach (Dane_Karty daneKarty in Karta_Ewidencji_Pracownika.Dane_Karty)
            {
                Pasujace_Daty.Add(new DateTime(Karta_Ewidencji_Pracownika.Rok, Karta_Ewidencji_Pracownika.Miesiac, daneKarty.Dzien));
            }
            DateTime startDate = new(Karta_Ewidencji_Pracownika.Rok, Karta_Ewidencji_Pracownika.Miesiac, 1);
            DateTime endDate = new(Karta_Ewidencji_Pracownika.Rok, Karta_Ewidencji_Pracownika.Miesiac, DateTime.DaysInMonth(Karta_Ewidencji_Pracownika.Rok, Karta_Ewidencji_Pracownika.Miesiac));
            for (DateTime dzien = startDate; dzien <= endDate; dzien = dzien.AddDays(1))
            {
                if (!Pasujace_Daty.Contains(dzien))
                {
                    Zrob_Insert_Obecnosc_Command(connection, transaction, dzien, TimeSpan.Zero, TimeSpan.Zero, Karta_Ewidencji_Pracownika, 1); // 1 - pusta strefa
                }
            }


            int ilosc_wpisow = 0;
            foreach (Dane_Karty Dane_Karty in Karta_Ewidencji_Pracownika.Dane_Karty)
            {
                if (DateTime.TryParse($"{Karta_Ewidencji_Pracownika.Rok}-{Karta_Ewidencji_Pracownika.Miesiac:D2}-{Dane_Karty.Dzien:D2}", out DateTime Data_Karty))
                {
                    if (Dane_Karty.Godziny_Rozpoczecia_Pracy.Count >= 1)
                    {
                        //Dane_Karty.Podziel_Nadgodziny();
                        
                        for (int j = 0; j < Dane_Karty.Godziny_Rozpoczecia_Pracy.Count; j++)
                        {
                            ilosc_wpisow += Zrob_Insert_Obecnosc_Command(connection, transaction, Data_Karty, Dane_Karty.Godziny_Rozpoczecia_Pracy[j], Dane_Karty.Godziny_Zakonczenia_Pracy[j], Karta_Ewidencji_Pracownika, 2);
                        }
                    }
                    else
                    {
                        if (Dane_Karty.Godziny_Nadliczbowe_Platne_Z_Dodatkiem_50 > 0 || Dane_Karty.Godziny_Nadliczbowe_Platne_Z_Dodatkiem_100 > 0)
                        {
                            TimeSpan baseTime = TimeSpan.FromHours(8);
                            Dane_Karty.Godziny_Rozpoczecia_Pracy.Add(baseTime);
                            Dane_Karty.Godziny_Zakonczenia_Pracy.Add(baseTime + TimeSpan.FromHours((double)(Dane_Karty.Godziny_Nadliczbowe_Platne_Z_Dodatkiem_50 + Dane_Karty.Godziny_Nadliczbowe_Platne_Z_Dodatkiem_100)));
                            for (int k = 0; k < Dane_Karty.Godziny_Rozpoczecia_Pracy.Count; k++)
                            {
                                ilosc_wpisow += Zrob_Insert_Obecnosc_Command(connection, transaction, Data_Karty, Dane_Karty.Godziny_Rozpoczecia_Pracy[k], Dane_Karty.Godziny_Zakonczenia_Pracy[k], Karta_Ewidencji_Pracownika, 2);
                            }
                        }
                        else
                        {
                            Zrob_Insert_Obecnosc_Command(connection, transaction, Data_Karty, TimeSpan.Zero, TimeSpan.Zero, Karta_Ewidencji_Pracownika, 1); // 1 - pusta strefa
                        }
                    }
                }
                
            }
            return ilosc_wpisow;
        }
        private static int Zrob_Insert_Obecnosc_Command(SqlConnection connection, SqlTransaction transaction, DateTime Data_Karty, TimeSpan startPodstawowy, TimeSpan endPodstawowy, Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika, int Typ_Pracy)
        {
            if (startPodstawowy == endPodstawowy && startPodstawowy != TimeSpan.Zero)
            {
                return 0;
            }

            try
            {
                DateTime godzOdDate = Helper.baseDate + startPodstawowy;
                DateTime godzDoDate = Helper.baseDate + endPodstawowy;
                bool duplicate = false;
                int IdPracownika = -1;
                try
                {
                    IdPracownika = Karta_Ewidencji_Pracownika.Pracownik.Get_PraId(connection, transaction);
                }
                catch (Exception ex)
                {
                    connection.Close();
                    Internal_Error_Logger.New_Custom_Error(ex.Message + " z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}" + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                }

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
                    using (SqlCommand insertCmd = new(@"
DECLARE @EXISTSDZIEN DATETIME = (SELECT PracPracaDni.PPR_Data FROM cdn.PracPracaDni WITH (NOLOCK) WHERE PPR_PraId = @PRI_PraId and PPR_Data = @DataInsert)
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
                    ,[PPR_Zrodlo])
                VALUES
                    (@PRI_PraId
                    ,@DataInsert
                    ,@DataMod
                    ,@DataMod
                    ,@ImieMod
                    ,@NazwiskoMod
                    ,@ImieMod
                    ,@NazwiskoMod
                    ,0)
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
		1);
", connection, transaction))
                    {
                        insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                        insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                        insertCmd.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = Data_Karty;
                        insertCmd.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                        insertCmd.Parameters.Add("@TypPracy", SqlDbType.Int).Value = Typ_Pracy;
                        insertCmd.Parameters.Add("@ImieMod", SqlDbType.NVarChar, 20).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20);
                        insertCmd.Parameters.Add("@NazwiskoMod", SqlDbType.NVarChar, 50).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 50);
                        insertCmd.Parameters.Add("@DataMod", SqlDbType.DateTime).Value = Internal_Error_Logger.Last_Mod_Time;
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
        private static int Dodaj_Absencje_do_Optimy(List<Absencja> Absencje, SqlTransaction tran, SqlConnection connection, Pracownik Pracownik)
        {
            int ilosc_wpisow = 0;
            List<List<Absencja>> ListyAbsencji = Podziel_Absencje_Na_Osobne(Absencje);
            foreach (List<Absencja> ListaAbsencji in ListyAbsencji)
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
                    Internal_Error_Logger.New_Custom_Error($"W programie brak dopasowanego kodu nieobecnosci: {ListaAbsencji[0].Rodzaj_Absencji} w dniu {new DateTime(ListaAbsencji[0].Rok, ListaAbsencji[0].Miesiac, ListaAbsencji[0].Dzien)} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki}. Absencja nie dodana.");
                    Exception e = new();
                    e.Data["Kod"] = 42069;
                    throw e;
                }
                int dni_robocze = Ile_Dni_Roboczych(ListaAbsencji);
                int dni_calosc = ListaAbsencji.Count;

                bool duplicate = false;

                int IdPracownika = -1;
                try
                {
                    IdPracownika = Pracownik.Get_PraId(connection, tran);
                }
                catch (Exception ex)
                {
                    connection.Close();
                    Internal_Error_Logger.New_Custom_Error(ex.Message + " z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}" + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                }

                using (SqlCommand cmd = new(@"IF EXISTS (
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
", connection, tran))
                {
                    cmd.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                    cmd.Parameters.Add("@NazwaNieobecnosci", SqlDbType.NVarChar, 50).Value = nazwa_nieobecnosci;
                    cmd.Parameters.Add("@DniPracy", SqlDbType.Int).Value = dni_robocze;
                    cmd.Parameters.Add("@DniKalendarzowe", SqlDbType.Int).Value = dni_calosc;
                    cmd.Parameters.Add("@Przyczyna", SqlDbType.NVarChar, 50).Value = przyczyna;
                    cmd.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = Data_Absencji_Start;
                    cmd.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = Helper.baseDate;
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
                            insertCmd.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                            insertCmd.Parameters.Add("@NazwaNieobecnosci", SqlDbType.NVarChar, 50).Value = nazwa_nieobecnosci;
                            insertCmd.Parameters.Add("@DniPracy", SqlDbType.Int).Value = dni_robocze;
                            insertCmd.Parameters.Add("@DniKalendarzowe", SqlDbType.Int).Value = dni_calosc;
                            insertCmd.Parameters.Add("@Przyczyna", SqlDbType.NVarChar, 50).Value = przyczyna;
                            insertCmd.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = Data_Absencji_Start;
                            insertCmd.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = Helper.baseDate;
                            insertCmd.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = Data_Absencji_End;
                            insertCmd.Parameters.Add("@ImieMod", SqlDbType.NVarChar, 20).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20);
                            insertCmd.Parameters.Add("@NazwiskoMod", SqlDbType.NVarChar, 50).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20);
                            insertCmd.Parameters.Add("@DataMod", SqlDbType.DateTime).Value = Internal_Error_Logger.Last_Mod_Time;
                            insertCmd.ExecuteScalar();
                        }
                    }

                    catch (FormatException ex)
                    {
                        Internal_Error_Logger.New_Custom_Error($"{ex.Message}");

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
        private static int Ile_Dni_Roboczych(List<Absencja> Absencje)
        {
            return Absencje.Count(Absencja =>
            {
                DateTime absenceDate = new(Absencja.Rok, Absencja.Miesiac, Absencja.Dzien);
                return absenceDate.DayOfWeek != DayOfWeek.Saturday && absenceDate.DayOfWeek != DayOfWeek.Sunday;
            });
        }
        private static int Dodaj_Godz_Odbior_Do_Optimy(Karta_Ewidencji_Pracownika karta, SqlTransaction transaction, SqlConnection connection)
        {
            int ilosc_wpisow = 0;
            foreach (var dane_Dni in karta.Dane_Karty)
            {
                if(dane_Dni.Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach > 0)
                {
                    int IdPracownika = karta.Pracownik.Get_PraId(connection, transaction);
                    var Ilosc_Godzin = dane_Dni.Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach;
                    DateTime godzOdDate = Helper.baseDate + TimeSpan.FromHours(8);
                    DateTime godzDoDate = Helper.baseDate + TimeSpan.FromHours(8) + TimeSpan.FromHours((double)dane_Dni.Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach);
                    bool duplicate = false;
                    using (SqlCommand cmd = new SqlCommand(@"
    DECLARE @EXISTSDZIEN INT;
    DECLARE @EXISTSDATA INT;
    SET @EXISTSDZIEN = (SELECT COUNT(PPR_Data) FROM cdn.PracPracaDni WHERE PPR_PraId = @PRI_PraId AND PPR_Data = @DataInsert);
    SET @EXISTSDATA = (
        SELECT COUNT(*)
        FROM CDN.PracPracaDniGodz 
        WHERE PGR_OdbNadg = 4
            AND PGR_Strefa = 2
            AND PGR_OdGodziny = DATEADD(MINUTE, 0, @GodzOdDate)
            AND PGR_DoGodziny = DATEADD(MINUTE, 0, @GodzDoDate)
            AND PGR_PprId = (SELECT PPR_PprId FROM cdn.PracPracaDni WHERE CAST(PPR_Data AS datetime) = @DataInsert AND PPR_PraId = @PRI_PraId)
    );
    SELECT CASE 
        WHEN @EXISTSDZIEN > 0 AND @EXISTSDATA > 0 THEN 1
        ELSE 0
    END;", connection, transaction))
                    {
                        cmd.Parameters.AddWithValue("@PRI_PraId", IdPracownika);
                        cmd.Parameters.AddWithValue("@TypPracy", 2);
                        cmd.Parameters.AddWithValue("@TypNadg", 4);
                        cmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                        cmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                        cmd.Parameters.AddWithValue("@DataInsert", DateTime.Parse($"{karta.Rok}-{karta.Miesiac:D2}-{dane_Dni.Dzien:D2}"));
                        if ((int)cmd.ExecuteScalar() == 1)
                        {
                            duplicate = true;
                        }
                    }
                    if (!duplicate)
                    {
                        try
                        {
                            if (dane_Dni.Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach > 0)
                            {
                                ilosc_wpisow++;
                                const string sqlQueryInsertOdbNadgodzin = @"
    DECLARE @PRA_PraId INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @PRI_PraId);
    DECLARE @EXISTSDZIEN DATETIME = (SELECT PracPracaDni.PPR_Data FROM cdn.PracPracaDni WHERE PPR_PraId = @PRA_PraId and PPR_Data = @DataInsert)
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
                        ,[PPR_Zrodlo])
                    VALUES
                        (@PRI_PraId
                        ,@DataInsert
                        ,GETDATE()
                        ,GETDATE()
                        ,'ADMIN'
                        ,'Administrator'
                        ,'ADMIN'
                        ,'Administrator'
                        ,0)
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
		    DATEADD(MINUTE, 0, @GodzOdDate),
		    DATEADD(MINUTE, 0, @GodzDoDate),
		    @TypPracy,
		    1,
		    1,
		    '',
		    @TypNadg);";
                                using (SqlCommand insertCmd = new SqlCommand(sqlQueryInsertOdbNadgodzin, connection, transaction))
                                {
                                    insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.Parse($"{karta.Rok}-{karta.Miesiac:D2}-{dane_Dni.Dzien:D2}"));
                                    insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                    insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                    insertCmd.Parameters.AddWithValue("@PRI_PraId", IdPracownika);
                                    insertCmd.Parameters.AddWithValue("@TypPracy", 2); // podstawowy
                                    insertCmd.Parameters.AddWithValue("@TypNadg", 4); // W.PŁ
                                    insertCmd.Parameters.AddWithValue("@ImieMod", Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20));
                                    insertCmd.Parameters.AddWithValue("@NazwiskoMod", Truncate(Internal_Error_Logger.Last_Mod_Osoba, 50));
                                    insertCmd.Parameters.AddWithValue("@DataMod", Internal_Error_Logger.Last_Mod_Time);
                                    insertCmd.ExecuteScalar();
                                }
                            }
                        }
                        catch (SqlException ex)
                        {
                            transaction.Rollback();
                            Internal_Error_Logger.New_Custom_Error(ex.Message + " z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                            throw new Exception(ex.Message + $" w pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}" + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                        }
                        catch (FormatException)
                        {
                            continue;
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            Internal_Error_Logger.New_Custom_Error(ex.Message + " z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                            throw new Exception(ex.Message + $" w pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}" + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                        }
                    }
                }
            }
            return ilosc_wpisow;
        }
    }
}
