using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Data.SqlClient;

//Jest przekopiowane praktycznie 1 do 1 ze starego programu dlatego jest taki murzyński kod
// TODO w przyszłości zrobić to lepiej

namespace Excel_Data_Importer_WARS
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
                    if (!DateTime.TryParse(Wartosc, out DateTime data))
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

            public void Set_Miesiac(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return;
                }
                value = value.ToLower().Trim();
                Dictionary<string, int> months = new()
                {
                    {"styczeń", 1}, {"i", 1}, {"luty", 2}, {"ii", 2}, {"marzec", 3}, {"iii", 3}, {"kwiecień", 4}, {"iv", 4},
                    {"maj", 5}, {"v", 5}, {"czerwiec", 6}, {"vi", 6}, {"lipiec", 7}, {"vii", 7}, {"sierpień", 8}, {"viii", 8},
                    {"wrzesień", 9}, {"ix", 9}, {"październik", 10}, {"x", 10}, {"listopad", 11}, {"xi", 11}, {"grudzień", 12}, {"xii", 12}
                };
                Miesiac = months.FirstOrDefault(kvp => value.Contains(kvp.Key)).Value;
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

                List<TimeSpan> newGodziny_Pracy_Od = [];
                List<TimeSpan> newGodziny_Pracy_Do = [];


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
            string Uwaga = Get_Uwaga_Karty(Zakladka);
            foreach (Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika in Karty_Ewidencji_Pracownika)
            {
                Dodaj_Dane_Do_Optimy(Karta_Ewidencji_Pracownika, ref Uwaga);
            }


        }
        private static void Get_Dane_Naglowka_Karty(ref Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika, Helper.Current_Position StartKarty, IXLWorksheet Zakladka)
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
                    if (dane.EndsWith('r'))
                    {
                        dane = dane[..^1].Trim();
                    }
                    if (dane.EndsWith("r."))
                    {
                        dane = dane[..^2].Trim();
                    }

                    string[] dateFormats = ["dd.MM.yyyy"];
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
                                if (dane.Split(" ").Length > 1)
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
            }

            dane = string.Empty;
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
                            dane = dane.Replace("akronim", "").Replace(":", "");
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
                Internal_Error_Logger.New_Error(dane, "Imie nazwisko akronim", StartKarty.Row - 2, StartKarty.Col, "Nieznany format", false);
                throw new Exception($"{ex}, {Internal_Error_Logger.Get_Error_String()}");
            }

        }
        private static Pracownik Get_Pracownik(int sposob, Helper.Current_Position StartKarty, IXLWorksheet Zakladka)
        {
            string[] wordsToRemove = ["IMIĘ:", "IMIE:", "NAZWISKO:", "NAZWISKO", " IMIE", "IMIĘ", ":"];
            string dane;
            Pracownik pracownik = new();
            switch (sposob)
            {
                // case 5:
                // KARTA  PRACY: NAZWISKO IMIĘ ||| Imie | NAzwisko, akronim ;

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
                                if (int.TryParse(parts[0], out _))
                                {
                                    pracownik.Akronim = parts[0];
                                    pracownik.Nazwisko = dane.Trim().Split(' ')[1];
                                    pracownik.Imie = dane.Trim().Split(' ')[2];
                                }
                                else if (int.TryParse(parts[2], out _))
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
        private static void Get_Dane_Dni(ref Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika, Helper.Current_Position Pozycja, IXLWorksheet Zakladka)
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
                        }
                    }
                }

                // absencja
                dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 3).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    Absencja Absencja = new()
                    {
                        Dzien = Dane_Karty.Dzien,
                        Miesiac = Karta_Ewidencji_Pracownika.Miesiac,
                        Rok = Karta_Ewidencji_Pracownika.Rok
                    };
                    if (!Helper.Try_Get_Type_From_String<string>(dane.ToUpper(), ref Absencja.Nazwa))
                    {
                        Internal_Error_Logger.New_Error(dane, "Nazwa Absencji", Pozycja.Col + 3, Pozycja.Row, "Zły format Nazwy absencji");
                    }
                    if (!Absencja.RodzajAbsencji.TryParse(Absencja.Nazwa, out Absencja.Rodzaj_Absencji))
                    {
                        Internal_Error_Logger.New_Error(dane, "Rodzaj Absencji", Pozycja.Col + 3, Pozycja.Row, "Nierozpoznany rodzaj absencji");
                    }

                    dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 7).GetFormattedString().Trim().Replace("  ", " ");
                    
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Absencja.Liczba_Godzin_Przepracowanych))
                    {
                        //Internal_Error_Logger.New_Error(dane, "Liczba godz Absencji", Pozycja.Col + 18, Pozycja.Row + Row_Offset, "Zły format lub bark Liczba godz Absencji");
                        //throw new Exception(Internal_Error_Logger.Get_Error_String());
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
                    }
                }

                // Godziny_Nadliczbowe_Platne_Z_Dodatkiem_50
                dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 9).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Karty.Godziny_Nadliczbowe_Platne_Z_Dodatkiem_50))
                    {
                        Internal_Error_Logger.New_Error(dane, "Godziny Nadliczbowe Platne Z Dodatkiem 50", Pozycja.Col + 9, Pozycja.Row, "Zły format Godziny Nadliczbowe Platne Z Dodatkiem 50");
                    }
                }

                // Godziny_Nadliczbowe_Platne_Z_Dodatkiem_100
                dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col + 10).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    if (!Helper.Try_Get_Type_From_String<decimal>(dane, ref Dane_Karty.Godziny_Nadliczbowe_Platne_Z_Dodatkiem_100))
                    {
                        Internal_Error_Logger.New_Error(dane, "Godziny Nadliczbowe Platne Z Dodatkiem 100", Pozycja.Col + 10, Pozycja.Row, "Zły format Godziny Nadliczbowe Platne Z Dodatkiem 100");
                    }
                }
                Pozycja.Row++;
                Karta_Ewidencji_Pracownika.Dane_Karty.Add(Dane_Karty);
            }
        }
        private static void Dodaj_Dane_Do_Optimy(Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika, ref string Uwaga)
        {
            using (SqlConnection connection = new(DbManager.Connection_String))
            {
                connection.Open();
                using (SqlTransaction transaction = connection.BeginTransaction(System.Data.IsolationLevel.ReadUncommitted))
                {
                    if (Dodaj_Obecnosci_do_Optimy(Karta_Ewidencji_Pracownika, transaction, connection) > 0)
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
                    if (Absencja.Dodaj_Absencje_do_Optimy(Karta_Ewidencji_Pracownika.Absencje, transaction, connection, Karta_Ewidencji_Pracownika.Pracownik, Internal_Error_Logger) > 0)
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
                    if (Dodaj_Godz_Odbior_Do_Optimy(Karta_Ewidencji_Pracownika, transaction, connection) > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Poprawnie dodano odbiory nadgodzin z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Nie dodano żadnych odbiorow nadgodzin");
                        Console.ForegroundColor = ConsoleColor.White;
                    }

                    
                    if (!string.IsNullOrEmpty(Uwaga))
                    {
                        if (Update_Opis_Karty(Karta_Ewidencji_Pracownika, Uwaga, connection, transaction))
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine($"Poprawnie dodano uwagę z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"Nie dodano żadnej uwagi");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        Uwaga = string.Empty; // Dodanie tego rekordu tylko raz
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
                    Zrob_Insert_Obecnosc_Command(connection, transaction, dzien, TimeSpan.Zero, TimeSpan.Zero, Karta_Ewidencji_Pracownika, Helper.Strefa.undefined); // 1 - pusta strefa
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
                            if (Dane_Karty.Godziny_Rozpoczecia_Pracy[j] != Dane_Karty.Godziny_Zakonczenia_Pracy[j])
                            {
                                ilosc_wpisow += Zrob_Insert_Obecnosc_Command(connection, transaction, Data_Karty, Dane_Karty.Godziny_Rozpoczecia_Pracy[j], Dane_Karty.Godziny_Zakonczenia_Pracy[j], Karta_Ewidencji_Pracownika, Helper.Strefa.Czas_Pracy_Podstawowy);
                            }
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
                                if (Dane_Karty.Godziny_Rozpoczecia_Pracy[k] != Dane_Karty.Godziny_Zakonczenia_Pracy[k])
                                {
                                    ilosc_wpisow += Zrob_Insert_Obecnosc_Command(connection, transaction, Data_Karty, Dane_Karty.Godziny_Rozpoczecia_Pracy[k], Dane_Karty.Godziny_Zakonczenia_Pracy[k], Karta_Ewidencji_Pracownika, Helper.Strefa.Czas_Pracy_Podstawowy);
                                }
                            }
                        }
                        else
                        {
                            Zrob_Insert_Obecnosc_Command(connection, transaction, Data_Karty, TimeSpan.Zero, TimeSpan.Zero, Karta_Ewidencji_Pracownika, Helper.Strefa.undefined); // 1 - pusta strefa
                        }
                    }
                }
                
            }
            return ilosc_wpisow;
        }
        private static int Zrob_Insert_Obecnosc_Command(SqlConnection connection, SqlTransaction transaction, DateTime Data_Karty, TimeSpan startPodstawowy, TimeSpan endPodstawowy, Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika, Helper.Strefa Strefa)
        {
            try
            {
                DateTime godzOdDate = DbManager.Base_Date + startPodstawowy;
                DateTime godzDoDate = DbManager.Base_Date + endPodstawowy;
                bool duplicate = false;
                int IdPracownika = -1;
                try
                {
                    IdPracownika = Karta_Ewidencji_Pracownika.Pracownik.Get_PraId(connection, transaction);
                }
                catch (Exception ex)
                {
                    connection.Close();
                    Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                }

                using (SqlCommand command = new(DbManager.Check_Duplicate_Obecnosc, connection, transaction))
                {
                    command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                    command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                    command.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = Data_Karty;
                    command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                    command.Parameters.Add("@Strefa", SqlDbType.Int).Value = Strefa;
                    duplicate = (int)command.ExecuteScalar() == 1;
                }

                if (!duplicate)
                {
                    using (SqlCommand command = new(DbManager.Insert_Obecnosci, connection, transaction))
                    {
                        command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                        command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                        command.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = Data_Karty;
                        command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                        command.Parameters.Add("@Strefa", SqlDbType.Int).Value = Strefa;
                        command.Parameters.Add("@ImieMod", SqlDbType.NVarChar, 20).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20);
                        command.Parameters.Add("@NazwiskoMod", SqlDbType.NVarChar, 50).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 50);
                        command.Parameters.Add("@DataMod", SqlDbType.DateTime).Value = Internal_Error_Logger.Last_Mod_Time;
                        command.ExecuteScalar();
                    }
                    return 1;
                }
            }
            catch (SqlException ex)
            {
                transaction.Rollback();
                Internal_Error_Logger.New_Custom_Error($"Error podczas operacji w bazie(Zrob_Insert_Obecnosc_Command): {ex.Message}");
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                Internal_Error_Logger.New_Custom_Error($"Error: {ex.Message}");
            }
            return 0;
        }
        private static int Dodaj_Godz_Odbior_Do_Optimy(Karta_Ewidencji_Pracownika karta, SqlTransaction transaction, SqlConnection connection)
        {
            int ilosc_wpisow = 0;
            foreach (Dane_Karty dane_Dni in karta.Dane_Karty)
            {
                if(dane_Dni.Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach > 0)
                {
                    int IdPracownika = karta.Pracownik.Get_PraId(connection, transaction);
                    decimal Ilosc_Godzin = dane_Dni.Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach;
                    DateTime godzOdDate = DbManager.Base_Date + TimeSpan.FromHours(8);
                    DateTime godzDoDate = DbManager.Base_Date + TimeSpan.FromHours(8) + TimeSpan.FromHours((double)dane_Dni.Liczba_Godzin_Do_Odbioru_Za_Prace_W_Nadgodzinach);
                    bool duplicate = false;
                    using (SqlCommand command = new(DbManager.Check_Duplicate_Odbior_Nadgodzin, connection, transaction))
                    {
                        command.Parameters.AddWithValue("@PRI_PraId", IdPracownika);
                        command.Parameters.AddWithValue("@Strefa", Helper.Strefa.Czas_Pracy_Podstawowy);
                        command.Parameters.AddWithValue("@Odb_Nadg", Helper.Odb_Nadg.W_PŁ);
                        command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                        command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                        command.Parameters.AddWithValue("@DataInsert", DateTime.Parse($"{karta.Rok}-{karta.Miesiac:D2}-{dane_Dni.Dzien:D2}"));
                        if ((int)command.ExecuteScalar() == 1)
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
                                using (SqlCommand command = new(DbManager.Insert_Odbior_Nadgodzin, connection, transaction))
                                {
                                    command.Parameters.AddWithValue("@DataInsert", DateTime.Parse($"{karta.Rok}-{karta.Miesiac:D2}-{dane_Dni.Dzien:D2}"));
                                    command.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                    command.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                    command.Parameters.AddWithValue("@PRI_PraId", IdPracownika);
                                    command.Parameters.AddWithValue("@Strefa", Helper.Strefa.Czas_Pracy_Podstawowy); 
                                    command.Parameters.AddWithValue("@Odb_Nadg", Helper.Odb_Nadg.W_PŁ);
                                    command.Parameters.AddWithValue("@ImieMod", Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20));
                                    command.Parameters.AddWithValue("@NazwiskoMod", Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 50));
                                    command.Parameters.AddWithValue("@DataMod", Internal_Error_Logger.Last_Mod_Time);
                                    command.ExecuteScalar();
                                }
                            }
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
            }
            return ilosc_wpisow;
        }
        private static string Get_Uwaga_Karty(IXLWorksheet Zakladka)
        {
            List<Helper.Current_Position> Pozycje = Helper.Find_Starting_Points(Zakladka, "Godziny nocne:", false);
            foreach (Helper.Current_Position Pozycja in Pozycje)
            {
                Pozycja.Row+=6;
                string dane = Zakladka.Cell(Pozycja.Row, Pozycja.Col).GetFormattedString().Trim().ToLower();
                if (!string.IsNullOrEmpty(dane))
                {
                    string Uwaga = string.Empty;
                    if (!Helper.Try_Get_Type_From_String<string>(dane, ref Uwaga))
                    {
                        Internal_Error_Logger.New_Error(dane, "Uwaga karty", Pozycja.Col, Pozycja.Row, "Błąd w trakcie wczytywania uwagi karty");
                    }
                    return Uwaga;
                }
            }
            return "";
        }
        private static bool Update_Opis_Karty(Karta_Ewidencji_Pracownika Karta_Ewidencji_Pracownika, string Uwaga, SqlConnection connection, SqlTransaction transaction)
        {
            using (SqlCommand command = new(DbManager.Update_Uwaga, connection, transaction))
            {
                command.Parameters.Add("@Uwaga", SqlDbType.NVarChar, 1024).Value = Helper.Truncate(Uwaga, 1024);
                command.Parameters.Add("@PracId", SqlDbType.Int).Value = Karta_Ewidencji_Pracownika.Pracownik.Get_PraId(connection, transaction);
                command.Parameters.Add("@Data", SqlDbType.DateTime).Value = new DateTime(Karta_Ewidencji_Pracownika.Rok, Karta_Ewidencji_Pracownika.Miesiac, 1);
                try
                {
                    int rowsAffected = command.ExecuteNonQuery();
                    return rowsAffected > 0;
                }
                catch (SqlException ex)
                {
                    transaction.Rollback();
                    Internal_Error_Logger.New_Custom_Error($"Error podczas operacji w bazie(Update_Opis_Karty): {ex.Message}");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    Internal_Error_Logger.New_Custom_Error($"Error(Update_Opis_Karty): {ex.Message}");
                }
            }
            return false;
        }
    }
}
