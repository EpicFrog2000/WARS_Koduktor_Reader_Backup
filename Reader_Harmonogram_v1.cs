using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Konduktor_Reader
{
    internal static class Reader_Harmonogram_v1
    {
        public class Godzina_Pracy
        {
            public int Dzien = 0;
            public TimeSpan Godzina_Rozpoczecia_Pracy = TimeSpan.Zero;
            public TimeSpan Godzina_Zakonczenia_Pracy = TimeSpan.Zero;
        }
        public class Harmonogram
        {
            public int Miesiac = 0;
            public int Rok = 0;
            public Helper.Pracownik Konduktor = new();
            public List<Relacja> Relacje = [];
            public void Set_Miesiac(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return;
                }
                value = value.ToLower().Trim().Split(' ')[0];
                switch (value)
                {
                    case "styczeń":
                        Miesiac = 1;
                        break;
                    case "luty":
                        Miesiac = 2;
                        break;
                    case "marzec":
                        Miesiac = 3;
                        break;
                    case "kwiecień":
                        Miesiac = 4;
                        break;
                    case "maj":
                        Miesiac = 5;
                        break;
                    case "czerwiec":
                        Miesiac = 6;
                        break;
                    case "lipiec":
                        Miesiac = 7;
                        break;
                    case "sierpień":
                        Miesiac = 8;
                        break;
                    case "wrzesień":
                        Miesiac = 9;
                        break;
                    case "październik":
                        Miesiac = 10;
                        break;
                    case "listopad":
                        Miesiac = 11;
                        break;
                    case "grudzień":
                        Miesiac = 12;
                        break;
                    default:
                        break;
                }
            }
            public void Set_Rok(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return;
                }
                value = value.ToLower().Trim();
                var parts = value.Split(' ');
                if (parts.Length == 1)
                {
                    value = parts[0];
                }else if(parts.Length == 2)
                {
                    value = parts[1];
                }
                if (int.TryParse(value, out int Parsed_Value))
                {
                    Rok = Parsed_Value;
                }
            }
            public void Set_Date(string value)
            {
                Set_Miesiac(value);
                Set_Rok(value);
            }
        }
        private class Current_Position
        {
            private int col = 1;
            private int row = 1;
            public int Col
            {
                get => col;
                set
                {
                    if (value < 1)
                    {
                        Program.error_logger.New_Custom_Error("Błąd w programie, próba czytania komórki w kolumnie mniejszej niż 1");
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
                        Program.error_logger.New_Custom_Error("Błąd w programie, próba czytania komórki w rzędzie mniejszym niż 1");
                        throw new ArgumentOutOfRangeException(nameof(Col), "Błąd w programie, próba czytania komórki w rzędzie mniejszym niż 1");
                    }
                    row = value;
                }
            }
        }
        public static void Process_Zakladka(IXLWorksheet Zakladka)
        {
            List<Current_Position> Pozcje_Harmonogramów_W_Zakladce = Find_Staring_Points_Harmonogramy(Zakladka);
            List<Harmonogram> Harmonogramy = [];
            foreach (Current_Position pozycja in Pozcje_Harmonogramów_W_Zakladce)
            {
                Harmonogram Harmonogram = new();
                Get_Dane_Naglowka(ref Harmonogram, pozycja, Zakladka);
                pozycja.Row += 4;
                Get_Relacje_Harmonogramu(ref Harmonogram, pozycja, Zakladka);
                Harmonogramy.Add(Harmonogram);
            }

            foreach (Harmonogram Harmonogram in Harmonogramy)
            {
                foreach (Relacja Relacja in Harmonogram.Relacje)
                {
                    Relacja.Insert_Relacja();
                    Helper.Insert_Harmonogram(Harmonogram, Relacja)
                    Helper.Insert_Harmonogram_Godziny(Harmonogram, Relacja);
                }
            }
        }
        private static List<Current_Position> Find_Staring_Points_Harmonogramy(IXLWorksheet Zakladka)
        {
            List<Current_Position> starty = new();
            int Limiter = 1000;
            int counter = 0;
            foreach (IXLCell? cell in Zakladka.CellsUsed())
            {
                try
                {
                    if (cell.HasFormula && !cell.Address.ToString()!.Equals(cell.FormulaA1))
                    {
                        counter++;
                        if (counter > Limiter)
                        {
                            break;
                        }
                        continue;
                    }
                    if (cell.Value.ToString().Contains("Dzień miesiąca"))
                    {
                        starty.Add(new Current_Position()
                        {
                            Row = cell.Address.RowNumber,
                            Col = cell.Address.ColumnNumber
                        });
                    }
                }
                catch
                {
                    continue;
                }
            }
            return starty;
        }
        private static void Get_Dane_Naglowka(ref Harmonogram Harmonogram, Current_Position pozycja, IXLWorksheet Zakladka)
        {
            // Imie Nazwisko
            string dane = Zakladka.Cell(pozycja.Row - 2, pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "Imię Nazwisko", pozycja.Col + 2, pozycja.Row - 2, "Program nie znalazł imienia i nazwiska w tym harmonogramie");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            string[] parts = dane.Split(' ');
            if(parts.Length == 2)
            {
                Harmonogram.Konduktor.Imie = parts[0].Trim();
                Harmonogram.Konduktor.Imie = parts[1].Trim();
            }
            else
            {
                Program.error_logger.New_Error(dane, "Imię Nazwisko", pozycja.Col + 2, pozycja.Row - 2, "Niepoprawny format");
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            // Data harmonogramu
            dane = Zakladka.Cell(pozycja.Row - 1, pozycja.Col + 4).GetFormattedString().Trim().Replace("  ", " ");
            if (!Helper.Try_Get_Type_From_String<string>(dane, Harmonogram.Set_Date))
            {
                Program.error_logger.New_Error(dane, "miesiąc rok", pozycja.Col + 4, pozycja.Row - 1, "Program nie znalazł miesiac rok w tym harmonogramie");
                throw new Exception(Program.error_logger.Get_Error_String());
            }

        }
        private static void Get_Relacje_Harmonogramu(ref Harmonogram Harmonogram, Current_Position pozycja, IXLWorksheet Zakladka)
        {
            int Obecny_Dzien = 0;
            string dane = Zakladka.Cell(pozycja.Row + 1, pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "dzień", pozycja.Col, pozycja.Row + 1, "Program nie znalazł dnia harmonogramu");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            if(!Helper.Try_Get_Type_From_String<int>(dane, ref Obecny_Dzien))
            {
                Program.error_logger.New_Error(dane, "dzień", pozycja.Col, pozycja.Row + 1, "Niepoprawny format dnia");
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            while (Obecny_Dzien <= 31) // max
            {

                dane = Zakladka.Cell(pozycja.Row + Obecny_Dzien, pozycja.Col).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(dane))
                {
                    break;
                }
                dane = Zakladka.Cell(pozycja.Row + Obecny_Dzien, pozycja.Col + 1).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(dane))
                {
                    Relacja Relacja = new();
                    Relacja.Numer_Relacji = dane;
                    dane = Zakladka.Cell(pozycja.Row + Obecny_Dzien, pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                    if (!Helper.Try_Get_Type_From_String<string>(dane, ref Relacja.Opis_Relacji_1))
                    {
                        Program.error_logger.New_Error(dane, "Opis relacji", pozycja.Col + 2, pozycja.Row + Obecny_Dzien, "Brak opisu relacji obok numeru relacji");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }

                    dane = Zakladka.Cell(pozycja.Row + Obecny_Dzien, pozycja.Col + 3).GetFormattedString().Trim().Replace("  ", " ");
                    if (string.IsNullOrEmpty(dane))
                    {
                        Program.error_logger.New_Error(dane, "Godzina rozpoczęcia relacji", pozycja.Col + 3, pozycja.Row + Obecny_Dzien, "Brak Godzina rozpoczęcia relacji relacji obok opisu relacji");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }

                    if(!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref Relacja.Godzina_Rozpoczecia_Relacji))
                    {
                        Program.error_logger.New_Error(dane, "Godzina rozpoczęcia relacji", pozycja.Col + 5, pozycja.Row + Obecny_Dzien, "Godzina rozpoczęcia relacji w nieprawidłowym formacie");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                    Relacja.Dzien_Rozpoczenia_Relacji = Obecny_Dzien;
                    int Dni_Relacji = 1;
                    while (true)
                    {
                        dane = Zakladka.Cell(pozycja.Row + Obecny_Dzien + Dni_Relacji, pozycja.Col + 2).GetFormattedString().Trim().Replace("  ", " ");
                        Dni_Relacji++;
                        if (string.IsNullOrEmpty(dane))
                        {
                            break;
                        }
                    }

                    for (int i = 1; i <= Dni_Relacji; i++)
                    {
                        Obecny_Dzien++;
                        Godzina_Pracy godzina_Pracy = new();
                        dane = Zakladka.Cell(pozycja.Row + Obecny_Dzien, pozycja.Col + 4).GetFormattedString().Trim().Replace("  ", " ");
                        if (!string.IsNullOrEmpty(dane))
                        {
                            if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref godzina_Pracy.Godzina_Rozpoczecia_Pracy))
                            {
                                Program.error_logger.New_Error(dane, "Godzina rozpoczęcia pracy", pozycja.Col + 4, pozycja.Row + Obecny_Dzien, "Godzina rozpoczęcia pracy w nieprawidłowym formacie");
                                throw new Exception(Program.error_logger.Get_Error_String());
                            }
                        }

                        dane = Zakladka.Cell(pozycja.Row + Obecny_Dzien, pozycja.Col + 5).GetFormattedString().Trim().Replace("  ", " ");
                        if (!string.IsNullOrEmpty(dane))
                        {
                            if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref godzina_Pracy.Godzina_Zakonczenia_Pracy))
                            {
                                Program.error_logger.New_Error(dane, "Godzina zakonczenia pracy", pozycja.Col + 5, pozycja.Row + Obecny_Dzien, "Godzina zakonczenia pracy w nieprawidłowym formacie");
                                throw new Exception(Program.error_logger.Get_Error_String());
                            }
                        }
                        godzina_Pracy.Dzien = Obecny_Dzien;
                        Relacja.Godziny_Pracy.Add(godzina_Pracy);
                    }
                    Harmonogram.Relacje.Add(Relacja);
                    continue;
                }
                Obecny_Dzien++;
            }
        }
    }
}
