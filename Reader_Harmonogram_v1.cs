using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace Konduktor_Reader
{
    internal static class Reader_Harmonogram_v1
    {
        public class Godzina_Pracy
        {
            public  int Dzien = 0;
            public TimeSpan Godzina_Rozpoczecia_Pracy = TimeSpan.Zero;
            public TimeSpan Godzina_Zakonczenia_Pracy = TimeSpan.Zero;
        }
        private class Harmonogram
        {
            public int Miesiac = 0;
            public int Rok = 0;
            public Pracownik Konduktor = new();
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
        public static void Process_Zakladka(IXLWorksheet Zakladka)
        {
            List<Helper.Current_Position> Pozcje_Harmonogramów_W_Zakladce = Find_Staring_Points_Harmonogramy(Zakladka);
            List<Harmonogram> Harmonogramy = [];
            foreach (Helper.Current_Position pozycja in Pozcje_Harmonogramów_W_Zakladce)
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
                    Insert_Harmonogram(Harmonogram, Relacja);
                    Insert_Harmonogram_Godziny(Harmonogram, Relacja);
                }
            }
        }
        private static List<Helper.Current_Position> Find_Staring_Points_Harmonogramy(IXLWorksheet Zakladka)
        {
            List<Helper.Current_Position> starty = [];
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
                        starty.Add(new Helper.Current_Position()
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
        private static void Get_Dane_Naglowka(ref Harmonogram Harmonogram, Helper.Current_Position pozycja, IXLWorksheet Zakladka)
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
        private static void Get_Relacje_Harmonogramu(ref Harmonogram Harmonogram, Helper.Current_Position pozycja, IXLWorksheet Zakladka)
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
                    Relacja Relacja = new()
                    {
                        Numer_Relacji = dane
                    };
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
        private static int Get_Harmonogram_Id(Harmonogram Harmonogram, Relacja Relacja)
        {
            string query = @"select H_Id FROM Harmonogramy WHERE H_PraId = @Pra_Id AND H_RId = @Relacja_Id AND H_Rok = @Rok AND H_Miesiac = @Miesiac;";
            using (SqlConnection connection = new(Program.Optima_Conection_String))
            {
                using (SqlCommand command = new(query, connection))
                {
                    int Relacja_Id = Relacja.Get_Relacja_Id();
                    int Prac_Id = Harmonogram.Konduktor.Get_Pracownik_Id();
                    command.Parameters.AddWithValue("@Pra_Id", Relacja_Id);
                    command.Parameters.AddWithValue("@Relacja_Id", Prac_Id);
                    command.Parameters.AddWithValue("@Rok", Harmonogram.Rok);
                    command.Parameters.AddWithValue("@Miesiac", Harmonogram.Miesiac);
                    connection.Open();
                    object result = command.ExecuteScalar();
                    if (result != null)
                    {
                        return Convert.ToInt32(result);
                    }
                    else
                    {
                        Program.error_logger.New_Custom_Error($"Nie ma takiego harmonogramu w bazie o danych Rok: {Harmonogram.Rok}, Miesiac: {Harmonogram.Miesiac}, Relacja_Id: {Relacja_Id}, Prac_Id: {Prac_Id}");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }
            }
        }
        private static void Insert_Harmonogram(Harmonogram Harmonogram, Relacja Relacja)
        {
            try
            {
                Get_Harmonogram_Id(Harmonogram, Relacja);
                return;
            }
            catch { }

            string query = @"INSERT INTO Harmonogramy
           (H_PraId
           ,H_Miesiac
           ,H_Rok
           ,H_RId
           ,H_Opis_1
           ,H_Opis_2
           ,H_Data_Mod
           ,H_Os_Mod)
     VALUES
           (@Pra_Id
           ,@Miesiac
           ,@Rok
           ,@Relacja_Id
           ,@Opis_1
           ,@Opis_2
           ,@Data_Mod
           ,@Os_Mod)";
            using (SqlConnection connection = new(Program.Optima_Conection_String))
            {
                using (SqlCommand command = new(query, connection))
                {
                    command.Parameters.AddWithValue("@Pra_Id", Harmonogram.Konduktor.Get_Pracownik_Id());
                    command.Parameters.AddWithValue("@Miesiac", Harmonogram.Miesiac);
                    command.Parameters.AddWithValue("@Rok", Harmonogram.Rok);
                    command.Parameters.AddWithValue("@Rok", Relacja.Get_Relacja_Id());
                    command.Parameters.AddWithValue("@Opis_1", "");
                    command.Parameters.AddWithValue("@Opis_2", "");
                    command.Parameters.AddWithValue("@Data_Mod", DateTime.Now);
                    command.Parameters.AddWithValue("@Os_Mod", "Norbert Tasarz");
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }
        private static void Insert_Harmonogram_Godziny(Harmonogram Harmonogram, Relacja Relacja)
        {
            string query = @"INSERT INTO Harmonogramy_Godziny
           (HG_HId
           ,HG_Dzien
           ,HG_Godzina_Rozpoczecia_Pracy
           ,HG_Godzina_Zakonczenia_Pracy
           ,HG_Data_Mod
           ,HG_Os_Mod)
     VALUES
           (@Harmonogram_Id
           ,@Dzien
           ,@Godzina_Rozpoczecia_Pracy
           ,@Godzina_Zakonczenia_Pracy
           ,@Data_Mod
           ,@Os_Mod)";
            using (SqlConnection connection = new(Program.Optima_Conection_String))
            {
                foreach (var relacja in Relacja.Godziny_Pracy)
                {
                    using (SqlCommand command = new(query, connection))
                    {
                        command.Parameters.AddWithValue("@Harmonogram_Id", Get_Harmonogram_Id(Harmonogram, Relacja));
                        command.Parameters.AddWithValue("@Dzien", relacja.Dzien);
                        command.Parameters.AddWithValue("@Godzina_Rozpoczecia_Pracy", relacja.Godzina_Rozpoczecia_Pracy);
                        command.Parameters.AddWithValue("@Godzina_Zakonczenia_Pracy", relacja.Godzina_Zakonczenia_Pracy);
                        command.Parameters.AddWithValue("@Data_Mod", DateTime.Now);
                        command.Parameters.AddWithValue("@Os_Mod", "Norbert Tasarz");
                        connection.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
        private static void Insert_Harmonogram_Nieobecnosci()
        {
        }
    }
}
