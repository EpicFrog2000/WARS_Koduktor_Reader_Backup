

using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace Konduktor_Reader
{
    internal static class Helper
    {
        public class Pracownik
        {
            public string Imie = string.Empty;
            public string Nazwisko = string.Empty;
            public string Akronim = string.Empty;
        }
        //if (!IsValidTypeFromString<int>("123", ref int intValue))
        //{
        //    Program.error ...
        //}
        public static bool Try_Get_Type_From_String<T>(string? value, ref T result)
        {
            result = default!;
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }
            try
            {
                if (typeof(T) == typeof(TimeSpan) && TimeSpan.TryParse(value, out var Time_Span_Value))
                {
                    result = (T)(object)Time_Span_Value;
                    return true;
                }
                if (typeof(T) == typeof(DateTime) && DateTime.TryParse(value, out var Date_Time_Value))
                {
                    result = (T)(object)Date_Time_Value;
                    return true;
                }
                var Converted_Value = Convert.ChangeType(value, typeof(T));
                if (Converted_Value != null)
                {
                    result = (T)Converted_Value;
                    return true;
                }
            }
            catch
            {
            }
            return false;
        }
        public static bool Try_Get_Type_From_String<T>(string? value, Action<T> method)
        {
            T refvalue = default!;
            if (Try_Get_Type_From_String<T>(value, ref refvalue))
            {
                method(refvalue);
                return true;
            }
            return false;
        }
        public static string Read_Value_Diffrent_Possible_Cells_In_Row(int Rows, int StartRow, int StartCol, IXLWorksheet Zakladka)
        {
            for (int i = 0; i < Rows; i++)
            {
                string Result = Zakladka.Cell(StartRow, StartCol).GetFormattedString().Trim().Replace("  ", " ");
                if (!string.IsNullOrEmpty(Result)){
                    return Result;
                }
            }
            return string.Empty;
        }
        public static int Get_Pracownik_Id(Pracownik Pracownik)
        {
            string query = @"SELECT Pra_Id FROM Pracownicy WHERE
                            Pra_Akronim = @Akronim OR
                            (Pra_Imie = @Imie AND Pra_Nazwisko = @Nazwisko) OR
                            (Pra_Imie = @Nazwisko AND Pra_Nazwisko = @Imie)";
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Akronim", Pracownik.Akronim);
                    command.Parameters.AddWithValue("@Imie", Pracownik.Imie);
                    command.Parameters.AddWithValue("@Nazwisko", Pracownik.Nazwisko);
                    connection.Open();
                    object result = command.ExecuteScalar();
                    if (result != null)
                    {
                        return Convert.ToInt32(result);
                    }
                    else
                    {
                        Program.error_logger.New_Custom_Error($"Nie ma takiego pracownika w bazie o danych imie: {Pracownik.Imie}, Nazwisko: {Pracownik.Nazwisko}, Akronim: {Pracownik.Akronim}");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }
            }
        }
        public static int Get_Harmonogram_Id(Reader_Harmonogram_v1.Harmonogram Harmonogram, Relacja Relacja)
        {
            string query = @"select H_Id FROM Harmonogramy WHERE H_PraId = @Pra_Id AND H_RId = @Relacja_Id AND H_Rok = @Rok AND H_Miesiac = @Miesiac;";
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    int Relacja_Id = Relacja.Get_Relacja_Id();
                    int Prac_Id = Get_Pracownik_Id(Harmonogram.Konduktor);
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
        public static void Insert_Harmonogram(Reader_Harmonogram_v1.Harmonogram Harmonogram, Relacja Relacja)
        {
            try
            {
                Helper.Get_Harmonogram_Id(Harmonogram, Relacja);
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
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Pra_Id", Get_Pracownik_Id(Harmonogram.Konduktor));
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
        public static void Insert_Harmonogram_Godziny(Reader_Harmonogram_v1.Harmonogram Harmonogram, Relacja Relacja)
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
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                foreach (var relacja in Relacja.Godziny_Pracy)
                {
                    using (SqlCommand command = new SqlCommand(query, connection))
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
        public static void Insert_Harmonogram_Nieobecnosci()
        {
        }
        public static void Insert_Stawka(Relacja Relacja)
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
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                foreach (Reader_Tabela_Stawek_v1.System_Obsługi_Relacji System_Obsługi_Relacji in Relacja.System_Obsługi_Relacji)
                {
                    using (SqlCommand command = new SqlCommand(query, connection))
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
        public static int Get_Praca_Id()
        {
            return -1;
        }
    }
}
