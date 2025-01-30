using Microsoft.Data.SqlClient;

namespace Konduktor_Reader
{
    internal class Pracownik
    {
        public string Imie = string.Empty;
        public string Nazwisko = string.Empty;
        public string Akronim = string.Empty;
        public int Get_Pracownik_Id()
        {
            string query = @"SELECT Pra_Id FROM Pracownicy WHERE
                            Pra_Akronim = @Akronim OR
                            (Pra_Imie = @Imie AND Pra_Nazwisko = @Nazwisko) OR
                            (Pra_Imie = @Nazwisko AND Pra_Nazwisko = @Imie)";
            using (SqlConnection connection = new(Program.Optima_Conection_String))
            {
                using (SqlCommand command = new(query, connection))
                {
                    command.Parameters.AddWithValue("@Akronim", Akronim);
                    command.Parameters.AddWithValue("@Imie", Imie);
                    command.Parameters.AddWithValue("@Nazwisko", Nazwisko);
                    connection.Open();
                    object result = command.ExecuteScalar();
                    if (result != null)
                    {
                        return Convert.ToInt32(result);
                    }
                    else
                    {
                        Program.error_logger.New_Custom_Error($"Nie ma takiego pracownika w bazie o danych imie: {Imie}, Nazwisko: {Nazwisko}, Akronim: {Akronim}");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }
            }
        }
    }
}
