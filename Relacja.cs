using Microsoft.Data.SqlClient;
using static Konduktor_Reader.Reader_Harmonogram_v1;
using static Konduktor_Reader.Reader_Tabela_Stawek_v1;

namespace Konduktor_Reader
{
    internal class Relacja
    {
        public string Numer_Relacji = string.Empty;
        public string Opis_Relacji_1 = string.Empty;
        public string Opis_Relacji_2 = string.Empty;
        public TimeSpan Godzina_Rozpoczecia_Relacji = TimeSpan.Zero;
        public int Dzien_Rozpoczenia_Relacji = 0;
        public List<System_Obsługi_Relacji> System_Obsługi_Relacji = [];
        public List<Godzina_Pracy> Godziny_Pracy = [];
        public int Get_Relacja_Id()
        {
            string query = @"select R_Id from Relacje where  R_Nazwa = @R_Nazwa AND R_Typ = @R_Typ;";
            using (SqlConnection connection = new(Program.Optima_Conection_String))
            {
                using (SqlCommand command = new(query, connection))
                {
                    command.Parameters.AddWithValue("@R_Nazwa", Numer_Relacji);
                    command.Parameters.AddWithValue("@R_Typ", "");
                    connection.Open();
                    object result = command.ExecuteScalar();
                    if (result != null)
                    {
                        return Convert.ToInt32(result);
                    }
                    else
                    {
                        Program.error_logger.New_Custom_Error($"Nie ma takiej relacji w bazie o danych Numer_Relacji: {Numer_Relacji}");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }
            }
        }
        public void Insert_Relacja()
        {
            try
            {
                Get_Relacja_Id();
            }
            catch
            {
                string query = @"INSERT INTO Relacje
           (R_Nazwa
           ,R_Typ
           ,R_Opis_1
           ,R_Opis_2
           ,R_Godz_Rozpoczecia
           ,R_Data_Mod
           ,R_Os_Mod)
     VALUES
           (@Nazwa_Relacji
           ,@Typ_Relacji
           ,@Opis_1
           ,@Opis_2
           ,@Godz_Rozpoczecia
           ,@Data_Mod
           ,@Os_Mod)";
                using (SqlConnection connection = new(Program.Optima_Conection_String))
                {
                    using (SqlCommand command = new(query, connection))
                    {
                        command.Parameters.AddWithValue("@R_Nazwa", Numer_Relacji);
                        command.Parameters.AddWithValue("@R_Typ", "");
                        command.Parameters.AddWithValue("@Opis_1", Opis_Relacji_1);
                        command.Parameters.AddWithValue("@Opis_2", Opis_Relacji_2);
                        command.Parameters.AddWithValue("@Godz_Rozpoczecia", Godzina_Rozpoczecia_Relacji);
                        command.Parameters.AddWithValue("@Data_Mod", DateTime.Now);
                        command.Parameters.AddWithValue("@Os_Mod", "Norbert Tasarz");
                        connection.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }

        }
    }
}
