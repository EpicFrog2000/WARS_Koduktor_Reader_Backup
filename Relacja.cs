using System.Data;
using Excel_Data_Importer_WARS;
using Microsoft.Data.SqlClient;
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

        public int Get_Relacja_Id_From_Optima(string Numer_Relacji)
        {
            using (SqlConnection connection = new(DbManager.Connection_String))
            {
                using (SqlCommand command = new(DbManager.Get_Relacja, connection))
                {
                    command.Parameters.Add("@R_Nazwa", SqlDbType.NVarChar, 20).Value = Numer_Relacji;
                    command.Parameters.Add("@R_Typ", SqlDbType.Int, 20).Value = DBNull.Value;
                    connection.Open();
                    object result = command.ExecuteScalar();
                    if (result != null)
                    {
                        return Convert.ToInt32(result);
                    }
                    else
                    {
                        //Program.error_logger.New_Custom_Error($"Nie ma takiej relacji w bazie o danych Numer_Relacji: {Numer_Relacji}");
                        throw new Exception($"Nie ma takiej relacji w bazie o danych Numer_Relacji: {Numer_Relacji}");
                    }
                }
            }
        }

        public void Insert_Relacja_Do_Optimy()
        {
            try
            {
                Get_Relacja_Id_From_Optima(Numer_Relacji);
            }
            catch
            {
                using (SqlConnection connection = new(DbManager.Connection_String))
                {
                    using (SqlCommand command = new(DbManager.Insert_Relacja, connection))
                    {
                        command.Parameters.Add("@Nazwa_Relacji", SqlDbType.NVarChar, 20).Value = Numer_Relacji;
                        //command.Parameters.Add("@R_Typ", SqlDbType.Int).Value = null;
                        command.Parameters.Add("@Opis_1", SqlDbType.NVarChar, 200).Value = Opis_Relacji_1;
                        command.Parameters.Add("@Opis_2", SqlDbType.NVarChar, 200).Value = Opis_Relacji_2;
                        command.Parameters.Add("@Godz_Rozpoczecia", SqlDbType.DateTime, 20).Value = DbManager.Base_Date + Godzina_Rozpoczecia_Relacji;
                        command.Parameters.Add("@Data_Mod", SqlDbType.DateTime, 20).Value = DateTime.Now;
                        command.Parameters.Add("@Os_Mod", SqlDbType.NVarChar, 20).Value = "Norbert Tasarz";
                        connection.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}