using System.Data;
using Microsoft.Data.SqlClient;
using static Excel_Data_Importer_WARS.Reader_Tabela_Stawek_v1;

namespace Excel_Data_Importer_WARS
{
    internal class Relacja
    {
        public string Numer_Relacji = string.Empty;
        public string Opis_Relacji_1 = string.Empty;
        public string Opis_Relacji_2 = string.Empty;
        public TimeSpan Godzina_Rozpoczecia_Relacji = TimeSpan.Zero;
        public int Dzien_Rozpoczenia_Relacji = 0;
        public List<System_Obsługi_Relacji> System_Obsługi_Relacji = [];

        public static int Get_Relacja_Id(string Numer_Relacji, SqlConnection connection, SqlTransaction transaction)
        {
            using (SqlCommand command = new(DbManager.Get_Relacja, connection, transaction))
            {
                command.Parameters.Add("@R_Nazwa", SqlDbType.NVarChar, 20).Value = Numer_Relacji;
                command.Parameters.Add("@R_Typ", SqlDbType.Int).Value = DBNull.Value;
                object result = command.ExecuteScalar();
                if (result != null)
                {
                    return Convert.ToInt32(result);
                }
                else
                {
                    throw new Exception($"Nie ma takiej relacji w bazie o danych Numer_Relacji: {Numer_Relacji}");
                }
            }
        }

        public void Insert_Relacja_Do_Optimy(Error_Logger Internal_Error_Logger, SqlConnection connection, SqlTransaction transaction)
        {
            try
            {
                Get_Relacja_Id(Numer_Relacji, connection, transaction);
            }
            catch
            {
                using (SqlCommand command = new(DbManager.Insert_Relacja, connection, transaction))
                {
                    command.Parameters.Add("@Nazwa_Relacji", SqlDbType.NVarChar, 20).Value = Numer_Relacji;
                    //command.Parameters.Add("@R_Typ", SqlDbType.Int).Value = null;
                    command.Parameters.Add("@Opis_1", SqlDbType.NVarChar, 200).Value = Opis_Relacji_1;
                    command.Parameters.Add("@Opis_2", SqlDbType.NVarChar, 200).Value = Opis_Relacji_2;
                    command.Parameters.Add("@Godz_Rozpoczecia", SqlDbType.DateTime).Value = DbManager.Base_Date + Godzina_Rozpoczecia_Relacji;
                    command.Parameters.Add("@Data_Mod", SqlDbType.DateTime).Value = DateTime.Now;
                    command.Parameters.Add("@Os_Mod", SqlDbType.NVarChar, 20).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20);
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}