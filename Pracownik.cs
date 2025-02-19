using System.Data;
using Microsoft.Data.SqlClient;

namespace Excel_Data_Importer_WARS
{
    internal class Pracownik
    {
        public string Imie = string.Empty;
        public string Nazwisko = string.Empty;
        public string Akronim = string.Empty;
        public int Get_PraId(SqlConnection connection, SqlTransaction transaction)
        {
            using SqlCommand command = new(DbManager.Get_PRI_PraId, connection, transaction);
            if (string.IsNullOrEmpty(Akronim))
            {
                command.Parameters.Add("@Akronim", SqlDbType.Int).Value = -1;
            }
            else
            {
                command.Parameters.Add("@Akronim", SqlDbType.Int).Value = int.Parse(Akronim);
            }
            command.Parameters.Add("@PracownikImieInsert", SqlDbType.NVarChar, 50).Value = Imie;
            command.Parameters.Add("@PracownikNazwiskoInsert", SqlDbType.NVarChar, 50).Value = Nazwisko;
            return command.ExecuteScalar() as int? ?? 0;
        }
    }
}