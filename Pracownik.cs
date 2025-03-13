using System.Data;
using Microsoft.Data.SqlClient;

namespace Excel_Data_Importer_WARS
{
    internal class Pracownik
    {
        public string Imie = string.Empty;
        public string Nazwisko = string.Empty;
        public string Akronim = string.Empty;
        public int Get_PraId()
        {
            using SqlCommand command = new(DbManager.Get_PRI_PraId, DbManager.GetConnection(), DbManager.Transaction_Manager.CurrentTransaction);
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
            int Pracid = command.ExecuteScalar() as int? ?? 0;
            return Pracid;
        }
    }
}