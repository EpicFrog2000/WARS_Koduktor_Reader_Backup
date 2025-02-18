using System.Data;
using DocumentFormat.OpenXml.Wordprocessing;
using Excel_Data_Importer_WARS;
using Microsoft.Data.SqlClient;

namespace Konduktor_Reader
{
    internal class Pracownik
    {
        public string Imie = string.Empty;
        public string Nazwisko = string.Empty;
        public string Akronim = string.Empty;

        public int Get_PraId(SqlConnection connection, SqlTransaction transaction)
        {
            using SqlCommand getCmd = new(DbManager.Get_PRI_PraId, connection, transaction);
            if (string.IsNullOrEmpty(Akronim))
            {
                getCmd.Parameters.Add("@Akronim", SqlDbType.Int).Value = -1;
            }
            else
            {
                getCmd.Parameters.Add("@Akronim", SqlDbType.Int).Value = int.Parse(Akronim);
            }
            getCmd.Parameters.Add("@PracownikImieInsert", SqlDbType.NVarChar, 50).Value = Imie;
            getCmd.Parameters.Add("@PracownikNazwiskoInsert", SqlDbType.NVarChar, 50).Value = Nazwisko;
            return getCmd.ExecuteScalar() as int? ?? 0;
        }
    }
}