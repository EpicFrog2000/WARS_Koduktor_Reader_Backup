using System.Data;
using DocumentFormat.OpenXml.Wordprocessing;
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
            const string sqlQuery = @"
DECLARE @PRI_PraId INT = NULL;

IF @Akronim IS NOT NULL AND @Akronim != 0
BEGIN
    SELECT @PRI_PraId = PRA_PraId FROM CDN.PracKod WHERE PRA_Kod = @Akronim;
END

IF @PRI_PraId IS NULL
BEGIN
    IF EXISTS (
        SELECT 1 FROM cdn.Pracidx
        WHERE ((PRI_Imie1 = @PracownikImieInsert AND PRI_Nazwisko = @PracownikNazwiskoInsert)
            OR (PRI_Imie1 = @PracownikNazwiskoInsert AND PRI_Nazwisko = @PracownikImieInsert))
        AND PRI_Typ = 1
        HAVING COUNT(PRI_PraId) > 1)
    BEGIN
        DECLARE @ErrorMessageA NVARCHAR(MAX);
        SET @ErrorMEssageA = N'Jest 2 pracownikow o takim imie i nazwisko: ' 
            + ISNULL(CAST(@PracownikNazwiskoInsert AS NVARCHAR(MAX)), N'') + N' ' 
            + ISNULL(CAST(@PracownikImieInsert AS NVARCHAR(MAX)), N'') + N' ' 
            + ISNULL(CAST(@Akronim AS NVARCHAR(MAX)), N'');
        THROW 50001, 
            @ErrorMessageA, 
            1;
    END

    SELECT @PRI_PraId = PRI_PraId
    FROM cdn.Pracidx
    WHERE (PRI_Imie1 = @PracownikImieInsert AND PRI_Nazwisko = @PracownikNazwiskoInsert)
       OR (PRI_Imie1 = @PracownikNazwiskoInsert AND PRI_Nazwisko = @PracownikImieInsert)
    AND PRI_Typ = 1;

    IF @PRI_PraId IS NULL
    BEGIN
        DECLARE @ErrorMessage NVARCHAR(MAX);
        SET @ErrorMessage = N'Brak pracownika o danych: '  
            + ISNULL(CAST(@PracownikNazwiskoInsert AS NVARCHAR(MAX)), N'') + N' '  
            + ISNULL(CAST(@PracownikImieInsert AS NVARCHAR(MAX)), N'') + N' '  
            + ISNULL(CAST(@Akronim AS NVARCHAR(MAX)), N'');

        THROW 50003, @ErrorMessage, 1;
    END
END

IF NOT EXISTS (SELECT 1 FROM CDN.PracKod WHERE PRA_Kod = @PRI_PraId)
BEGIN
    INSERT INTO CDN.PracKod (PRA_Kod, PRA_Archiwalny, PRA_Nadrzedny, PRA_EPEmail, PRA_EPTelefon, PRA_EPNrPokoju, PRA_EPDostep, PRA_HasloDoWydrukow)
    VALUES (@PRI_PraId, 0, 0, '', '', '', 0, '');
END

SELECT @PRI_PraId;";
            using SqlCommand getCmd = new(sqlQuery, connection, transaction);
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