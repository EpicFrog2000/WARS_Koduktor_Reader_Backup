using System.Transactions;
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
        public int Get_PraId()
        {
            const string sqlQuery = @"
-- Get PRA_PraId from Akronim if available
DECLARE @PRI_PraId INT = NULL;

IF @Akronim IS NOT NULL AND @Akronim != 0
BEGIN
    SELECT @PRI_PraId = PRA_PraId FROM CDN.PracKod WHERE PRA_Kod = @Akronim;
END

-- If Akronim lookup fails, try by Name & Surname
IF @PRI_PraId IS NULL
BEGIN
    IF EXISTS (
        SELECT 1 FROM cdn.Pracidx 
        WHERE ((PRI_Imie1 = @PracownikImieInsert AND PRI_Nazwisko = @PracownikNazwiskoInsert) 
            OR (PRI_Imie1 = @PracownikNazwiskoInsert AND PRI_Nazwisko = @PracownikImieInsert)) 
        AND PRI_Typ = 1 
        HAVING COUNT(PRI_PraId) > 1)
    BEGIN
        THROW 50001, 'Duplicate employees found for the given name & surname without a unique acronym.', 1;
    END

    SELECT @PRI_PraId = PRI_PraId 
    FROM cdn.Pracidx 
    WHERE (PRI_Imie1 = @PracownikImieInsert AND PRI_Nazwisko = @PracownikNazwiskoInsert) 
       OR (PRI_Imie1 = @PracownikNazwiskoInsert AND PRI_Nazwisko = @PracownikImieInsert) 
    AND PRI_Typ = 1;

    -- If still null, throw an error
    IF @PRI_PraId IS NULL
    BEGIN
        THROW 50003, 'No employee found with the provided details.', 1;
    END
END

-- Ensure PRA_PraId exists in CDN.PracKod
IF NOT EXISTS (SELECT 1 FROM CDN.PracKod WHERE PRA_Kod = @PRI_PraId)
BEGIN
    INSERT INTO CDN.PracKod (PRA_Kod, PRA_Archiwalny, PRA_Nadrzedny, PRA_EPEmail, PRA_EPTelefon, PRA_EPNrPokoju, PRA_EPDostep, PRA_HasloDoWydrukow)
    VALUES (@PRI_PraId, 0, 0, '', '', '', 0, '');
END

SELECT @PRI_PraId;";
            try
            {
                using SqlConnection connection = new();
                using SqlCommand getCmd = new(sqlQuery, connection);
                getCmd.Parameters.AddWithValue("@Akronim", Akronim);
                getCmd.Parameters.AddWithValue("@PracownikImieInsert", Imie);
                getCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", Nazwisko);
                return getCmd.ExecuteScalar() as int? ?? 0;
            }
            catch (Exception ex)
            {
                Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
            }
        }
        public int Get_PraId(SqlConnection connection, SqlTransaction transaction)
        {
            const string sqlQuery = @"
-- Get PRA_PraId from Akronim if available
DECLARE @PRI_PraId INT = NULL;

IF @Akronim IS NOT NULL AND @Akronim != 0
BEGIN
    SELECT @PRI_PraId = PRA_PraId FROM CDN.PracKod WHERE PRA_Kod = @Akronim;
END

-- If Akronim lookup fails, try by Name & Surname
IF @PRI_PraId IS NULL
BEGIN
    IF EXISTS (
        SELECT 1 FROM cdn.Pracidx 
        WHERE ((PRI_Imie1 = @PracownikImieInsert AND PRI_Nazwisko = @PracownikNazwiskoInsert) 
            OR (PRI_Imie1 = @PracownikNazwiskoInsert AND PRI_Nazwisko = @PracownikImieInsert)) 
        AND PRI_Typ = 1 
        HAVING COUNT(PRI_PraId) > 1)
    BEGIN
        THROW 50001, 'Duplicate employees found for the given name & surname without a unique acronym.', 1;
    END

    SELECT @PRI_PraId = PRI_PraId 
    FROM cdn.Pracidx 
    WHERE (PRI_Imie1 = @PracownikImieInsert AND PRI_Nazwisko = @PracownikNazwiskoInsert) 
       OR (PRI_Imie1 = @PracownikNazwiskoInsert AND PRI_Nazwisko = @PracownikImieInsert) 
    AND PRI_Typ = 1;

    -- If still null, throw an error
    IF @PRI_PraId IS NULL
    BEGIN
        THROW 50003, 'No employee found with the provided details.', 1;
    END
END

-- Ensure PRA_PraId exists in CDN.PracKod
IF NOT EXISTS (SELECT 1 FROM CDN.PracKod WHERE PRA_Kod = @PRI_PraId)
BEGIN
    INSERT INTO CDN.PracKod (PRA_Kod, PRA_Archiwalny, PRA_Nadrzedny, PRA_EPEmail, PRA_EPTelefon, PRA_EPNrPokoju, PRA_EPDostep, PRA_HasloDoWydrukow)
    VALUES (@PRI_PraId, 0, 0, '', '', '', 0, '');
END

SELECT @PRI_PraId;";
            try
            {
                using var getCmd = new SqlCommand(sqlQuery, connection, transaction);
                getCmd.Parameters.AddWithValue("@Akronim", Akronim);
                getCmd.Parameters.AddWithValue("@PracownikImieInsert", Imie);
                getCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", Nazwisko);
                return getCmd.ExecuteScalar() as int? ?? 0;
            }
            catch (Exception ex)
            {
                connection.Close();
                Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
            }
        }
    }
}
