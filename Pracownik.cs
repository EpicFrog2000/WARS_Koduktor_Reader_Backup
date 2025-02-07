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
            string sqlQueryGetPRI_PraId = @"
-- weź @PRA_PraId z akronimu
IF @Akronim IS NOT NULL AND @Akronim != 0
BEGIN
	DECLARE @AkroRes INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @Akronim);
	IF @AkroRes IS NOT NULL
	BEGIN
		SELECT @AkroRes;
	END
END

-- weż @PRA_PraId z imie i nazwisko
IF (
    (
        SELECT DISTINCT COUNT(PRI_PraId)
        FROM cdn.Pracidx
        WHERE
            (PRI_Imie1 = @PracownikImieInsert AND PRI_Nazwisko = @PracownikNazwiskoInsert AND PRI_Typ = 1)
            OR
            (PRI_Imie1 = @PracownikNazwiskoInsert AND PRI_Nazwisko = @PracownikImieInsert AND PRI_Typ = 1)
    ) > 1
)
BEGIN
    DECLARE @ErrorMessageC NVARCHAR(500) = 'Jest 2 pracowników w bazie o takim samym imieniu i nazwisku, a takiego akronimu nie ma w bazie: ' + @PracownikImieInsert + ' ' + @PracownikNazwiskoInsert + ' ' + Convert(VARCHAR(200), @Akronim);
    THROW 50001, @ErrorMessageC, 1;
END

DECLARE @PRI_PraId INT = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1);
IF @PRI_PraId IS NULL
BEGIN
	SET @PRI_PraId = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikNazwiskoInsert  and PRI_Nazwisko = @PracownikImieInsert and PRI_Typ = 1);
	IF @PRI_PraId IS NULL
	BEGIN
		DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie o imieniu, nazwisku i akronimie: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert + ' ' + Convert(VARCHAR(200), @Akronim);
		THROW 50003, @ErrorMessage, 1;
	END
END

DECLARE @EXISTSPRACTEST INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @PRI_PraId)

IF @EXISTSPRACTEST IS NULL
BEGIN
    INSERT INTO [CDN].[PracKod] ([PRA_Kod] ,[PRA_Archiwalny],[PRA_Nadrzedny],[PRA_EPEmail],[PRA_EPTelefon],[PRA_EPNrPokoju],[PRA_EPDostep],[PRA_HasloDoWydrukow])
    VALUES (@PRI_PraId,0,0,'','','',0,'');
END
SELECT @PRI_PraId;";
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                try
                {
                    connection.Open();
                    using (SqlCommand getCmd = new SqlCommand(sqlQueryGetPRI_PraId, connection))
                    {
                        getCmd.Parameters.AddWithValue("@Akronim ", Akronim);
                        getCmd.Parameters.AddWithValue("@PracownikImieInsert", Imie);
                        getCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", Nazwisko);
                        object result = getCmd.ExecuteScalar();
                        return result != null ? Convert.ToInt32(result) : 0;
                    }
                }
                catch (Exception ex)
                {
                    connection.Close();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                }
            }
        }
        public int Get_PraId(SqlConnection connection, SqlTransaction transaction)
        {
            string sqlQueryGetPRI_PraId = @"
-- weź @PRA_PraId z akronimu
IF @Akronim IS NOT NULL AND @Akronim != 0
BEGIN
	DECLARE @AkroRes INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @Akronim);
	IF @AkroRes IS NOT NULL
	BEGIN
		SELECT @AkroRes;
	END
END

-- weż @PRA_PraId z imie i nazwisko
IF (
    (
        SELECT DISTINCT COUNT(PRI_PraId)
        FROM cdn.Pracidx
        WHERE
            (PRI_Imie1 = @PracownikImieInsert AND PRI_Nazwisko = @PracownikNazwiskoInsert AND PRI_Typ = 1)
            OR
            (PRI_Imie1 = @PracownikNazwiskoInsert AND PRI_Nazwisko = @PracownikImieInsert AND PRI_Typ = 1)
    ) > 1
)
BEGIN
    DECLARE @ErrorMessageC NVARCHAR(500) = 'Jest 2 pracowników w bazie o takim samym imieniu i nazwisku, a takiego akronimu nie ma w bazie: ' + @PracownikImieInsert + ' ' + @PracownikNazwiskoInsert + ' ' + Convert(VARCHAR(200), @Akronim);
    THROW 50001, @ErrorMessageC, 1;
END

DECLARE @PRI_PraId INT = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1);
IF @PRI_PraId IS NULL
BEGIN
	SET @PRI_PraId = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikNazwiskoInsert  and PRI_Nazwisko = @PracownikImieInsert and PRI_Typ = 1);
	IF @PRI_PraId IS NULL
	BEGIN
		DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie o imieniu, nazwisku i akronimie: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert + ' ' + Convert(VARCHAR(200), @Akronim);
		THROW 50003, @ErrorMessage, 1;
	END
END

DECLARE @EXISTSPRACTEST INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @PRI_PraId)

IF @EXISTSPRACTEST IS NULL
BEGIN
    INSERT INTO [CDN].[PracKod] ([PRA_Kod] ,[PRA_Archiwalny],[PRA_Nadrzedny],[PRA_EPEmail],[PRA_EPTelefon],[PRA_EPNrPokoju],[PRA_EPDostep],[PRA_HasloDoWydrukow])
    VALUES (@PRI_PraId,0,0,'','','',0,'');
END
SELECT @PRI_PraId;";
            try
            {
                using (SqlCommand getCmd = new SqlCommand(sqlQueryGetPRI_PraId, connection, transaction))
                {
                    getCmd.Parameters.AddWithValue("@Akronim ", Akronim);
                    getCmd.Parameters.AddWithValue("@PracownikImieInsert", Imie);
                    getCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", Nazwisko);
                    object result = getCmd.ExecuteScalar();
                    return result != null ? Convert.ToInt32(result) : 0;
                }
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
