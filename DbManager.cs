using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using System.Threading;
using Microsoft.Data.SqlClient;

namespace Excel_Data_Importer_WARS
{
    internal static class DbManager
    {
        public static string Connection_String = string.Empty;
        public static readonly string Insert_Odbior_Nadgodzin = @"
    DECLARE @PRA_PraId INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @PRI_PraId);
    DECLARE @EXISTSDZIEN DATETIME = (SELECT PracPracaDni.PPR_Data FROM cdn.PracPracaDni WHERE PPR_PraId = @PRA_PraId and PPR_Data = @DataInsert)
    IF @EXISTSDZIEN is null
    BEGIN
        BEGIN TRY
            INSERT INTO [CDN].[PracPracaDni]
                        ([PPR_PraId]
                        ,[PPR_Data]
                        ,[PPR_TS_Zal]
                        ,[PPR_TS_Mod]
                        ,[PPR_OpeModKod]
                        ,[PPR_OpeModNazwisko]
                        ,[PPR_OpeZalKod]
                        ,[PPR_OpeZalNazwisko]
                        ,[PPR_Zrodlo])
                    VALUES
                        (@PRI_PraId
                        ,@DataInsert
                        ,GETDATE()
                        ,GETDATE()
                        ,'ADMIN'
                        ,'Administrator'
                        ,'ADMIN'
                        ,'Administrator'
                        ,0)
        END TRY
        BEGIN CATCH
        END CATCH
    END

    INSERT INTO CDN.PracPracaDniGodz
		    (PGR_PprId,
		    PGR_Lp,
		    PGR_OdGodziny,
		    PGR_DoGodziny,
		    PGR_Strefa,
		    PGR_DzlId,
		    PGR_PrjId,
		    PGR_Uwagi,
		    PGR_OdbNadg)
	    VALUES
		    ((select PPR_PprId from cdn.PracPracaDni where CAST(PPR_Data as datetime) = @DataInsert and PPR_PraId = @PRI_PraId),
		    1,
		    DATEADD(MINUTE, 0, @GodzOdDate),
		    DATEADD(MINUTE, 0, @GodzDoDate),
		    @Strefa,
		    1,
		    1,
		    '',
		    @Odb_Nadg);";
        public static readonly string Check_Duplicate_Odbior_Nadgodzin = @"
    DECLARE @EXISTSDZIEN INT;
    DECLARE @EXISTSDATA INT;
    SET @EXISTSDZIEN = (SELECT COUNT(PPR_Data) FROM cdn.PracPracaDni WHERE PPR_PraId = @PRI_PraId AND PPR_Data = @DataInsert);
    SET @EXISTSDATA = (
        SELECT COUNT(*)
        FROM CDN.PracPracaDniGodz 
        WHERE PGR_OdbNadg = 4
            AND PGR_Strefa = 2
            AND PGR_OdGodziny = DATEADD(MINUTE, 0, @GodzOdDate)
            AND PGR_DoGodziny = DATEADD(MINUTE, 0, @GodzDoDate)
            AND PGR_PprId = (SELECT PPR_PprId FROM cdn.PracPracaDni WHERE CAST(PPR_Data AS datetime) = @DataInsert AND PPR_PraId = @PRI_PraId)
    );
    SELECT CASE 
        WHEN @EXISTSDZIEN > 0 AND @EXISTSDATA > 0 THEN 1
        ELSE 0
    END;";
        public static readonly string Check_Duplicate_Nieobecnosci = @"SELECT CASE 
    WHEN EXISTS (
        SELECT 1
        FROM [CDN].[PracNieobec] PNB
        WHERE PNB.PNB_PraId = @PRI_PraId
        AND PNB.PNB_OkresOd = @DataOd
        AND PNB.PNB_OkresDo = @DataDo
    ) THEN 1
    ELSE 0
END AS ExistsFlag;
";
        public static readonly string Insert_Nieobecnosci = @$"
DECLARE @TNBID INT = (SELECT TNB_TnbId FROM cdn.TypNieobec WHERE TNB_Nazwa = @NazwaNieobecnosci);
    INSERT INTO [CDN].[PracNieobec]
               ([PNB_PraId]
               ,[PNB_TnbId]
               ,[PNB_TyuId]
               ,[PNB_NaPodstPoprzNB]
               ,[PNB_OkresOd]
               ,[PNB_Seria]
               ,[PNB_Numer]
               ,[PNB_OkresDo]
               ,[PNB_Opis]
               ,[PNB_Rozliczona]
               ,[PNB_RozliczData]
               ,[PNB_ZwolFPFGSP]
               ,[PNB_UrlopNaZadanie]
               ,[PNB_Przyczyna]
               ,[PNB_DniPracy]
               ,[PNB_DniKalend]
               ,[PNB_Calodzienna]
               ,[PNB_ZlecZasilekPIT]
               ,[PNB_PracaRodzic]
               ,[PNB_Dziecko]
               ,[PNB_OpeZalId]
               ,[PNB_StaZalId]
               ,[PNB_TS_Zal]
               ,[PNB_TS_Mod]
               ,[PNB_OpeModKod]
               ,[PNB_OpeModNazwisko]
               ,[PNB_OpeZalKod]
               ,[PNB_OpeZalNazwisko]
               ,[PNB_Zrodlo])
         VALUES
               (@PRI_PraId
               ,@TNBID
               ,99999
               ,0
               ,@DataOd
               ,''
               ,''
               ,@DataDo
               ,''
               ,0
               ,@BaseDate
               ,0
               ,0
               ,@Przyczyna
               ,@DniPracy
               ,@DniKalendarzowe
               ,1
               ,0
               ,0
               ,''
               ,1
               ,1
               ,@DataMod
               ,@DataMod
               ,@ImieMod
               ,@NazwiskoMod
               ,@ImieMod
               ,@NazwiskoMod
               ,0);";
        public static readonly string Check_Duplicate_Obecnosc = @"
        IF EXISTS (
            SELECT 1
            FROM cdn.PracPracaDni P
            INNER JOIN CDN.PracPracaDniGodz G ON P.PPR_PprId = G.PGR_PprId
            WHERE P.PPR_PraId = @PRI_PraId 
              AND P.PPR_Data = @DataInsert
              AND G.PGR_OdGodziny = @GodzOdDate
              AND G.PGR_DoGodziny = @GodzDoDate
              AND G.PGR_Strefa = @Strefa
        )
        BEGIN
            SELECT 1;
        END
        ELSE
        BEGIN
            SELECT 0;
        END";
        public static readonly string Insert_Obecnosci = @"
DECLARE @EXISTSDZIEN DATETIME = (SELECT PracPracaDni.PPR_Data FROM cdn.PracPracaDni WITH (NOLOCK) WHERE PPR_PraId = @PRI_PraId and PPR_Data = @DataInsert)
IF @EXISTSDZIEN is null
BEGIN
    BEGIN TRY
        INSERT INTO [CDN].[PracPracaDni]
                    ([PPR_PraId]
                    ,[PPR_Data]
                    ,[PPR_TS_Zal]
                    ,[PPR_TS_Mod]
                    ,[PPR_OpeModKod]
                    ,[PPR_OpeModNazwisko]
                    ,[PPR_OpeZalKod]
                    ,[PPR_OpeZalNazwisko]
                    ,[PPR_Zrodlo])
                VALUES
                    (@PRI_PraId
                    ,@DataInsert
                    ,@DataMod
                    ,@DataMod
                    ,@ImieMod
                    ,@NazwiskoMod
                    ,@ImieMod
                    ,@NazwiskoMod
                    ,0)
    END TRY
    BEGIN CATCH
    END CATCH
END

INSERT INTO CDN.PracPracaDniGodz
		(PGR_PprId,
		PGR_Lp,
		PGR_OdGodziny,
		PGR_DoGodziny,
		PGR_Strefa,
		PGR_DzlId,
		PGR_PrjId,
		PGR_Uwagi,
		PGR_OdbNadg)
	VALUES
		((select PPR_PprId from cdn.PracPracaDni where CAST(PPR_Data as datetime) = @DataInsert and PPR_PraId = @PRI_PraId),
		1,
		@GodzOdDate,
		@GodzDoDate,
		@Strefa,
		1,
		1,
		'',
		1);
";
        public static readonly string Insert_Atrybuty = @$"
                            WITH CTE AS (
                                SELECT OAT_OatId
            FROM cdn.OAtrybuty
            WHERE OAT_AtkId = (SELECT ATK_AtkId FROM cdn.OAtrybutyKlasy WHERE ATK_Nazwa = @NazwaAtrybutu)
                            )

                            MERGE cdn.OAtrybutyHist AS target
                            USING CTE AS source
                            ON target.ATH_OatId = source.OAT_OatId
                               AND target.ATH_DataOd = @ATHDataOd
                               AND target.ATH_DataDo = @ATHDataDo
                            WHEN MATCHED THEN
                                UPDATE SET ATH_Wartosc = @NowaWartosc
                            WHEN NOT MATCHED THEN
                                INSERT (ATH_PrcId, ATH_AtkId, ATH_OatId, ATH_Wartosc, ATH_DataOd, ATH_DataDo)
                                VALUES (0, 4, source.OAT_OatId, @NowaWartosc, @ATHDataOd, @ATHDataDo);";
        public static readonly string Insert_Relacja = @"INSERT INTO CDN.Relacje
            (R_Nazwa
            --,R_Typ
            ,R_Opis_1
            ,R_Opis_2
            ,R_Godz_Rozpoczecia
            ,R_Data_Mod
            ,R_Os_Mod)
        VALUES
            (@Nazwa_Relacji
            --,@R_Typ
            ,@Opis_1
            ,@Opis_2
            ,@Godz_Rozpoczecia
            ,@Data_Mod
            ,@Os_Mod)";
        public static readonly string Get_Relacja = @"select R_Id from cdn.Relacje where R_Nazwa = @R_Nazwa AND R_Typ = @R_Typ;";
        public static readonly string Insert_Prowizje = @"WITH CTE AS (
    SELECT OAT_OatId 
            FROM cdn.OAtrybuty 
            WHERE OAT_AtkId = (SELECT ATK_AtkId FROM cdn.OAtrybutyKlasy WHERE ATK_Nazwa = @NazwaAtrybutu) and
			OAT_PrcId = @PracID
)

MERGE cdn.OAtrybutyHist AS target
USING CTE AS source
ON target.ATH_OatId = source.OAT_OatId
   AND target.ATH_DataOd = @ATHDataOd
   AND target.ATH_DataDo = @ATHDataDo
WHEN MATCHED THEN
    UPDATE SET ATH_Wartosc = @NowaWartosc
WHEN NOT MATCHED THEN
    INSERT (ATH_PrcId, ATH_AtkId, ATH_OatId, ATH_Wartosc, ATH_DataOd, ATH_DataDo)
    VALUES (0, 4, source.OAT_OatId, @NowaWartosc, @ATHDataOd, @ATHDataDo);";
        public static readonly string Insert_Plan_Pracy = @"
DECLARE @id int;
DECLARE @EXISTSDZIEN INT = (SELECT COUNT([CDN].[PracPlanDni].[PPL_Data]) FROM cdn.PracPlanDni WHERE cdn.PracPlanDni.PPL_PraId = @PRI_PraId and [CDN].[PracPlanDni].[PPL_Data] = @DataInsert)
IF @EXISTSDZIEN = 0
BEGIN
BEGIN TRY
INSERT INTO [CDN].[PracPlanDni]
        ([PPL_PraId]
        ,[PPL_Data]
        ,[PPL_TS_Zal]
        ,[PPL_TS_Mod]
        ,[PPL_OpeModKod]
        ,[PPL_OpeModNazwisko]
        ,[PPL_OpeZalKod]
        ,[PPL_OpeZalNazwisko]
        ,[PPL_Zrodlo]
        ,[PPL_TypDnia])
VALUES
        (@PRI_PraId
        ,@DataInsert
        ,@DataMod
        ,@DataMod
        ,@ImieMod
        ,@NazwiskoMod
        ,@ImieMod
        ,@NazwiskoMod
        ,0
        ,ISNULL((SELECT TOP 1 KAD_TypDnia FROM cdn.KalendDni WHERE KAD_Data = @DataInsert), 1))
END TRY
BEGIN CATCH
END CATCH
END

SET @id = (select [cdn].[PracPlanDni].[PPL_PplId] from [cdn].[PracPlanDni] where [cdn].[PracPlanDni].[PPL_Data] = @DataInsert and [cdn].[PracPlanDni].[PPL_PraId] = @PRI_PraId);
INSERT INTO CDN.PracPlanDniGodz
	        (PGL_PplId,
	        PGL_Lp,
	        PGL_OdGodziny,
	        PGL_DoGodziny,
	        PGL_Strefa,
	        PGL_DzlId,
	        PGL_PrjId,
	        PGL_UwagiPlanu)
        VALUES
	        (@id,
	        1,
	        @GodzOdDate,
	        @GodzDoDate,
	        @Strefa,
	        1,
	        1,
	        '');";
        public static readonly string Check_Duplicate_Plan_Pracy = @"
IF EXISTS (
SELECT 1 
FROM cdn.PracPlanDni 
WHERE PPL_Data = @DataInsert 
    AND PPL_PraId = @PRI_PraId
)
BEGIN
IF EXISTS (
    SELECT 1 
    FROM cdn.PracPlanDniGodz 
    WHERE PGL_PplId = (
        SELECT PPL_PplId 
        FROM cdn.PracPlanDni 
        WHERE PPL_Data = @DataInsert 
            AND PPL_PraId = @PRI_PraId
    )
        AND PGL_OdGodziny = @GodzOdDate 
        AND PGL_DoGodziny = @GodzDoDate
)
BEGIN
    SELECT 1;
END
ELSE
BEGIN
    SELECT 0;
END
END
ELSE
BEGIN
SELECT 0;
END";
        public static readonly string Get_PRI_PraId = @"
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
        public static readonly string Update_Uwaga = @"UPDATE pdg
SET pdg.PGR_Uwagi = @Uwaga
FROM cdn.PracPracaDniGodz pdg
INNER JOIN cdn.PracPracaDni ppd 
    ON pdg.PGR_PprId = ppd.PPR_PprId
WHERE ppd.PPR_PraId = @PracId
AND ppd.PPR_Data = @Data;
";
        public static readonly string Get_Id_Dni_Godz_Pracy = @"select PGR_PgrId from cdn.PracPracaDniGodz where PGR_PprId in (select PPR_PprId from cdn.PracPracaDni where PPR_Data = @DataInsert and PPR_PraId = @PRI_PraId);";
        public static readonly string Update_Dzien_Pracy_Strefa = @"IF EXISTS (
    SELECT 1 
    FROM cdn.PracPracaDniGodz 
    WHERE PGR_PgrId = @IdDniaGodz 
    AND PGR_OdGodziny = '1899-12-30 00:00:00' 
    AND PGR_DoGodziny = '1899-12-30 00:00:00'
)
BEGIN
    UPDATE cdn.PracPracaDniGodz 
    SET PGR_Strefa = @NowaStrefa,
	PGR_OdGodziny = @NewOdGodz,
	PGR_DoGodziny = @NewDoGodz
    WHERE PGR_PgrId = @IdDniaGodz;
END
ELSE
BEGIN
    UPDATE cdn.PracPracaDniGodz 
    SET PGR_Strefa = @NowaStrefa
    WHERE PGR_PgrId = @IdDniaGodz;
END
";
        public static bool Valid_SQLConnection_String()
        {
            try
            {
                using (SqlConnection connection = new(Connection_String))
                {
                    connection.Open();
                    connection.Close();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
        public static readonly DateTime Base_Date = new(1899, 12, 30); // Do zapytan sql (zostawić z powodów historycznych xdd, tak powstało pół godzinki)
    }
}
