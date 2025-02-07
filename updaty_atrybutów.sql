DECLARE @NazwaAtrybutu VARCHAR(200) = 'Dodatek wyjazdowy';
DECLARE @ATHDataOd DATETIME = '2024-01-01 00:00:00.000';
DECLARE @ATHDataDo DATETIME = '2025-01-01 00:00:00.000';
DECLARE @NowaWartosc NVARCHAR(101) = '1.50';

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
    UPDATE SET 
        ATH_Wartosc = @NowaWartosc,
        ATH_DataOd = @ATHDataOd,
        ATH_DataDo = @ATHDataDo
WHEN NOT MATCHED THEN
    INSERT (ATH_PrcId, ATH_AtkId, ATH_OatId, ATH_Wartosc, ATH_DataOd, ATH_DataDo)
    VALUES (0, 4, source.OAT_OatId, @NowaWartosc, @ATHDataOd, @ATHDataDo);

SELECT * 
FROM cdn.OAtrybutyHist join cdn.OAtrybuty on  OAT_OatId = ATH_OatId where OAT_NazwaKlasy = @NazwaAtrybutu

-- TODO ZROBIÆ ¯EBY TYLKO PRACOWNIKOM Z ODPOWIEDNIM NR RELACJI(TABELE TEZ DO UTWORZENIA)

--ALTER TABLE cdn.PracPracaDni
--ADD PPR_Relacja varchar(20) null default null;