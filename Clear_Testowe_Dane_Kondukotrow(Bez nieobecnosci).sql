DELETE FROM cdn.PracPracaDniGodz 
WHERE PGR_PprId IN (SELECT DISTINCT PPR_PprId 
    FROM cdn.PracPracaDni 
    WHERE PPR_PraId IN (
        125, 257, 579, 1117, 1190, 1534, 
        1567, 1608, 2100, 3008, 3403, 3429, 3573
    ));

DELETE FROM cdn.PracPracaDni 
    WHERE PPR_PraId IN (
        125, 257, 579, 1117, 1190, 1534, 
        1567, 1608, 2100, 3008, 3403, 3429, 3573
    );

delete from cdn.OAtrybutyHist where ATH_AtkId IN (select ATK_AtkId from cdn.OAtrybutyKlasy where ATK_Nazwa IN (
'Wynagrodzenie rycza³towe - Podstawowe',
'Wynagrodzenie rycza³towe - Nadgodziny',
'Wynagrodzenie rycza³towe - Nocki',
'Dodatek wyjazdowy',
'Prowizja za towar',
'Prowizja za wydane napoje awaryjne'
));