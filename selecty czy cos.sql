select PRI_Imie1, PRI_Nazwisko, PPR_Data, PPR_Relacja, OAT_NazwaKlasy, ATH_Wartosc from cdn.PracPracaDni
left join cdn.PracKod on PPR_PraId = PRA_PraId
left join cdn.OAtrybuty on PRA_PraId = OAT_PrcId
left join cdn.OAtrybutyHist on OAT_OatId = ATH_OatId
left join cdn.Pracidx on  PRA_PraId = PRI_PraId 
where PRI_Typ = 1

delete from cdn.PracPracaDni
delete from cdn.PracPracaDniGodz
delete from cdn.Relacje
delete from cdn.OAtrybutyHist
