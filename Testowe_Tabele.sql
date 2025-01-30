CREATE TABLE Relacje (
    R_Id INT IDENTITY(1,1),
    R_Nazwa VARCHAR(50),
	R_Typ INT,
    R_Opis_1 VARCHAR(100),
	R_Opis_2 VARCHAR(100),
	R_Godz_Rozpoczecia DATETIME,
	R_Data_Mod VARCHAR(100),
	R_Os_Mod VARCHAR(100)
);

CREATE TABLE Stawki (
    S_Id INT IDENTITY(1,1),
    S_RId INT,
	S_Calkowity DECIMAL,
	S_Ogolem DECIMAL,
	S_Podstawowe DECIMAL,
	S_Godz_Nadliczbowe_50 DECIMAL,
	S_Godz_Nadliczbowe_100 DECIMAL,
	S_Czas_Odpoczynku DECIMAL,
	S_Podstawowa_Stawka_Godzinowa DECIMAL,
	S_Wynagrodzenie_Ryczaltowe_Podstawowe DECIMAL,
	S_Wynagrodzenie_Ryczaltowe_Za_Godz_Nadlicznowe DECIMAL,
	S_Wynagrodzenie_Ryczaltowe_Dodatek_Za_Prace_W_Nocy DECIMAL,
	S_Wynagrodzenie_Ryczaltowe_Calkowite DECIMAL,
	S_Dodatek_Wyjazdowy DECIMAL,
	S_Data_Mod VARCHAR(100),
	S_Os_Mod VARCHAR(100)
);

CREATE TABLE Harmonogramy(
	H_Id INT IDENTITY(1,1),
	H_PId INT,
	H_Miesiac INT,
	H_Rok INT,
	H_RId INT,
	H_Opis_1 VARCHAR(100),
	H_Opis_2 VARCHAR(100),
	H_Data_Mod VARCHAR(100),
	H_Os_Mod VARCHAR(100)
);

CREATE TABLE Harmonogramy_Godziny(
	HG_Id INT IDENTITY(1,1),
	HG_HId INT,
	HG_Dzien INT,
	HG_Godzina_Rozpoczecia_Pracy DATETIME,
	HG_Godzina_Zakonczenia_Pracy DATETIME,
	HG_Data_Mod VARCHAR(100),
	HG_Os_Mod VARCHAR(100)
);

CREATE TABLE Harmonogramy_Nieobecnosci(
	HN_Id INT IDENTITY(1,1),
	HN_HId INT,
	HN_Kod VARCHAR(50),
	HN_Opis VARCHAR(200),
	HN_Data_Mod VARCHAR(100),
	HN_Os_Mod VARCHAR(100)
);

CREATE TABLE Praca(
	P_Id INT IDENTITY(1,1),
	P_PraId INT,
	P_Miesiac INT,
	P_Rok INT,
	P_RId INT,
	P_Opis_1 VARCHAR(100),
	P_Opis_2 VARCHAR(100),
	P_Data_Mod VARCHAR(100),
	P_Os_Mod VARCHAR(100)
);

CREATE TABLE Praca_Godziny(
	PG_Id INT IDENTITY(1,1),
	PG_Opis_1 VARCHAR(100),
	PG_Opis_2 VARCHAR(100),
	PG_Godzina_Rozpoczecia_Pracy DATETIME,
	PG_Godzina_Zakonczenia_Pracy DATETIME,
	PG_Godzina_Odpoczynku_Od DATETIME,
	PG_Godzina_Odpoczynku_Do DATETIME,
	PG_Data_Mod VARCHAR(100),
	PG_Os_Mod VARCHAR(100)
);

CREATE TABLE Pracownicy(
	Pra_Id INT IDENTITY(1,1),
	Pra_Imie VARCHAR(100),
	Pra_Nazwisko VARCHAR(100),
	Pra_Akronim VARCHAR(50),
	Pra_Data_Mod VARCHAR(100),
	Pra_Os_Mod VARCHAR(100)
);

DECLARE @Nazwisko VARCHAR(100) = 'Tasarz';
DECLARE @Imie VARCHAR(100) = 'Norbert';
DECLARE @Akronim VARCHAR(50) = '32154';


SELECT Pra_Id FROM Pracownicy WHERE
Pra_Akronim = @Akronim OR
(Pra_Imie = @Imie AND Pra_Nazwisko = @Nazwisko) OR
(Pra_Imie = @Nazwisko AND Pra_Nazwisko = @Imie)