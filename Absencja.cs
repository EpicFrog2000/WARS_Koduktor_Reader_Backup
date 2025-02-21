using System.Data;
using Microsoft.Data.SqlClient;

namespace Excel_Data_Importer_WARS
{


    internal class Absencja
    {
        public int Dzien = 0;
        public int Miesiac = 0;
        public int Rok = 0;
        public string Nazwa = string.Empty;
        public decimal Liczba_Godzin_Absencji = 0;
        public RodzajAbsencji Rodzaj_Absencji = 0;
        public enum RodzajAbsencji
        {
            DE,     // Delegacja
            DM,     // Dodatkowy urlop macierzyński
            DR,     // Urlop rodzicielski
            IK,     // Izolacja - Koronawirus
            NB,     // Badania lekarskie - okresowe
            NN,     // Nieobecność nieusprawiedliwiona
            NR,     // Badania lekarskie - z tyt. niepełnosprawności
            NU,     // Nieobecność usprawiedliwiona
            OD,     // Oddelegowanie do prac w ZZ
            OG,     // Odbiór godzin dyżuru
            ON,     // Odbiór nadgodzin
            OO,     // Odbiór pracy w niedziele
            OP,     // Urlop opiekuńczy (niepłatny)
            OS,     // Odbiór pracujących sobót
            PP,     // Poszukiwanie pracy
            PZ,     // Praca zdalna okazjonalna
            SW,     // Urlop/zwolnienie z tyt. siły wyższej
            SZ,     // Szkolenie
            SP,     // Zwolniony z obowiązku świadcz. pracy
            U9,     // Urlop rodzicielski 9 tygodni
            UA,     // Długotrwały urlop bezpłatny
            UB,     // Urlop bezpłatny
            UC,     // Urlop ojcowski
            UD,     // Na opiekę nad dzieckiem art.K.P.188
            UJ,     // Ćwiczenia wojskowe
            UK,     // Urlop dla krwiodawcy
            UL,     // Służba wojskowa
            ULawnika, // Praca ławnika w sądzie
            UM,     // Urlop macierzyński
            UN,     // Urlop z tyt. niepełnosprawności
            UO,     // Urlop okolicznościowy
            UP,     // Dodatkowy urlop osoby represjonowanej
            UR,     // Dodatkowe dni na turnus rehabilitacyjny
            US,     // Urlop szkoleniowy
            UV,     // Urlop weterana
            UW,     // Urlop wypoczynkowy
            UY,     // Urlop wychowawczy
            UZ,     // Urlop na żądanie
            WY,     // Wypoczynek skazanego
            ZC,     // Opieka nad członkiem rodziny (ZLA)
            ZD,     // Opieka nad dzieckiem (ZUS ZLA)
            ZK,     // Opieka nad dzieckiem Koronawirus
            ZL,     // Zwolnienie lekarskie (ZUS ZLA)
            ZN,     // Zwolnienie lekarskie niepłatne (ZLA)
            ZP,     // Kwarantanna sanepid
            ZR,     // Zwolnienie na rehabilitację (ZUS ZLA)
            ZS,     // Zwolnienie szpitalne (ZUS ZLA)
            ZY,     // Zwolnienie powypadkoe (ZUS ZLA)
            ZZ,     // Zwolnienie lek. (ciąża) (ZUS ZLA)
            SZK     // Szkolenie
        }
        public static int Dodaj_Absencje_do_Optimy(List<Absencja> Absencje, SqlTransaction tran, SqlConnection connection, Pracownik Pracownik, Error_Logger Internal_Error_Logger)
        {
            int ilosc_wpisow = 0;
            List<List<Absencja>> ListyAbsencji = Podziel_Absencje_Na_Osobne(Absencje);
            foreach (List<Absencja> ListaAbsencji in ListyAbsencji)
            {
                DateTime Data_Absencji_Start;
                DateTime Data_Absencji_End;

                try
                {
                    Data_Absencji_Start = new DateTime(ListaAbsencji[0].Rok, ListaAbsencji[0].Miesiac, ListaAbsencji[0].Dzien);
                    Data_Absencji_End = new DateTime(ListaAbsencji[^1].Rok, ListaAbsencji[^1].Miesiac, ListaAbsencji[^1].Dzien);
                }
                catch
                {
                    continue;
                }

                int przyczyna = Dopasuj_TKN_Nazwa(ListaAbsencji[0].Rodzaj_Absencji);
                string nazwa_absencji = Dopasuj_TBN_Nazwa(ListaAbsencji[0].Rodzaj_Absencji);

                if (string.IsNullOrEmpty(nazwa_absencji))
                {
                    Internal_Error_Logger.New_Custom_Error($"W programie brak dopasowanego kodu Absencji: {ListaAbsencji[0].Rodzaj_Absencji} w dniu {new DateTime(ListaAbsencji[0].Rok, ListaAbsencji[0].Miesiac, ListaAbsencji[0].Dzien)} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki}. Absencja nie dodana.");
                    Exception e = new();
                    e.Data["Kod"] = 42069;
                    throw e;
                }
                int dni_robocze = Ile_Dni_Roboczych(ListaAbsencji);
                int dni_calosc = ListaAbsencji.Count;

                bool duplicate = false;

                int IdPracownika = -1;
                try
                {
                    IdPracownika = Pracownik.Get_PraId(connection, tran);
                }
                catch (Exception ex)
                {
                    connection.Close();
                    Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                    throw new Exception($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                }

                using (SqlCommand command = new(DbManager.Check_Duplicate_Absencje, connection, tran))
                {
                    command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                    command.Parameters.Add("@NazwaNieobecnosci", SqlDbType.NVarChar, 50).Value = nazwa_absencji;
                    command.Parameters.Add("@DniPracy", SqlDbType.Int).Value = dni_robocze;
                    command.Parameters.Add("@DniKalendarzowe", SqlDbType.Int).Value = dni_calosc;
                    command.Parameters.Add("@Przyczyna", SqlDbType.NVarChar, 50).Value = przyczyna;
                    command.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = Data_Absencji_Start;
                    command.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = DbManager.Base_Date;
                    command.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = Data_Absencji_End;
                    if ((int)command.ExecuteScalar() == 1)
                    {
                        duplicate = true;
                        return 0;
                    }
                }

                if (!duplicate)
                {
                    try
                    {
                        using (SqlCommand command = new(DbManager.Insert_Nieobecnosci, connection, tran))
                        {
                            command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                            command.Parameters.Add("@NazwaNieobecnosci", SqlDbType.NVarChar, 50).Value = nazwa_absencji;
                            command.Parameters.Add("@DniPracy", SqlDbType.Int).Value = dni_robocze;
                            command.Parameters.Add("@DniKalendarzowe", SqlDbType.Int).Value = dni_calosc;
                            command.Parameters.Add("@Przyczyna", SqlDbType.NVarChar, 50).Value = przyczyna;
                            command.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = Data_Absencji_Start;
                            command.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = DbManager.Base_Date;
                            command.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = Data_Absencji_End;
                            command.Parameters.Add("@ImieMod", SqlDbType.NVarChar, 20).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20);
                            command.Parameters.Add("@NazwiskoMod", SqlDbType.NVarChar, 50).Value = Helper.Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20);
                            command.Parameters.Add("@DataMod", SqlDbType.DateTime).Value = Internal_Error_Logger.Last_Mod_Time;
                            command.ExecuteScalar();
                        }
                    }

                    catch (FormatException ex)
                    {
                        Internal_Error_Logger.New_Custom_Error($"{ex.Message}");

                        continue;
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                    ilosc_wpisow++;
                }
            }
            return ilosc_wpisow;
        }
        private static List<List<Absencja>> Podziel_Absencje_Na_Osobne(List<Absencja> Absencje)
        {
            List<List<Absencja>> OsobneAbsencje = [];
            List<Absencja> currentGroup = [];

            foreach (Absencja Absencja in Absencje)
            {
                if (currentGroup.Count == 0 || Absencja.Dzien == currentGroup[^1].Dzien + 1)
                {
                    currentGroup.Add(Absencja);
                }
                else
                {
                    OsobneAbsencje.Add(new List<Absencja>(currentGroup));
                    currentGroup = [Absencja];
                }
            }

            if (currentGroup.Count > 0)
            {
                OsobneAbsencje.Add(currentGroup);
            }

            return OsobneAbsencje;
        }
        private static int Ile_Dni_Roboczych(List<Absencja> Absencje)
        {
            return Absencje.Count(Absencja =>
            {
                DateTime absenceDate = new(Absencja.Rok, Absencja.Miesiac, Absencja.Dzien);
                return absenceDate.DayOfWeek != DayOfWeek.Saturday && absenceDate.DayOfWeek != DayOfWeek.Sunday;
            });
        }
        private static int Dopasuj_TKN_Nazwa(RodzajAbsencji rodzaj)
        {
            return rodzaj switch
            {
                RodzajAbsencji.ZL => 1,       // Zwolnienie lekarskie
                RodzajAbsencji.DM => 2,       // Urlop macierzyński
                RodzajAbsencji.DR => 13,      // Urlop opiekuńczy
                RodzajAbsencji.NB => 1,       // Zwolnienie lekarskie
                RodzajAbsencji.NN => 5,       // Nieobecność nieusprawiedliwiona
                RodzajAbsencji.UC => 21,      // Urlop opiekuńczy
                RodzajAbsencji.UD => 21,      // Urlop opiekuńczy
                RodzajAbsencji.UJ => 10,      // Służba wojskowa
                RodzajAbsencji.UL => 10,      // Służba wojskowa
                RodzajAbsencji.UM => 2,       // Urlop macierzyński
                RodzajAbsencji.UO => 4,       // Urlop okolicznościowy
                RodzajAbsencji.UN => 3,       // Urlop rehabilitacyjny
                RodzajAbsencji.UR => 3,       // Urlop rehabilitacyjny
                RodzajAbsencji.ZC => 21,      // Urlop opiekuńczy
                RodzajAbsencji.ZD => 21,      // Urlop opiekuńczy
                RodzajAbsencji.ZK => 21,      // Urlop opiekuńczy
                RodzajAbsencji.ZN => 1,       // Zwolnienie lekarskie
                RodzajAbsencji.ZR => 3,       // Urlop rehabilitacyjny
                RodzajAbsencji.ZZ => 1,       // Zwolnienie lekarskie ciążowe
                RodzajAbsencji.SZK => 20,     // SZ
                _ => 9                        // Nie dotyczy dla pozostałych przypadków
            };
        }
        private static string Dopasuj_TBN_Nazwa(RodzajAbsencji rodzaj)
        {
            return rodzaj switch
            {
                RodzajAbsencji.UO => "Urlop okolicznościowy",
                RodzajAbsencji.ZL => "Zwolnienie chorobowe/F",
                RodzajAbsencji.ZY => "Zwolnienie chorobowe/wyp.w drodze/F",
                RodzajAbsencji.ZS => "Zwolnienie chorobowe/wyp.przy pracy/F",
                RodzajAbsencji.ZN => "Zwolnienie chorobowe/bez prawa do zas.",
                RodzajAbsencji.ZP => "Zwolnienie chorobowe/pozbawiony prawa",
                RodzajAbsencji.UR => "Urlop rehabilitacyjny",
                RodzajAbsencji.ZR => "Urlop rehabilitacyjny/wypadek w drodze..",
                RodzajAbsencji.ZD => "Urlop rehabilitacyjny/wypadek przy pracy",
                RodzajAbsencji.UM => "Urlop macierzyński",
                RodzajAbsencji.UC => "Urlop ojcowski",
                RodzajAbsencji.OP => "Urlop opiekuńczy (zasiłek)",
                RodzajAbsencji.UY => "Urlop wychowawczy (121)",
                RodzajAbsencji.UW => "Urlop wypoczynkowy",
                RodzajAbsencji.NU => "Nieobecność usprawiedliwiona (151)",
                RodzajAbsencji.NN => "Nieobecność nieusprawiedliwiona (152)",
                RodzajAbsencji.UL => "Służba wojskowa",
                RodzajAbsencji.DR => "Urlop rodzicielski",
                RodzajAbsencji.DM => "Urlop macierzyński dodatkowy",
                RodzajAbsencji.PP => "Dni wolne na poszukiwanie pracy",
                RodzajAbsencji.UK => "Dni wolne z tyt. krwiodawstwa",
                RodzajAbsencji.IK => "Covid19",
                RodzajAbsencji.SZK => "Szkolenie",
                RodzajAbsencji.PZ => "Praca zdalna",
                //RodzajAbsencji.ZZ => "" Kurwa niewiem POMOCY xd
                //RodzajAbsencji.UD => "", NIE MA opieki nad dzieckiem 
                _ => "Nieobecność (B2B)"
            };
        }
    }
}
