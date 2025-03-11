using System.Data;
using System.Transactions;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace Excel_Data_Importer_WARS
{


    internal class Absencja
    {
        public int Dzien = 0;
        public int Miesiac = 0;
        public int Rok = 0;
        public string Nazwa = string.Empty;
        public decimal Liczba_Godzin_Absencji = 0;
        public RodzajAbsencji Rodzaj_Absencji = RodzajAbsencji.DEFAULT;
        public decimal Liczba_Godzin_Przepracowanych = 0;
        public enum RodzajAbsencji
        {
            DEFAULT = -1,
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
            UŻ,     // Urlop na żądanie
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
            SZK,    // Szkolenie
        }
        public static int Dodaj_Absencje_do_Optimy(List<Absencja> Absencje, SqlTransaction transaction, SqlConnection connection, Pracownik Pracownik, Error_Logger Internal_Error_Logger)
        {
            int ilosc_wpisow = 0;

            foreach(Absencja absencja in Absencje)
            {
                if (absencja.Rodzaj_Absencji == RodzajAbsencji.DE)
                {

                    int IdPrac = -1;
                    try
                    {
                        IdPrac = Pracownik.Get_PraId(connection, transaction);
                    }
                    catch (Exception ex)
                    {
                        connection.Close();
                        Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                    }

                    //sprawdz czy jest juz delegacja w tym dniu




                    //get id dni
                    List<int> Lista_Dni_Godz_Pracy = [];
                    using (SqlCommand command = new(DbManager.Get_Id_Dni_Godz_Pracy, connection, transaction))
                    {
                        command.Parameters.Add("@DataInsert", SqlDbType.DateTime).Value = new DateTime(absencja.Rok, absencja.Miesiac, absencja.Dzien);
                        command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPrac;
                        object result = command.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                        {
                            if (result is int singleValue)
                            {
                                Lista_Dni_Godz_Pracy.Add(singleValue);
                            }
                            else if (result is int[] array)
                            {
                                Lista_Dni_Godz_Pracy.AddRange(array);
                            }
                        }
                    }
                    foreach (int Dzien_Godz in Lista_Dni_Godz_Pracy)
                    {
                        //update strefe dnia na delegacje
                        try
                        {
                            using (SqlCommand command = new(DbManager.Update_Dzien_Pracy_Strefa, connection, transaction))
                            {
                                


                                command.Parameters.Add("@NowaStrefa", SqlDbType.Int).Value = Helper.Strefa.Czas_Pracy_W_Delegacji;
                                command.Parameters.Add("@IdDniaGodz", SqlDbType.Int).Value = Dzien_Godz;
                                command.Parameters.Add("@NewOdGodz", SqlDbType.DateTime).Value = DbManager.Base_Date + TimeSpan.FromHours(8);
                                command.Parameters.Add("@NewDoGodz", SqlDbType.DateTime).Value = DbManager.Base_Date + TimeSpan.FromHours(8) + TimeSpan.FromHours((double)absencja.Liczba_Godzin_Przepracowanych);
                                command.ExecuteScalar();
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
            }




            List<List<Absencja>> ListyAbsencji = Podziel_Absencje_Na_Osobne(Absencje);
            foreach (List<Absencja> ListaAbsencji in ListyAbsencji)
            {
                if (ListaAbsencji[0].Rodzaj_Absencji == RodzajAbsencji.DEFAULT)
                {
                    continue;
                }

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

                if (ListaAbsencji[0].Rodzaj_Absencji == RodzajAbsencji.ON)
                {
                    continue;
                }

                if (ListaAbsencji[0].Rodzaj_Absencji == RodzajAbsencji.DE)
                {
                    continue;
                }

                



                // Pierdole to, w zależnpści od nazwy absencji są rozne przyczyny które nie są opisane w dokumentacji, nie mam sposobu na wgl zrobienie tego w 
                //  sposób który nie zmiażdży mi jąder :(.
                //int przyczyna = Dopasuj_TKN_Nazwa(ListaAbsencji[0].Rodzaj_Absencji);
                int przyczyna = 1; // Nie dotyczy

                string nazwa_absencji = Dopasuj_TBN_Nazwa(ListaAbsencji[0].Rodzaj_Absencji);
                if (string.IsNullOrEmpty(nazwa_absencji))
                {
                    Internal_Error_Logger.New_Custom_Error($"W programie brak dopasowanego kodu Absencji: {ListaAbsencji[0].Rodzaj_Absencji} w dniu {new DateTime(ListaAbsencji[0].Rok, ListaAbsencji[0].Miesiac, ListaAbsencji[0].Dzien)} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki}. Absencja nie dodana.", false);
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
                    IdPracownika = Pracownik.Get_PraId(connection, transaction);
                }
                catch (Exception ex)
                {
                    connection.Close();
                    Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                }

                using (SqlCommand command = new(DbManager.Check_Duplicate_Nieobecnosci, connection, transaction))
                {
                    command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                    command.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = Data_Absencji_Start;
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
                        using (SqlCommand command = new(DbManager.Insert_Nieobecnosci, connection, transaction))
                        {
                            command.Parameters.Add("@PRI_PraId", SqlDbType.Int).Value = IdPracownika;
                            command.Parameters.Add("@NazwaNieobecnosci", SqlDbType.NVarChar, 40).Value = nazwa_absencji;
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

                        transaction.Rollback();
                        Internal_Error_Logger.New_Custom_Error($"{ex.Message} z pliku: {Internal_Error_Logger.Nazwa_Pliku} z zakladki: {Internal_Error_Logger.Nr_Zakladki} nazwa zakladki: {Internal_Error_Logger.Nazwa_Zakladki}");
                    }
                    catch
                    {
                        transaction.Rollback();
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
                if (currentGroup.Count == 0 ||
                    (Absencja.Dzien == currentGroup[^1].Dzien + 1 && Absencja.Rodzaj_Absencji == currentGroup[^1].Rodzaj_Absencji))
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
            //1 – Nie dotyczy
            //2 – Zwolnienie lekarskie
            //3 – Wypadek w pracy / choroba zawodowa
            //4 – Wypadek w drodze do/ z pracy
            //5 – Zwolnienie w okresie ciąży
            //6 – Zwolnienie spowodowane gruźlicą
            //7 – Nadużycie alkoholu
            //8 – Przestępstwa / wykroczenie
            //9 – Opieka nad dzieckiem do lat 14
            //10 – Opieka nad inną osobą
            //11 – Leczenie szpitalne
            //12 - Badanie dawcy / pobranie organów
            return rodzaj switch
            {
                RodzajAbsencji.DM => 9,
                RodzajAbsencji.DR => 9,
                RodzajAbsencji.NB => 2,
                RodzajAbsencji.NR => 2,
                RodzajAbsencji.U9 => 9,
                RodzajAbsencji.UC => 9,
                RodzajAbsencji.UD => 9,
                RodzajAbsencji.UK => 12,
                RodzajAbsencji.UM => 9,
                RodzajAbsencji.UR => 11,
                RodzajAbsencji.ZC => 10,
                RodzajAbsencji.ZD => 9,
                RodzajAbsencji.ZK => 9,
                RodzajAbsencji.ZL => 2,
                RodzajAbsencji.ZN => 2,
                RodzajAbsencji.ZR => 11,
                RodzajAbsencji.ZS => 11,
                RodzajAbsencji.ZZ => 5,
                _ => 1
            };
        }
        private static string Dopasuj_TBN_Nazwa(RodzajAbsencji rodzaj)
        {
            return rodzaj switch
            {
                RodzajAbsencji.UO => "Urlop okolicznościowy",
                RodzajAbsencji.ZL => "Zwolnienie chorobowe",
                RodzajAbsencji.ZY => "Zwolnienie chorobowe",
                RodzajAbsencji.ZS => "Zwolnienie chorobowe",
                RodzajAbsencji.ZN => "Zwolnienie chorobowe.",
                RodzajAbsencji.ZP => "Zwolnienie chorobowe",
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
                RodzajAbsencji.UZ => "Urlop inny",
                RodzajAbsencji.UŻ => "Urlop inny",
                //RodzajAbsencji.ZZ => "" Kurwa niewiem POMOCY xd
                //RodzajAbsencji.UD => "", NIE MA opieki nad dzieckiem 
                _ => "Nieobecność (B2B)"
            };
        }
    }
}