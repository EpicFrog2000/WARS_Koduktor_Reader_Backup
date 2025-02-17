using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Konduktor_Reader;
using Microsoft.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using static Konduktor_Reader.Helper;

namespace Excel_Data_Importer_WARS
{
    internal class Reader_Grafik_Pracy_Pracownika_2025_v3
    {
        private class Grafik
        {
            public Pracownik Pracownik  = new();
            public int Miesiac  = 0;
            public int Rok  = 0;
            public List<Dane_Dnia> Dane_Dni  = new();
            public string Nazwa_Pliku = "";
            public int Nr_Zakladki = 1;
            public void Set_Miesiac(string wartosc)
            {
                if (string.IsNullOrEmpty(wartosc))
                {
                    Miesiac = -1;
                    return;
                }
                wartosc = wartosc.Trim().ToLower();
                if (wartosc.Contains("styczeń"))
                {
                    Miesiac = 1;
                }
                else if (wartosc.Contains("luty"))
                {
                    Miesiac = 2;
                }
                else if (wartosc.Contains("marzec"))
                {
                    Miesiac = 3;
                }
                else if (wartosc.Contains("kwiecień"))
                {
                    Miesiac = 4;
                }
                else if (wartosc.Contains("maj"))
                {
                    Miesiac = 5;
                }
                else if (wartosc.Contains("czerwiec"))
                {
                    Miesiac = 6;
                }
                else if (wartosc.Contains("lipiec"))
                {
                    Miesiac = 7;
                }
                else if (wartosc.Contains("sierpień"))
                {
                    Miesiac = 8;
                }
                else if (wartosc.Contains("wrzesień"))
                {
                    Miesiac = 9;
                }
                else if (wartosc.Contains("październik"))
                {
                    Miesiac = 10;
                }
                else if (wartosc.Contains("listopad"))
                {
                    Miesiac = 11;
                }
                else if (wartosc.Contains("grudzień"))
                {
                    Miesiac = 12;
                }
                else
                {
                    Miesiac = 0;
                }
            }
        }
        private class Dane_Dnia
        {
            public int Nr_Dnia  = 0;
            public TimeSpan Godzina_Pracy_Od  = TimeSpan.Zero;
            public TimeSpan Godzina_Pracy_Do  = TimeSpan.Zero;
        }
        public static Error_Logger Internal_Error_Logger = new(true);
        public static void Process_Zakladka(IXLWorksheet Zakladka, Error_Logger error_Logger)
        {
            Internal_Error_Logger = error_Logger;
            List<Current_Position> Lista_Pozycji_Grafików_Z_Zakladki = Helper.Find_Starting_Points(Zakladka, "Data");
            List<Grafik> grafiki = new();
            foreach (Current_Position Startpozycja in Lista_Pozycji_Grafików_Z_Zakladki)
            {
                Current_Position pozycja = Startpozycja;
                int counter = 0;
                while (true)
                {
                    int rowOffset = -5;
                    string dane = "";
                    Grafik grafik = new();
                    try
                    {
                        dane = Zakladka.Cell(pozycja.Row + rowOffset, pozycja.Col + 3).GetFormattedString().Trim();
                    }
                    catch
                    {
                        rowOffset = -4;
                        try
                        {
                            dane = Zakladka.Cell(pozycja.Row + rowOffset, pozycja.Col + 3).GetFormattedString().Trim();
                        }
                        catch
                        {
                            rowOffset = -3;
                            try
                            {
                                dane = Zakladka.Cell(pozycja.Row + rowOffset, pozycja.Col + 3).GetFormattedString().Trim();
                            }
                            catch
                            {
                                Internal_Error_Logger.New_Error(dane, "Naglowek", pozycja.Col + 3, pozycja.Row - 5, "Zły format pliku");
                                throw new Exception(Internal_Error_Logger.Get_Error_String());
                            }
                        }
                    }// xddddddddddd
                    grafik.Set_Miesiac(dane);
                    if (grafik.Miesiac < 1)
                    {
                        Internal_Error_Logger.New_Error(dane, "Miesiac", pozycja.Col + 3, pozycja.Row + rowOffset, "Błędna wartość w mieisac");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }


                    dane = Zakladka.Cell(pozycja.Row + rowOffset, pozycja.Col + 6).GetFormattedString().Trim();
                    if (string.IsNullOrEmpty(dane))
                    {
                        Internal_Error_Logger.New_Error(dane, "Rok", pozycja.Col + 5, pozycja.Row + rowOffset, "Błędna wartość w rok");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }

                    if (int.TryParse(dane, out int tmprok))
                    {
                        grafik.Rok = tmprok;
                    }
                    else
                    {
                        Internal_Error_Logger.New_Error(dane, "Rok", pozycja.Col + 5, pozycja.Row + rowOffset, "Błędna wartość w rok");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }

                    grafik.Pracownik = Get_Pracownik(Zakladka, new Current_Position { Row = Startpozycja.Row, Col = Startpozycja.Col + ((counter * 3) + 1) });
                    if (string.IsNullOrEmpty(grafik.Pracownik.Imie) && string.IsNullOrEmpty(grafik.Pracownik.Nazwisko) && string.IsNullOrEmpty(grafik.Pracownik.Akronim))
                    {
                        break;
                    }

                    List<Dane_Dnia> dane2 = Get_Dane_Dni(Zakladka, new Current_Position { Row = Startpozycja.Row + 4, Col = Startpozycja.Col + ((counter * 3) + 1) });
                    foreach (Dane_Dnia d in dane2)
                    {
                        grafik.Dane_Dni.Add(d);
                    }
                    grafiki.Add(grafik);
                    counter++;
                }
            }
            if (grafiki.Count > 0)
            {
                Dodaj_Plany_do_Optimy(grafiki);
            }
            else
            {
                Internal_Error_Logger.New_Custom_Error("Zły format pliku, nie znaleniono żadnych grafików z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                throw new Exception("Zły format pliku, nie znaleniono żadnych grafików z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
            }
        }
        private static Pracownik Get_Pracownik(IXLWorksheet worksheet, Current_Position pozycja)
        {
            Pracownik pracownik = new Pracownik();
            string pole1 = "";
            string pole2 = "";
            int offset = 0;
            while (true)
            {
                try
                {
                    pozycja.Row--;
                    if (pozycja.Row < 1)
                    {
                        return pracownik;
                    }
                    try
                    {
                        pole1 = worksheet.Cell(pozycja.Row, pozycja.Col).GetFormattedString().Trim();
                    }
                    catch
                    {
                        return new Pracownik();
                    }
                    if (pole1 != "Godziny pracy od")
                    {
                        offset++;
                        for (int i = 0; i < 3; i++)
                        {
                            pole1 = worksheet.Cell(pozycja.Row, pozycja.Col + i).GetFormattedString().Trim();
                            if (!string.IsNullOrEmpty(pole1))
                            {
                                if (offset == 1)
                                {
                                    pole2 = worksheet.Cell(pozycja.Row - 1, pozycja.Col).GetFormattedString().Trim();
                                }
                                if (offset == 2)
                                {
                                    pole2 = pole1;
                                }
                                break;
                            }
                        }
                        if (!string.IsNullOrEmpty(pole2))
                        {
                            break;
                        }
                    }
                }
                catch
                {
                    return new Pracownik();
                }
                
            }
            if (!string.IsNullOrEmpty(pole1) && int.TryParse(pole1, out int impa))
            {
                pracownik.Akronim = pole1;
                if (!string.IsNullOrEmpty(pole2))
                {
                    string[] parts = pole2.Split(" ");
                    if (parts.Length == 2)
                    {
                        pracownik.Nazwisko = parts[0];
                        pracownik.Imie = parts[1];
                    }
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(pole2))
                {
                    string[] parts = pole2.Split(" ");
                    if (parts.Length == 2)
                    {
                        pracownik.Nazwisko = parts[0];
                        pracownik.Imie = parts[1];
                    }
                    else if (parts.Length == 3)
                    {
                        pracownik.Nazwisko = parts[0];
                        pracownik.Imie = parts[1];
                        if (int.TryParse(parts[2], out int tmpint))
                        {
                            pracownik.Akronim = parts[2];
                        }
                    }
                    else
                    {
                        Internal_Error_Logger.New_Error(pole1, "Imie nazwisko akronim", pozycja.Col, pozycja.Row, "Błędny format danych w komórkach imie nazwisko akronim");
                        throw new Exception(Internal_Error_Logger.Get_Error_String());
                    }
                }
            }
            return pracownik;
        }
        private static List<Dane_Dnia> Get_Dane_Dni(IXLWorksheet worksheet, Current_Position pozycja)
        {
            List<Dane_Dnia> Dane_Dni = new();
            for (int i = 0; i < 31; i++)
            {
                string dane = "";
                string danedzien = worksheet.Cell(pozycja.Row, 1).GetFormattedString().Trim(); ;
                if (string.IsNullOrEmpty(danedzien))
                {
                    break;
                }
                Dane_Dnia dane_Dnia = new Dane_Dnia();
                dane_Dnia.Nr_Dnia = i + 1;
                dane = worksheet.Cell(pozycja.Row, pozycja.Col).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    pozycja.Row += 1;
                    continue;
                }

                if(!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref dane_Dnia.Godzina_Pracy_Od))
                {
                    Internal_Error_Logger.New_Error(dane, "Godzina pracy od", pozycja.Col, pozycja.Row, "Błędna wartość w godzinie pracy od");
                    throw new Exception(Internal_Error_Logger.Get_Error_String());
                }

                dane = worksheet.Cell(pozycja.Row, pozycja.Col + 1).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    pozycja.Row += 1;
                    continue;
                }
                if (!Helper.Try_Get_Type_From_String<TimeSpan>(dane, ref dane_Dnia.Godzina_Pracy_Do))
                {
                    Internal_Error_Logger.New_Error(dane, "Godzina pracy do", pozycja.Col + 1, pozycja.Row, "Błędna wartość w godzinie pracy do");
                    throw new Exception(Internal_Error_Logger.Get_Error_String());
                }
                Dane_Dni.Add(dane_Dnia);
                pozycja.Row += 1;
                
            }
            return Dane_Dni;
        }
        private static void Dodaj_Plany_do_Optimy(List<Grafik> grafiki)
        {
            int dodano = 0;
            using (SqlConnection connection = new SqlConnection(Program.config.Optima_Conection_String))
            {
                connection.Open();
                using (SqlTransaction tran = connection.BeginTransaction())
                {
                    foreach (Grafik grafik in grafiki)
                    {
                        foreach (Dane_Dnia dane_DniA in grafik.Dane_Dni)
                        {
                            try
                            {
                                dodano += Zrob_Insert_Plan_command(connection, tran, grafik, grafik.Pracownik, DateTime.ParseExact($"{grafik.Rok}-{grafik.Miesiac:D2}-{dane_DniA.Nr_Dnia:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture), dane_DniA.Godzina_Pracy_Od, dane_DniA.Godzina_Pracy_Do);
                            }
                            catch (SqlException ex)
                            {
                                tran.Rollback();
                                Internal_Error_Logger.New_Custom_Error(ex.Message + " z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                                throw new Exception(ex.Message + $" w pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}" + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                            }
                            catch (FormatException)
                            {
                                continue;
                            }
                            catch (Exception ex)
                            {
                                tran.Rollback();
                                Internal_Error_Logger.New_Custom_Error(ex.Message + " z pliku: " + Internal_Error_Logger.Nazwa_Pliku + " z zakladki: " + Internal_Error_Logger.Nr_Zakladki + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                                throw new Exception(ex.Message + $" w pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}" + " nazwa zakladki: " + Internal_Error_Logger.Nazwa_Zakladki);
                            }
                        }
                    }
                    tran.Commit();
                }
                if (dodano > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Poprawnie dodawno plan z pliku {Internal_Error_Logger.Nazwa_Pliku} z zakladki {Internal_Error_Logger.Nr_Zakladki}");
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }
        }
        private static int Zrob_Insert_Plan_command(SqlConnection connection, SqlTransaction transaction, Grafik grafik, Pracownik pracownik, DateTime data, TimeSpan startGodz, TimeSpan endGodz)
        {
            int IdPracownika = -1;
            IdPracownika = pracownik.Get_PraId(connection, transaction);

            using (SqlCommand cmd = new(@"
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
END", connection, transaction))
            {
                cmd.Parameters.AddWithValue("@DataInsert", data);
                cmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = (DateTime)(Helper.baseDate + startGodz);
                cmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = (DateTime)(Helper.baseDate + endGodz);
                cmd.Parameters.AddWithValue("@PRI_PraId", IdPracownika);
                if ((int)cmd.ExecuteScalar() == 1)
                {
                    return 0;
                }
            }
            using (SqlCommand insertCmd = new($@"
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
	        2,
	        1,
	        1,
	        '');", connection, transaction))
            {
                insertCmd.Parameters.AddWithValue("@DataInsert", data);
                insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = (DateTime)(Helper.baseDate + startGodz);
                insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = (DateTime)(Helper.baseDate + endGodz);
                insertCmd.Parameters.AddWithValue("@PRI_PraId", IdPracownika);
                insertCmd.Parameters.AddWithValue("@ImieMod", Truncate(Internal_Error_Logger.Last_Mod_Osoba, 20));
                insertCmd.Parameters.AddWithValue("@NazwiskoMod", Truncate(Internal_Error_Logger.Last_Mod_Osoba, 50));
                insertCmd.Parameters.AddWithValue("@DataMod", Internal_Error_Logger.Last_Mod_Time);

                insertCmd.ExecuteScalar();
            }
            return 1;
        }
    }
}
