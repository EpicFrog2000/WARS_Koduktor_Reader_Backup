using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;
using System.Data;
using System.Diagnostics;

namespace Excel_Data_Importer_WARS
{
    static class Program
    {
        private static Error_Logger error_logger = new(true); // true - Console write message on creating new error
        private static Config config = new();
        private static readonly bool LOG_TO_TERMINAL = true;
        private static readonly bool Do_Stuff_In_loop = false;
        private enum Typ_Zakladki
        {
            Nierozopznana = -1,
            Tabela_Stawek = 0,
            Karta_Ewidencji_Konduktora = 1,
            Karta_Ewidencji_Pracownika = 2,
            Grafik_Pracy_Pracownika = 3
        }

        public static int Main()
        {
            try
            {
                Do_The_Thing();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Pomiar.Display_Times();
                Console.ReadLine();
            }

            return 0;
        }
        private static void Do_The_Thing()
        {
            // Start measuring time
            Stopwatch stopwatch = new();
            stopwatch.Start();

            TextWriter originalOut = Console.Out;
            if (!LOG_TO_TERMINAL)
            {
                Console.SetOut(TextWriter.Null);
            }

            do
            {
                if (!config.GetConfigFromFile()) // zwraca czy plik konfiguracyjny był utowrzony
                {
                    // Zakończ pracy aby użytkownik mógł zmienić konfigurację
                    return;
                }

                if (!DbManager.Valid_SQLConnection_String())
                {
                    Console.WriteLine($"Invalid connection string: {DbManager.Connection_String}");
                    Console.ReadLine();
                }

                if (config.Files_Folders.Count < 1)
                {
                    Console.WriteLine($"Program nie ma ustawionych folderów");
                    Console.ReadLine();
                }

                foreach (string Folder_Path in config.Files_Folders)
                {
                    if (!Directory.Exists(Folder_Path))
                    {
                        Console.WriteLine($"Program nie znalazł folderu: {Folder_Path}");
                        continue;
                    }
                    string[] Files_Paths = Directory.GetFiles(Folder_Path);
                    if (Files_Paths.Length < 1)
                    {
                        Console.WriteLine($"Program nie znalazł żadnych plików w folderze: {Folder_Path}");
                        continue;
                    }

                    Check_Base_Dirs(Folder_Path);

                    foreach (string filepath in Files_Paths)
                    {
                        Process_Files(filepath);
                    }

                    //await Parallel.ForEachAsync(Files_Paths, async (filePath, _) =>
                    //{
                    //    await Task.Run(() => Process_Files(filePath));
                    //});

                    //await Task.WhenAll(Files_Paths.Select(filePath => Process_Files(filePath)).ToList());
                }
            } while (Do_Stuff_In_loop);


            if (!LOG_TO_TERMINAL)
            {
                Console.SetOut(originalOut);
            }
            Console.WriteLine("Czas wykonania programu: " + stopwatch.Elapsed.ToString(@"hh\:mm\:ss\:fff"));
            Console.WriteLine("Kliknij aby zakończyć...");
            Console.ReadLine();
        }
        private static void Process_Files(string File_Path)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"Czytanie: {Path.GetFileNameWithoutExtension(File_Path)} {DateTime.Now}");
            Console.ForegroundColor = ConsoleColor.White;
            (XLWorkbook? Workbook, FileStream? Stream) = Open_Workbook(File_Path);
            if (Workbook == null)
            {
                if (Stream != null)
                {
                    Stream!.Close();
                }
                Move_File(File_Path, 0);
                Pomiar.Avg_Process_Files = PomiaryStopWatch.Elapsed;
                return;
            }

            Usun_Ukryte_Karty(Workbook);

            Error_Logger Internal_Error_Logger = new(true)
            {
                Good_Files_Folder = error_logger.Good_Files_Folder,
                Current_Processed_Files_Folder = error_logger.Current_Processed_Files_Folder,
                Current_Bad_Files_Folder = error_logger.Current_Bad_Files_Folder,
                ErrorFilePath = error_logger.ErrorFilePath,
                Nazwa_Pliku = File_Path
            };

            (Internal_Error_Logger.Last_Mod_Osoba, Internal_Error_Logger.Last_Mod_Time) = Get_Metadane_Pliku(Workbook, File_Path);
            error_logger.Nazwa_Pliku = Internal_Error_Logger.Nazwa_Pliku;
            error_logger.Last_Mod_Osoba = Internal_Error_Logger.Last_Mod_Osoba;
            error_logger.Last_Mod_Time = Internal_Error_Logger.Last_Mod_Time;

            int Ilosc_Zakladek_W_Workbook = Workbook.Worksheets.Count;
            if (Ilosc_Zakladek_W_Workbook < 1)
            {
                Pomiar.Avg_Process_Files = PomiaryStopWatch.Elapsed;
                Workbook!.Dispose();
                Stream!.Close();
                return;
            }
            bool Contains_Any_Bad_Data = false;
            for (int Obecny_Numer_Zakladki = 1; Obecny_Numer_Zakladki <= Ilosc_Zakladek_W_Workbook; Obecny_Numer_Zakladki++)
            {
                Internal_Error_Logger.Nr_Zakladki = Obecny_Numer_Zakladki;
                IXLWorksheet Zakladka = Workbook.Worksheet(Obecny_Numer_Zakladki);
                Internal_Error_Logger.Nazwa_Zakladki = Zakladka.Name;
                Typ_Zakladki Typ_Zakladki = Get_Typ_Zakladki(Zakladka);
                try
                {
                    switch (Typ_Zakladki)
                    {
                        case Typ_Zakladki.Tabela_Stawek:
                            Reader_Tabela_Stawek_v1.Process_Zakladka(Zakladka, Internal_Error_Logger);
                            break;
                        case Typ_Zakladki.Karta_Ewidencji_Konduktora:
                            Obecny_Numer_Zakladki = Ilosc_Zakladek_W_Workbook + 1;
                            Reader_Karta_Ewidencji_Konduktora_v1.Process_Zakladka(Zakladka, Internal_Error_Logger);
                            break;
                        case Typ_Zakladki.Karta_Ewidencji_Pracownika:
                            Reader_Karta_Ewidencji_Pracownika.Process_Zakladka(Zakladka, Internal_Error_Logger);
                            break;
                        case Typ_Zakladki.Grafik_Pracy_Pracownika:
                            Reader_Grafik_Pracy_Pracownika_2025_v3.Process_Zakladka(Zakladka, Internal_Error_Logger);
                            break;
                        case Typ_Zakladki.Nierozopznana:
                            _ = Copy_Bad_Sheet_To_Files_Folder(Workbook.Properties, Zakladka, File_Path);
                            Workbook!.Dispose();
                            Stream!.Close();
                            Move_File(File_Path, 2);
                            error_logger.New_Custom_Error($"Nie rozpoznano tego typu zakładki w pliku: \"{error_logger.Nazwa_Pliku}\" zakladka: \"{error_logger.Nazwa_Zakladki}\" numer zakładki: \"{error_logger.Nr_Zakladki}\"");
                            Pomiar.Avg_Process_Files = PomiaryStopWatch.Elapsed;
                            return;
                    }
                }
                catch
                {
                    _ = Copy_Bad_Sheet_To_Files_Folder(Workbook.Properties, Zakladka, File_Path);
                    Contains_Any_Bad_Data = true;
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Pomiar.Avg_Process_Files = PomiaryStopWatch.Elapsed;
                }
            }
            Workbook!.Dispose();
            Stream!.Close();
            if (!Contains_Any_Bad_Data)
            {
                Move_File(File_Path, 1);
            }
            else
            {
                Move_File(File_Path, 2);
            }
        }
        private static Typ_Zakladki Get_Typ_Zakladki(IXLWorksheet Worksheet)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            string Cell_Value = Worksheet.Cell(3, 5).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("Harmonogram pracy"))
            {
                Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Typ_Zakladki.Grafik_Pracy_Pracownika;
            }

            Cell_Value = Worksheet.Cell(1, 3).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("Tabela Stawek"))
            {
                Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Typ_Zakladki.Tabela_Stawek;
            }

            Cell_Value = Worksheet.Cell(1, 1).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("KARTA EWIDENCJI CZASU PRACY"))
            {
                Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Typ_Zakladki.Karta_Ewidencji_Konduktora;
            }

            Cell_Value = Worksheet.Cell(1, 1).Value.ToString();
            if (Cell_Value.Trim().StartsWith("GRAFIK PRACY MIESIĄC")) // grafik v2024 v2
            {
                Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Typ_Zakladki.Grafik_Pracy_Pracownika;
            }

            Cell_Value = Worksheet.Cell(1, 2).Value.ToString();
            if (Cell_Value.Trim().StartsWith("GRAFIK PRACY MIESIĄC")) // grafik v2024 v2
            {
                Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Typ_Zakladki.Grafik_Pracy_Pracownika;
            }

            foreach (IXLCell cell in Worksheet.CellsUsed()) // karta pracy NIE konduktora
            {
                try
                {
                    if (cell.GetString().Trim() == "Dzień")
                    {
                        Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                        return Typ_Zakladki.Karta_Ewidencji_Pracownika;
                    }
                }
                catch { }
            }
            Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
            return Typ_Zakladki.Nierozopznana;
        }
        private static void Usun_Ukryte_Karty(XLWorkbook workbook)
        {
            Stopwatch PomiaryStopWatch = Stopwatch.StartNew();
            IXLWorksheet[] sheetsToRemove = workbook.Worksheets.Where(s => s.Visibility == XLWorksheetVisibility.Hidden).ToArray();
            foreach (IXLWorksheet sheet in sheetsToRemove)
            {
                workbook.Worksheets.Delete(sheet.Name);
            }
            workbook.Save();
            Pomiar.Avg_Usun_Ukryte_Karty = PomiaryStopWatch.Elapsed;
        }
        private static (string, DateTime) Get_Metadane_Pliku(XLWorkbook Workbook, string File_Path)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            DateTime lastWriteTime = File.GetLastWriteTime(File_Path);
            string lastModifiedBy = Workbook.Properties.LastModifiedBy ?? "";
            Pomiar.Avg_Get_Metadane_Pliku = PomiaryStopWatch.Elapsed;
            return (lastModifiedBy, lastWriteTime);
        }
        private static (XLWorkbook?, FileStream?) Open_Workbook(string File_Path)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            XLWorkbook Workbook;
            try
            {
                FileStream stream = new(File_Path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite, bufferSize: 8192);
                Workbook = new(stream);
                Pomiar.Avg_Open_Workbook = PomiaryStopWatch.Elapsed;
                return (Workbook, stream);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"Program nie może odczytać pliku {File_Path}");
                Console.ForegroundColor = ConsoleColor.White;
                Pomiar.Avg_Open_Workbook = PomiaryStopWatch.Elapsed;
                return (null, null);
            }
        }
        private static void Move_File(string filePath, int opcja)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            if (!config.Move_Files_To_Processed_Folder)
            {
                Pomiar.Avg_MoveFile = PomiaryStopWatch.Elapsed;
                return;
            }
            try
            {

                string destinationPath = string.Empty;
                if (opcja == 0)
                {
                    destinationPath = Path.Combine(error_logger.Current_Bad_Files_Folder, Path.GetFileName(filePath));
                }
                else if (opcja == 1)
                {
                    destinationPath = Path.Combine(error_logger.Good_Files_Folder, Path.GetFileName(filePath));
                }
                else if (opcja == 2)
                {
                    destinationPath = Path.Combine(error_logger.Current_Processed_Files_Folder, Path.GetFileName(filePath));
                    if (File.Exists(destinationPath))
                    {
                        File.Delete(destinationPath);
                    }
                    File.Move(filePath, destinationPath);
                    Pomiar.Avg_MoveFile = PomiaryStopWatch.Elapsed;
                    return;
                }
                else
                {
                    return;
                }

                if (File.Exists(destinationPath))
                {
                    File.Delete(destinationPath);
                }
                File.Copy(filePath, destinationPath);
                destinationPath = Path.Combine(error_logger.Current_Processed_Files_Folder, Path.GetFileName(filePath));
                if (File.Exists(destinationPath))
                {
                    File.Delete(destinationPath);
                }
                File.Move(filePath, destinationPath);
            }
            catch
            {
                error_logger.New_Custom_Error($"Nie udało się przenieść pliku: {filePath}");
            }
            finally
            {
                Pomiar.Avg_MoveFile = PomiaryStopWatch.Elapsed;
            }
        }
        private static async Task Copy_Bad_Sheet_To_Files_Folder(XLWorkbookProperties op, IXLWorksheet sheetToCopy, string filePath)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            string newFilePath = Path.Combine(error_logger.Current_Bad_Files_Folder, $"DO_POPRAWY_{Path.GetFileName(filePath)}");
            string newSheetName = sheetToCopy.Name.Length > 31 ? sheetToCopy.Name[..31] : sheetToCopy.Name;
            using (XLWorkbook workbook = File.Exists(newFilePath) ? new(newFilePath) : new())
            {
                if (!workbook.Worksheets.Contains(newSheetName))
                {
                    sheetToCopy.CopyTo(workbook, newSheetName);
                    op.Author = "Kopia wykonana przez importer";
                    op.Modified = DateTime.Now;
                    await Task.Run(() => workbook.SaveAs(newFilePath));
                }
            }
            Pomiar.Avg_Copy_Bad_Sheet_To_Files_Folder = PomiaryStopWatch.Elapsed;
        }
        private static void Check_Base_Dirs(string path)
        {
            string[] directories =
            [
                Path.Combine(path, "Errors"),
                Path.Combine(path, "Bad_Files"),
                Path.Combine(path, "Processed_Files"),
                Path.Combine(path, "Good_Files")
            ];
            string errorFilePath = Path.Combine(directories[0], "Errors.txt");
            error_logger.Current_Processed_Files_Folder = directories[2];
            error_logger.Current_Bad_Files_Folder = directories[1];
            error_logger.Set_Error_File_Path(directories[0]);
            error_logger.Good_Files_Folder = directories[3];
            foreach (string dir in directories)
            {
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
            }

            if (config.Clear_Logs_On_Program_Restart)
            {
                foreach (string file in Directory.GetFiles(directories[0]))
                {
                    File.Delete(file);
                }
            }
            if (config.Clear_Bad_Files_On_Restart)
            {
                foreach (string file in Directory.GetFiles(directories[1]))
                {
                    File.Delete(file);
                }
            }
            if (config.Clear_Processed_Files_On_Restart)
            {
                foreach (string file in Directory.GetFiles(directories[2]))
                {
                    File.Delete(file);
                }
            }
            if (config.Clear_Good_Files_On_Restart)
            {
                foreach (string file in Directory.GetFiles(directories[3]))
                {
                    File.Delete(file);
                }
            }

            if (!File.Exists(errorFilePath))
            {
                File.Create(errorFilePath);
            }
        }
        public static void Convert_To_Xlsx(string inputFilePath, string outputFilePath)
        {
            // nwm dlaczego textwrap jest zawsze true. Jebać to jest wystarczająco dobre.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            DataSet dataSet;
            using (FileStream stream = File.Open(inputFilePath, FileMode.Open, FileAccess.Read))
            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
            {
                ExcelDataSetConfiguration config = new()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                dataSet = reader.AsDataSet(config);
            }

            using XLWorkbook workbook = new();
            foreach (DataTable table in dataSet.Tables)
            {
                IXLWorksheet worksheet = workbook.Worksheets.Add(table.TableName);
                for (int i = 0; i < table.Columns.Count; i++)
                    worksheet.Cell(1, i + 1).Value = table.Columns[i].ColumnName;

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    for (int j = 0; j < table.Columns.Count; j++)
                    {
                        object value = table.Rows[i][j];

                        if (value == DBNull.Value)
                        {
                            worksheet.Cell(i + 2, j + 1).Value = string.Empty;
                        }
                        else
                        {
                            worksheet.Cell(i + 2, j + 1).Value = value.ToString();
                        }
                    }
                }
            }
            workbook.SaveAs(outputFilePath);
            //(string o, DateTime d) = Get_Metadane_Pliku(inputFilePath);
            //workbook.Properties.LastModifiedBy = o;
            //workbook.Properties.Modified = d;
            workbook.SaveAs(outputFilePath);
        } //Kiedyś używane do konwertowania plików xls na xlsx ale w sumie to wyjebane (pora umierać)
    }
    static class Pomiar
    {
        private static TimeSpan avg_Get_Metadane_Pliku = TimeSpan.Zero;
        private static TimeSpan avg_Process_Files = TimeSpan.Zero;
        private static TimeSpan avg_MoveFile = TimeSpan.Zero;
        private static TimeSpan avg_Copy_Bad_Sheet_To_Files_Folder = TimeSpan.Zero;
        private static TimeSpan avg_Open_Workbook = TimeSpan.Zero;
        private static TimeSpan avg_Get_Typ_Zakladki = TimeSpan.Zero;
        private static TimeSpan avg_Usun_Ukryte_Karty = TimeSpan.Zero;
        public static TimeSpan Avg_Get_Metadane_Pliku
        {
            get => avg_Get_Metadane_Pliku;
            set
            {
                if (avg_Get_Metadane_Pliku == TimeSpan.Zero)
                {
                    avg_Get_Metadane_Pliku = value;
                    return;
                }
                avg_Get_Metadane_Pliku = (value + avg_Get_Metadane_Pliku) / 2;
            }
        }
        public static TimeSpan Avg_Process_Files
        {
            get => avg_Process_Files;
            set
            {
                if (avg_Process_Files == TimeSpan.Zero)
                {
                    avg_Process_Files = value;
                    return;
                }
                avg_Process_Files = (value + avg_Process_Files) / 2;
            }
        }
        public static TimeSpan Avg_MoveFile
        {
            get => avg_MoveFile;
            set
            {
                if (avg_MoveFile == TimeSpan.Zero)
                {
                    avg_MoveFile = value;
                    return;
                }
                avg_MoveFile = (value + avg_MoveFile) / 2;
            }
        }
        public static TimeSpan Avg_Copy_Bad_Sheet_To_Files_Folder
        {
            get => avg_Copy_Bad_Sheet_To_Files_Folder;
            set
            {
                if (avg_Copy_Bad_Sheet_To_Files_Folder == TimeSpan.Zero)
                {
                    avg_Copy_Bad_Sheet_To_Files_Folder = value;
                    return;
                }
                avg_Copy_Bad_Sheet_To_Files_Folder = (value + avg_Copy_Bad_Sheet_To_Files_Folder) / 2;
            }
        }
        public static TimeSpan Avg_Open_Workbook
        {
            get => avg_Open_Workbook;
            set
            {
                if (avg_Open_Workbook == TimeSpan.Zero)
                {
                    avg_Open_Workbook = value;
                    return;
                }
                avg_Open_Workbook = (value + avg_Open_Workbook) / 2;
            }
        }
        public static TimeSpan Avg_Get_Typ_Zakladki
        {
            get => avg_Get_Typ_Zakladki;
            set
            {
                if (avg_Get_Typ_Zakladki == TimeSpan.Zero)
                {
                    avg_Get_Typ_Zakladki = value;
                    return;
                }
                avg_Get_Typ_Zakladki = (value + avg_Get_Typ_Zakladki) / 2;
            }
        }
        public static TimeSpan Avg_Usun_Ukryte_Karty
        {
            get => avg_Usun_Ukryte_Karty;
            set
            {
                if (avg_Usun_Ukryte_Karty == TimeSpan.Zero)
                {
                    avg_Usun_Ukryte_Karty = value;
                    return;
                }
                avg_Usun_Ukryte_Karty = (value + avg_Usun_Ukryte_Karty) / 2;
            }
        }
        public static void Display_Times()
        {
            Console.WriteLine($"Pomiar.Avg_Get_Typ_Zakladki: {Avg_Get_Typ_Zakladki}");
            Console.WriteLine($"Pomiar.Avg_Get_Metadane_Pliku: {Avg_Get_Metadane_Pliku}");
            Console.WriteLine($"Pomiar.Avg_Open_Workbook: {Avg_Open_Workbook}");
            Console.WriteLine($"Pomiar.Avg_Process_Files: {Avg_Process_Files}");
            Console.WriteLine($"Pomiar.Avg_MoveFile: {Avg_MoveFile}");
            Console.WriteLine($"Pomiar.Avg_Copy_Bad_Sheet_To_Files_Folder: {Avg_Copy_Bad_Sheet_To_Files_Folder}");
            Console.WriteLine($"Pomiar.Avg_Usun_Ukryte_Karty: {Avg_Usun_Ukryte_Karty}");
        }
    }

}

