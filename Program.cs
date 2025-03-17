using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using static Excel_Data_Importer_WARS.DbManager;

namespace Excel_Data_Importer_WARS
{
    static class Program
    {
        private static Error_Logger error_logger = new(true); // true - Console write message on creating new error
        private static Config config = new();
        private static readonly bool LOG_TO_TERMINAL = true;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
        [DllImportAttribute("user32.dll", CharSet = CharSet.Unicode)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        public static async Task Main()
        {
            Tylko_Jedna_Instancja();

            try
            {
                await Do_The_Thing();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Helper.Pomiar.Display_Times();
                Console.WriteLine("Kliknij aby zakończyć...");
                Console.ReadLine();
            }
        }
        private static async Task Do_The_Thing()
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
                    Console.WriteLine($"Invalid connection string");
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

                    Check_Base_Dirs(Folder_Path);


                    DbManager.OpenConnection();

                    string[] files = Directory.GetFiles(Folder_Path, "*.xlsx");

                    const int batchSize = 8;
                    for (int i = 0; i < files.Length; i += batchSize)
                    {
                        var batch = files.Skip(i).Take(batchSize).Select(file => Process_Files(file)); // Coś długo się ta funkcja wykonuje, TODO fix
                        await Task.WhenAll(batch);                                                      // Nvm, może jednak ok ale i tak coś mi tu śmierdzi z czasem wykonywania.  
                    }

                    DbManager.CloseConnection();
                }
            } while (config.Tryb_Zapetlony);


            if (!LOG_TO_TERMINAL)
            {
                Console.SetOut(originalOut);
            }
            Console.WriteLine("Czas wykonania programu: " + stopwatch.Elapsed.ToString(@"hh\:mm\:ss\:fff"));
        }
        private static async Task Process_Files(string File_Path)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"[{DateTime.Now}] Czytanie: {Path.GetFileNameWithoutExtension(File_Path)}");
            Console.ForegroundColor = ConsoleColor.White;
            
            XLWorkbook? Workbook = await Open_Workbook(File_Path);
            if (Workbook == null)
            {
                Move_File(File_Path, Move_File_Opcje.Bad_Files_Folder);
                return;
            }

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
                Workbook!.Dispose();
                return;
            }
            
            bool Contains_Any_Bad_Data = false;

            for (int Obecny_Numer_Zakladki = 1; Obecny_Numer_Zakladki <= Ilosc_Zakladek_W_Workbook; Obecny_Numer_Zakladki++)
            {
                Stopwatch PomiaryStopWatch_zakladka = new();
                PomiaryStopWatch_zakladka.Restart();

                Internal_Error_Logger.Nr_Zakladki = Obecny_Numer_Zakladki;
                IXLWorksheet Zakladka = Workbook.Worksheet(Obecny_Numer_Zakladki);
                if (Zakladka.Visibility == XLWorksheetVisibility.Hidden || Zakladka.Visibility == XLWorksheetVisibility.VeryHidden)
                {
                    continue;
                }
                Internal_Error_Logger.Nazwa_Zakladki = Zakladka.Name;
                Helper.Typ_Zakladki Typ_Zakladki = Get_Typ_Zakladki(Zakladka);

                try
                {
                    switch (Typ_Zakladki)
                    {
                        case Helper.Typ_Zakladki.Tabela_Stawek:
                            await Reader_Tabela_Stawek_v1.Process_Zakladka(Zakladka, Internal_Error_Logger);
                            break;
                        case Helper.Typ_Zakladki.Karta_Ewidencji_Konduktora:
                            Obecny_Numer_Zakladki = Ilosc_Zakladek_W_Workbook + 1;
                            await Reader_Karta_Ewidencji_Konduktora_v1.Process_Zakladka(Zakladka, Internal_Error_Logger);
                            break;
                        case Helper.Typ_Zakladki.Karta_Ewidencji_Pracownika:
                            await Reader_Karta_Ewidencji_Pracownika.Process_Zakladka(Zakladka, Internal_Error_Logger);
                            break;
                        case Helper.Typ_Zakladki.Grafik_Pracy_Pracownika:
                            await Reader_Grafik_Pracy_Pracownika_2025_v3.Process_Zakladka(Zakladka, Internal_Error_Logger);
                            break;
                        case Helper.Typ_Zakladki.Harmonogram_Pracy_Konduktora:
                            await Reader_Harmonogram_Pracy_Konduktora.Process_Zakladka(Zakladka, Internal_Error_Logger);
                            break;
                        case Helper.Typ_Zakladki.Nierozopznana:
                            error_logger.New_Custom_Error($"Nie rozpoznano tego typu zakładki w pliku: \"{error_logger.Nazwa_Pliku}\" zakladka: \"{error_logger.Nazwa_Zakladki}\" numer zakładki: \"{error_logger.Nr_Zakladki}\"", false);
                            Contains_Any_Bad_Data = true;
                            _ = Copy_Bad_Sheet_To_Files_Folder(Workbook.Properties, Zakladka, File_Path);
                            break;
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
                    Helper.Pomiar.Avg_Process_1_Zakladka = PomiaryStopWatch_zakladka.Elapsed;
                }
            }
            Workbook!.Dispose();
            if (!Contains_Any_Bad_Data)
            {
                Move_File(File_Path, Move_File_Opcje.Good_Files_Folder);
            }
            else
            {
                Move_File(File_Path, Move_File_Opcje.Processed_Files_Folder);
            }
            Helper.Pomiar.Avg_Process_Files = PomiaryStopWatch.Elapsed;

        }
        private static Helper.Typ_Zakladki Get_Typ_Zakladki(IXLWorksheet Worksheet)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            string Cell_Value = Worksheet.Cell(3, 5).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("Harmonogram pracy"))
            {
                foreach (IXLCell cell in Worksheet.CellsUsed())
                {
                    try
                    {
                        if (cell.GetString().Trim() == "Czas odpoczynku (wliczany do CP)")
                        {
                            Helper.Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                            return Helper.Typ_Zakladki.Harmonogram_Pracy_Konduktora;
                        }
                    }
                    catch { }
                }
                Helper.Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Helper.Typ_Zakladki.Grafik_Pracy_Pracownika;
            }

            Cell_Value = Worksheet.Cell(1, 3).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("Tabela Stawek"))
            {
                Helper.Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Helper.Typ_Zakladki.Tabela_Stawek;
            }

            Cell_Value = Worksheet.Cell(1, 1).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("KARTA EWIDENCJI CZASU PRACY"))
            {
                Helper.Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Helper.Typ_Zakladki.Karta_Ewidencji_Konduktora;
            }

            Cell_Value = Worksheet.Cell(1, 1).Value.ToString();
            if (Cell_Value.Trim().StartsWith("GRAFIK PRACY MIESIĄC")) // grafik v2024 v2
            {
                Helper.Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Helper.Typ_Zakladki.Grafik_Pracy_Pracownika;
            }

            Cell_Value = Worksheet.Cell(1, 2).Value.ToString();
            if (Cell_Value.Trim().StartsWith("GRAFIK PRACY MIESIĄC")) // grafik v2024 v2
            {
                Helper.Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                return Helper.Typ_Zakladki.Grafik_Pracy_Pracownika;
            }

            foreach (IXLCell cell in Worksheet.CellsUsed()) 
            {
                try
                {
                     if (cell.GetString().Trim() == "Dzień") // karta pracy NIE konduktora
                    {
                        Helper.Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
                        return Helper.Typ_Zakladki.Karta_Ewidencji_Pracownika;
                    }
                }
                catch { }
            }
            Helper.Pomiar.Avg_Get_Typ_Zakladki = PomiaryStopWatch.Elapsed;
            return Helper.Typ_Zakladki.Nierozopznana;
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
            Helper.Pomiar.Avg_Usun_Ukryte_Karty = PomiaryStopWatch.Elapsed;
        }
        private static (string, DateTime) Get_Metadane_Pliku(XLWorkbook Workbook, string File_Path)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            DateTime lastWriteTime = File.GetLastWriteTime(File_Path);
            string lastModifiedBy = Workbook.Properties.LastModifiedBy ?? "";
            Helper.Pomiar.Avg_Get_Metadane_Pliku = PomiaryStopWatch.Elapsed;
            return (lastModifiedBy, lastWriteTime);
        }
        private static async Task<XLWorkbook?> Open_Workbook(string File_Path)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            XLWorkbook Workbook;
            try
            {
                byte[] fileBytes = await File.ReadAllBytesAsync(File_Path);
                using var memoryStream = new MemoryStream(fileBytes);
                Workbook = new(memoryStream);
                Helper.Pomiar.Avg_Open_Workbook = PomiaryStopWatch.Elapsed;
                return Workbook;
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                //error_logger.New_Custom_Error($"Program nie może odczytać pliku {File_Path}", false);
                Console.WriteLine($"Program nie może odczytać pliku {File_Path}");
                Console.ForegroundColor = ConsoleColor.White;
                Helper.Pomiar.Avg_Open_Workbook = PomiaryStopWatch.Elapsed;
                return null;
            }
        }
        private static void Move_File(string filePath, Move_File_Opcje opcja)
        {
            Stopwatch PomiaryStopWatch = new();
            PomiaryStopWatch.Restart();
            if (!config.Move_Files_To_Processed_Folder)
            {
                Helper.Pomiar.Avg_MoveFile = PomiaryStopWatch.Elapsed;
                return;
            }
            try
            {
                string destinationPath = string.Empty;
                switch (opcja)
                {
                    case Move_File_Opcje.Bad_Files_Folder:
                        destinationPath = Path.Combine(error_logger.Current_Bad_Files_Folder, Path.GetFileName(filePath));
                        break;
                    case Move_File_Opcje.Good_Files_Folder:
                        destinationPath = Path.Combine(error_logger.Good_Files_Folder, Path.GetFileName(filePath));
                        break;
                    case Move_File_Opcje.Processed_Files_Folder:
                        destinationPath = Path.Combine(error_logger.Current_Processed_Files_Folder, Path.GetFileName(filePath));
                        if (File.Exists(destinationPath))
                        {
                            File.Delete(destinationPath);
                        }
                        File.Move(filePath, destinationPath);
                        Helper.Pomiar.Avg_MoveFile = PomiaryStopWatch.Elapsed;
                        return;
                    default:
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
                error_logger.New_Custom_Error($"Nie udało się przenieść pliku: {filePath}", false);
            }
            finally
            {
                Helper.Pomiar.Avg_MoveFile = PomiaryStopWatch.Elapsed;
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
            Helper.Pomiar.Avg_Copy_Bad_Sheet_To_Files_Folder = PomiaryStopWatch.Elapsed;
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
        //private static void Convert_To_Xlsx(string inputFilePath, string outputFilePath)
        //{
        //    // nwm dlaczego textwrap jest zawsze true. Jebać to jest wystarczająco dobre.
        //    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        //    DataSet dataSet;
        //    using (FileStream stream = File.Open(inputFilePath, FileMode.Open, FileAccess.Read))
        //    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
        //    {
        //        ExcelDataSetConfiguration config = new()
        //        {
        //            ConfigureDataTable = _ => new ExcelDataTableConfiguration
        //            {
        //                UseHeaderRow = true
        //            }
        //        };
        //        dataSet = reader.AsDataSet(config);
        //    }

        //    using XLWorkbook workbook = new();
        //    foreach (DataTable table in dataSet.Tables)
        //    {
        //        IXLWorksheet worksheet = workbook.Worksheets.Add(table.TableName);
        //        for (int i = 0; i < table.Columns.Count; i++)
        //            worksheet.Cell(1, i + 1).Value = table.Columns[i].ColumnName;

        //        for (int i = 0; i < table.Rows.Count; i++)
        //        {
        //            for (int j = 0; j < table.Columns.Count; j++)
        //            {
        //                object value = table.Rows[i][j];

        //                if (value == DBNull.Value)
        //                {
        //                    worksheet.Cell(i + 2, j + 1).Value = string.Empty;
        //                }
        //                else
        //                {
        //                    worksheet.Cell(i + 2, j + 1).Value = value.ToString();
        //                }
        //            }
        //        }
        //    }
        //    workbook.SaveAs(outputFilePath);
        //    //(string o, DateTime d) = Get_Metadane_Pliku(inputFilePath);
        //    //workbook.Properties.LastModifiedBy = o;
        //    //workbook.Properties.Modified = d;
        //    workbook.SaveAs(outputFilePath);
        //} //Kiedyś używane do konwertowania plików xls na xlsx ale w sumie to wyjebane (pora umierać)
        private static void Tylko_Jedna_Instancja()
        {
            // Jeśli jest uruchomiona już jedna instancja to zakończ program
            // Jeśli jest w trakie procesowania pliku możę być on uszkodzony jeśli program zostanie zakończony ale musiał bym się bawić w mutexy i to jest za dużo roboty -_-
            Process[] processes = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName);
            if (processes.Length > 1)
            {
                IntPtr hWnd = (IntPtr)MessageBox(IntPtr.Zero, "Program jest już uruchomiony. Następi jego zamknięcie", "Informacja", 0);
                SetWindowPos(hWnd, new IntPtr(-1), 0, 0, 0, 0, 0x0002 | 0x0001);
                foreach (Process process in processes) // Strzeli też samobója chyba
                {
                    process.Kill();
                }
                Environment.Exit(0);
            }
        }
        private enum Move_File_Opcje
        {
            Bad_Files_Folder = 0,
            Good_Files_Folder = 1,
            Processed_Files_Folder = 2
        }
    }

}

