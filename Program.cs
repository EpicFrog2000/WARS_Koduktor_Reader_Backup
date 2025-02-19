using ClosedXML.Excel;
using ExcelDataReader;
using System.Data;
using System.Diagnostics;

// TODO better error log messages
// TODO dodawanie komentarza z kart pracy
// TODO Lepsze oznaczenia Absencji

namespace Excel_Data_Importer_WARS
{
    static class Program
    {
        public static Error_Logger error_logger = new(true); // true - Console write message on creating new error
        public static Config config = new();
        public static Stopwatch stopwatch = new();
        public static readonly bool LOG_TO_TERMINAL = false;
        public static async Task<int> Main()
        {
            // Start measuring time

            stopwatch.Start();
            TextWriter originalOut = Console.Out;
            if (!LOG_TO_TERMINAL)
            {
                Console.SetOut(TextWriter.Null);
            }

            config.GetConfigFromFile();

            if (!DbManager.Valid_SQLConnection_String())
            {
                Console.WriteLine($"Invalid connection string: {DbManager.Connection_String}");
                Console.ReadLine();
                return -1;
            }

            if (config.Files_Folders.Count < 1)
            {
                Console.WriteLine($"Program nie ma ustawionych folderów");
                Console.ReadLine();
                return 0;
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

                /*foreach (string filepath in Files_Paths)
                {
                    await Process_Files(filepath).ConfigureAwait(false);
                }*/

                await Task.WhenAll(Files_Paths.Select(filePath => Process_Files(filePath)));

            }

            if (!LOG_TO_TERMINAL)
            {
                Console.SetOut(originalOut);
            }
            Console.WriteLine("Czas wykonania programu: " + stopwatch.Elapsed.ToString(@"hh\:mm\:ss\:fff"));
            Console.WriteLine("Kliknij aby zakończyć...");
            Console.ReadLine();
            return 0;
        }
        private static async Task Process_Files(string File_Path)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"Czytanie: {Path.GetFileNameWithoutExtension(File_Path)} {DateTime.Now}");
            Console.ForegroundColor = ConsoleColor.White;
            if (!Is_File_Valid(File_Path))
            {
                return;
            }
            Error_Logger Internal_Error_Logger = new(true)
            {
                Current_Processed_Files_Folder = error_logger.Current_Processed_Files_Folder,
                Current_Bad_Files_Folder = error_logger.Current_Bad_Files_Folder,
                ErrorFilePath = error_logger.ErrorFilePath,
                Nazwa_Pliku = File_Path
            };

            (Internal_Error_Logger.Last_Mod_Osoba, Internal_Error_Logger.Last_Mod_Time) = await Get_Metadane_Pliku(File_Path).ConfigureAwait(false);

            error_logger.Nazwa_Pliku = Internal_Error_Logger.Nazwa_Pliku;
            error_logger.Last_Mod_Osoba = Internal_Error_Logger.Last_Mod_Osoba;
            error_logger.Last_Mod_Time = Internal_Error_Logger.Last_Mod_Time;



            using var stream = File.Open(File_Path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            using (XLWorkbook Workbook = new(stream))
            {
                await Usun_Ukryte_Karty(Workbook);
                int Ilosc_Zakladek_W_Workbook = Workbook.Worksheets.Count;
                if (Ilosc_Zakladek_W_Workbook < 1)
                {
                    return;
                }
                for (int Obecny_Numer_Zakladki = 1; Obecny_Numer_Zakladki <= Ilosc_Zakladek_W_Workbook; Obecny_Numer_Zakladki++)
                {
                    Internal_Error_Logger.Nr_Zakladki = Obecny_Numer_Zakladki;
                    IXLWorksheet Zakladka = Workbook.Worksheet(Obecny_Numer_Zakladki);
                    Internal_Error_Logger.Nazwa_Zakladki = Zakladka.Name;
                    int Typ_Zakladki = Get_Typ_Zakladki(Zakladka);
                    try
                    {
                        switch (Typ_Zakladki)
                        {
                            case 2:
                                Reader_Tabela_Stawek_v1.Process_Zakladka(Zakladka, Internal_Error_Logger);
                                break;
                            case 3:
                                Obecny_Numer_Zakladki = Ilosc_Zakladek_W_Workbook + 1;
                                Reader_Karta_Ewidencji_Konduktora_v1.Process_Zakladka(Zakladka, Internal_Error_Logger);
                                break;
                            case 4:
                                Reader_Karta_Ewidencji_Pracownika.Process_Zakladka(Zakladka, Internal_Error_Logger);
                                break;
                            case 5:
                                Reader_Grafik_Pracy_Pracownika_2025_v3.Process_Zakladka(Zakladka, Internal_Error_Logger);
                                break;
                            default:
                                await Copy_Bad_Sheet_To_Files_Folder(File_Path, Obecny_Numer_Zakladki).ConfigureAwait(false);
                                error_logger.New_Custom_Error($"Nie rozpoznano tego typu zakładki w pliku: \"{error_logger.Nazwa_Pliku}\" zakladka: \"{error_logger.Nazwa_Zakladki}\" numer zakładki: \"{error_logger.Nr_Zakladki}\"");
                                return;
                        }
                    }
                    catch
                    {
                        await Copy_Bad_Sheet_To_Files_Folder(File_Path, Obecny_Numer_Zakladki).ConfigureAwait(false);
                    }
                }
            }
            await MoveFile(File_Path, 0).ConfigureAwait(false);
        }
        private static int Get_Typ_Zakladki(IXLWorksheet Worksheet)
        {

            string Cell_Value = Worksheet.Cell(3, 5).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("Harmonogram pracy"))
            {
                return 1;
            }

            Cell_Value = Worksheet.Cell(1, 3).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("Tabela Stawek"))
            {
                return 2;
            }

            Cell_Value = Worksheet.Cell(1, 1).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("KARTA EWIDENCJI CZASU PRACY"))
            {
                return 3;
            }

            Cell_Value = Worksheet.Cell(1, 1).Value.ToString();
            if (Cell_Value.Trim().StartsWith("GRAFIK PRACY MIESIĄC")) // grafik v2024 v2
            {
                return 5;
            }

            Cell_Value = Worksheet.Cell(1, 2).Value.ToString();
            if (Cell_Value.Trim().StartsWith("GRAFIK PRACY MIESIĄC")) // grafik v2024 v2
            {
                return 5;
            }

            foreach (IXLCell cell in Worksheet.CellsUsed()) // karta pracy NIE konduktora
            {
                try
                {
                    if (cell.GetString().Trim() == "Dzień")
                    {
                        return 4;
                    }
                }
                catch { }
            }

            return 0;
        }
        private static async Task Usun_Ukryte_Karty(XLWorkbook workbook)
        {
            await Task.Run(() =>
            {
                List<IXLWorksheet> hiddenSheets = [];
                foreach (IXLWorksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Visibility == XLWorksheetVisibility.Hidden)
                    {
                        hiddenSheets.Add(sheet);
                    }
                }
                foreach (IXLWorksheet sheet in hiddenSheets)
                {
                    workbook.Worksheets.Delete(sheet.Name);
                }
                workbook.Save();
            });
        }
        private static async Task<(string, DateTime)> Get_Metadane_Pliku(string File_Path)
        {
            DateTime lastWriteTime = File.GetLastWriteTime(File_Path);

            try
            {
                using (FileStream fs = new(File_Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var workbook = await Task.Run(() => new XLWorkbook(fs));

                    string lastModifiedBy = workbook.Properties.LastModifiedBy ?? "";
                    return (lastModifiedBy, lastWriteTime);
                }
            }
            catch
            {
                return ("", lastWriteTime);
            }
        }
        private static bool Is_File_Valid(string File_Path)
        {
            try
            {
                XLWorkbook Workbook = new(File_Path);
            }
            catch
            {
                MoveFile(File_Path, 1).ConfigureAwait(false);
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"Program nie może odczytać pliku {File_Path}");
                Console.ForegroundColor = ConsoleColor.White;
                return false;
            }
            string[] validExtensions = [".xlsx", ".xls"];
            string fileExtension = Path.GetExtension(File_Path).ToLowerInvariant();
            if (!validExtensions.Contains(fileExtension))
            {
                MoveFile(File_Path, 1).ConfigureAwait(false);
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"Program nie może odczytać pliku {File_Path}");
                Console.ForegroundColor = ConsoleColor.White;
                return false;
            }
            return true;
        }
        private static async Task MoveFile(string filePath, int option)
        {
            if (!config.Move_Files_To_Processed_Folder)
            {
                return;
            }
            string processedFilesFolder = string.Empty;
            processedFilesFolder = option switch
            {
                0 => error_logger.Current_Processed_Files_Folder,
                1 => error_logger.Current_Bad_Files_Folder,
                _ => error_logger.Current_Processed_Files_Folder,
            };
            Directory.CreateDirectory(processedFilesFolder);
            string destinationPath = Path.Combine(processedFilesFolder, Path.GetFileName(filePath));
            if (File.Exists(destinationPath))
            {
                File.Delete(destinationPath);
            }
            await Task.Run(() => File.Move(filePath, destinationPath));
        }
        private static async Task Copy_Bad_Sheet_To_Files_Folder(string filePath, int sheetIndex)
        {
            string newFilePath = Path.Combine(error_logger.Current_Bad_Files_Folder, "DO_POPRAWY_" + Path.GetFileName(filePath));
            try
            {
                await Task.Run(() =>
                {
                    using (XLWorkbook originalwb = new(filePath))
                    {
                        IXLWorksheet sheetToCopy = originalwb.Worksheet(sheetIndex);
                        string newSheetName = sheetToCopy.Name;
                        if (newSheetName.Length > 31)
                        {
                            newSheetName = newSheetName[..31];
                        }

                        using (XLWorkbook workbook = File.Exists(newFilePath) ? new XLWorkbook(newFilePath) : new XLWorkbook())
                        {
                            if (workbook.Worksheets.Contains(newSheetName))
                            {
                                return;
                            }
                            sheetToCopy.CopyTo(workbook, newSheetName);
                            XLWorkbookProperties properties = originalwb.Properties;
                            properties.Author = "Copied by program";
                            properties.Modified = DateTime.Now;
                            workbook.SaveAs(newFilePath);
                        }
                    }
                });
            }
            catch
            {
            }
        }
        private static void Check_Base_Dirs(string path)
        {
            string[] directories =
            [
                Path.Combine(path, "Errors"),
                Path.Combine(path, "Bad_Files"),
                Path.Combine(path, "Processed_Files")
            ];
            string errorFilePath = Path.Combine(directories[0], "Errors.txt");
            error_logger.Current_Processed_Files_Folder = directories[2];
            error_logger.Current_Bad_Files_Folder = directories[1];
            error_logger.Set_Error_File_Path(directories[0]);
            foreach (var dir in directories)
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

            if (!File.Exists(errorFilePath))
            {
                File.Create(errorFilePath);
            }
        }
        public static async Task Convert_To_Xlsx(string inputFilePath, string outputFilePath)
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
            (string o, DateTime d) = await Get_Metadane_Pliku(inputFilePath).ConfigureAwait(false);
            workbook.Properties.LastModifiedBy = o;
            workbook.Properties.Modified = d;
            workbook.SaveAs(outputFilePath);
        }
    }
}

