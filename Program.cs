using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Diagnostics;

namespace Konduktor_Reader{
    static class Program
    {
        public static Error_Logger error_logger = new(true); // true - Console write message on creating new error
        public static string[] Path_To_Folders_With_Files = ["C:\\Users\\norbert.tasarz\\Desktop\\Arkusze konduktorzy\\Wynagrodzenia1\\", "C:\\Users\\norbert.tasarz\\Desktop\\Arkusze konduktorzy\\Ewidencja 2\\"];
        public static readonly DateTime baseDate = new(1899, 12, 30); // Do zapytan sql
        public static string Optima_Conection_String = "Server=ITEGERNT;Database=CDN_Wars_prod_ITEGER_22012025;Encrypt=True;TrustServerCertificate=True;Integrated Security=True;";
        public static bool Clear_Logs_On_Program_Restart = false;
        public static bool Clear_Processed_Files_On_Restart = false;
        public static bool Clear_Bad_Files_On_Restart = false;
        public static bool Move_Files_To_Processed_Folder = false;
        public static readonly bool TEST_CLEAR_DB_TABLES = true;
        public static readonly bool LOG_TO_Terminal = true;
        public static int Main()
        {
            Stopwatch stopwatch = new Stopwatch();

            // Start measuring time
            stopwatch.Start();
            TextWriter originalOut = Console.Out;
            if (!LOG_TO_Terminal)
            {
                Console.SetOut(TextWriter.Null);
            }
            Clear_Tables(); // FOR TESTING ONLY


            // TODO DODAC PETLE WHILE TRUE

            Config Config = new();
            Config.GetConfigFromFile();
            Config.Set_Program_Config();

            if (!Helper.Valid_SQLConnection_String(Optima_Conection_String))
            {
                Console.WriteLine($"Invalid connection string: {Optima_Conection_String}");
                Console.ReadLine();
                return -1;
            }

            if (Path_To_Folders_With_Files.Length < 1)
            {
                Console.WriteLine($"Program nie ma ustawionych folderów");
                return 0;
            }


            foreach (string Folder_Path in Path_To_Folders_With_Files)
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
                    string File_Path = filepath;
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine($"Czytanie: {Path.GetFileNameWithoutExtension(File_Path)} {DateTime.Now}");
                    Console.ForegroundColor = ConsoleColor.White;
                    if (!Is_File_Valid(File_Path))
                    {
                        //try
                        //{
                        //    Convert_To_Xlsx(File_Path, Path.Combine(Path.GetDirectoryName(File_Path)!, Path.GetFileNameWithoutExtension(File_Path) + ".xlsx"));
                        //    File_Path = Path.Combine(Path.GetDirectoryName(File_Path)!, Path.GetFileNameWithoutExtension(File_Path) + ".xlsx");
                        //}
                        //catch
                        //{
                        //    MoveFile(File_Path, 1);
                        //    Console.ForegroundColor = ConsoleColor.DarkYellow;
                        //    Console.WriteLine($"Program nie możę odczytać pliku {File_Path}");
                        //    Console.ForegroundColor = ConsoleColor.White;
                        //    continue;
                        //}
                        MoveFile(File_Path, 1);
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine($"Program nie możę odczytać pliku {File_Path}");
                        Console.ForegroundColor = ConsoleColor.White;
                        continue;
                    }
                    error_logger.Nazwa_Pliku = File_Path;
                    (error_logger.Last_Mod_Osoba, error_logger.Last_Mod_Time) = Get_Metadane_Pliku(File_Path);
                    using (XLWorkbook Workbook = new(File_Path))
                    {
                        Usun_Ukryte_Karty(Workbook);
                        int Ilosc_Zakladek_W_Workbook = Workbook.Worksheets.Count;
                        if (Ilosc_Zakladek_W_Workbook < 1)
                        {
                            continue;
                        }
                        for (int Obecny_Numer_Zakladki = 1; Obecny_Numer_Zakladki <= Ilosc_Zakladek_W_Workbook; Obecny_Numer_Zakladki++)
                        {
                            error_logger.Nr_Zakladki = Obecny_Numer_Zakladki;
                            IXLWorksheet Zakladka = Workbook.Worksheet(Obecny_Numer_Zakladki);
                            error_logger.Nazwa_Zakladki = Zakladka.Name;
                            int Typ_Zakladki = Get_Typ_Zakladki(Zakladka);
                            switch (Typ_Zakladki)
                            {

                                case 2:
                                    try
                                    {
                                        Reader_Tabela_Stawek_v1.Process_Zakladka(Zakladka);
                                    }
                                    catch
                                    {
                                        Copy_Bad_Sheet_To_Files_Folder(File_Path, Obecny_Numer_Zakladki);
                                    }
                                    break;
                                case 3:
                                    try
                                    {
                                        Obecny_Numer_Zakladki = Ilosc_Zakladek_W_Workbook+1; // Czytanie tylko pierwszej zakłdki
                                        Reader_Karta_Ewidencji_Konduktora_v1.Process_Zakladka(Zakladka);
                                    }
                                    catch
                                    {
                                        Copy_Bad_Sheet_To_Files_Folder(File_Path, Obecny_Numer_Zakladki);
                                    }
                                    break;
                                default:
                                    Copy_Bad_Sheet_To_Files_Folder(File_Path, Obecny_Numer_Zakladki);
                                    error_logger.New_Custom_Error($"Nie rozpoznano tego typu zakładki w pliku: \"{error_logger.Nazwa_Pliku}\" zakladka: \"{error_logger.Nazwa_Zakladki}\" numer zakładki: \"{error_logger.Nr_Zakladki}\"");
                                    continue;
                            }
                        }
                    }
                    MoveFile(File_Path, 0);
                }
            }
            if (!LOG_TO_Terminal)
            {
                Console.SetOut(originalOut);
            }
            TimeSpan elapsedTime = stopwatch.Elapsed;
            Console.WriteLine("Elapsed Time: " + elapsedTime.ToString(@"hh\:mm\:ss\:fff"));

            Console.WriteLine("Kliknij aby zakończyć...");
            Console.ReadLine();
            return 0;
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
            return 0;
        }
        private static void Usun_Ukryte_Karty(XLWorkbook workbook)
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
        }
        private static (string, DateTime) Get_Metadane_Pliku(string File_Path)
        {
            try
            {
                using (XLWorkbook workbook = new(File_Path))
                {
                    DateTime lastWriteTime = File.GetLastWriteTime(File_Path);

                    if (workbook.Properties.LastModifiedBy == null)
                    {
                        return ("", lastWriteTime);
                    }
                    return (workbook.Properties.LastModifiedBy, lastWriteTime);
                }
            }
            catch
            {
                DateTime lastWriteTime = File.GetLastWriteTime(File_Path);
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
                return false;
            }
            string[] validExtensions = [".xlsx", ".xls"];
            string fileExtension = Path.GetExtension(File_Path).ToLowerInvariant();
            return validExtensions.Contains(fileExtension);
        }
        private static void MoveFile(string filePath, int option)
        {
            if (!Move_Files_To_Processed_Folder)
            {
                return;
            }
            string processedFilesFolder = string.Empty;
            switch (option)
            {
                case 0:
                    processedFilesFolder = error_logger.Current_Processed_Files_Folder;
                    break;
                case 1:
                    processedFilesFolder = error_logger.Current_Bad_Files_Folder;
                    break;
                default:
                    processedFilesFolder = error_logger.Current_Processed_Files_Folder;
                    break;
            }
            Directory.CreateDirectory(processedFilesFolder);
            string destinationPath = Path.Combine(processedFilesFolder, Path.GetFileName(filePath));
            if (File.Exists(destinationPath))
            {
                File.Delete(destinationPath);
            }
            File.Move(filePath, destinationPath);
        }
        private static void Copy_Bad_Sheet_To_Files_Folder(string filePath, int sheetIndex)
        {
            string newFilePath = System.IO.Path.Combine(error_logger.Current_Bad_Files_Folder, "DO_POPRAWY_" + System.IO.Path.GetFileName(filePath));
            try
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

            if (Clear_Logs_On_Program_Restart)
            {
                foreach (string file in Directory.GetFiles(directories[0]))
                {
                    File.Delete(file);
                }
            }
            if (Clear_Bad_Files_On_Restart)
            {
                foreach (string file in Directory.GetFiles(directories[1]))
                {
                    File.Delete(file);
                }
            }
            if (Clear_Processed_Files_On_Restart)
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

            using XLWorkbook workbook = new XLWorkbook();
            foreach (System.Data.DataTable table in dataSet.Tables)
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
            (string o, DateTime d) = Get_Metadane_Pliku(inputFilePath);
            workbook.Properties.LastModifiedBy = o;
            workbook.Properties.Modified = d;
            workbook.SaveAs(outputFilePath);
        }

        private static void Clear_Tables()
        {
            if (TEST_CLEAR_DB_TABLES)
            {
                using SqlConnection connection = new(Optima_Conection_String);
                connection.Open();
                using SqlCommand command = new("delete from cdn.PracPracaDniGodz; delete from cdn.PracPracaDni", connection);
                command.ExecuteScalar();
                connection.Close();
            }
        }
    }
}

