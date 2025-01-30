using System.Linq;
using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;


namespace Konduktor_Reader{
    static class Program
    {
        public static Error_Logger error_logger = new(true); // true - Console write message on creating new error
        private static string[] Path_To_Folders_With_Files = ["C:\\Users\\norbert.tasarz\\Desktop\\Arkusze konduktorzy\\Wynagrodzenia1\\"];
        public static readonly DateTime baseDate = new(1899, 12, 30); // Do zapytan sql
        public static string Optima_Conection_String = "Server=ITEGERNT;Database=CDN_Wars_prod_ITEGER;Encrypt=True;TrustServerCertificate=True;Integrated Security=True;\r\n";

        public static int Main()
        {
            Config Config = new();
            Config.GetConfigFromFile();
            // SET VARIABLES TO THOSE FROM CONFIG
            Optima_Conection_String = Config.Optima_Conection_String;
            // itd...


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

                foreach (string File_Path in Files_Paths)
                {
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine($"Czytanie: {Path.GetFileNameWithoutExtension(File_Path)} {DateTime.Now}");
                    Console.ForegroundColor = ConsoleColor.White;
                    if (!Is_File_Valid(File_Path))
                    {
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine($"Program nie możę odczytać pliku {File_Path}");
                        Console.ForegroundColor = ConsoleColor.White;
                        continue;
                    }
                    // TODO konwersja pliku na xlsx czy coś
                    error_logger.Nazwa_Pliku = File_Path;
                    (error_logger.Last_Mod_Osoba, error_logger.Last_Mod_Time) = Get_Metadane_Pliku(File_Path);
                    using (XLWorkbook Workbook = new(File_Path))
                    {
                        Usun_Ukryte_Karty(Workbook);
                        int Ilosc_Zakladek_W_Workbook = Workbook.Worksheets.Count;
                        if(Ilosc_Zakladek_W_Workbook < 1)
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
                                case 1: // Reader_Harmonogram_v1
                                    try
                                    {
                                        Reader_Harmonogram_v1.Process_Zakladka(Zakladka);
                                    }
                                    catch
                                    {
                                        Copy_Bad_Sheet_To_Files_Folder(File_Path, Obecny_Numer_Zakladki);
                                    }
                                    break;
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
                                default:
                                    Copy_Bad_Sheet_To_Files_Folder(File_Path, Obecny_Numer_Zakladki);
                                    error_logger.New_Custom_Error($"Nie rozpoznano tego typu zakładki w pliku: \"{error_logger.Nazwa_Pliku}\" zakladka: \"{error_logger.Nazwa_Zakladki}\" numer zakładki: \"{error_logger.Nr_Zakladki}\"");
                                    continue;
                            }
                        }
                    }
                    //MoveFile(File_Path);
                }
            }
            return 0;
        }
        private static int Get_Typ_Zakladki(IXLWorksheet Worksheet)
        {
            string Cell_Value = Worksheet.Cell(3,5).GetFormattedString().Trim().Replace("  ", " ");
            if(Cell_Value.Contains("Harmonogram pracy"))
            {
                return 1;
            }

            Cell_Value = Worksheet.Cell(1, 3).GetFormattedString().Trim().Replace("  ", " ");
            if (Cell_Value.Contains("Tabela Stawek"))
            {
                return 2;
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
                using (XLWorkbook workbook = new XLWorkbook(File_Path))
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
        private static bool Is_File_Valid(string filePath)
        {
            string[] validExtensions = { ".xlsb", ".xlsx", ".xls" };
            string fileExtension = Path.GetExtension(filePath).ToLowerInvariant();
            return validExtensions.Contains(fileExtension);
        }
        private static void MoveFile(string filePath)
        {
            string processedFilesFolder = error_logger.Current_Processed_Files_Folder;
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
                        newSheetName = newSheetName.Substring(0, 31);
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
            {
                Path.Combine(path, "Errors"),
                Path.Combine(path, "Bad_Files"),
                Path.Combine(path, "Processed_Files")
            };
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
            if (!File.Exists(errorFilePath))
            {
                File.Create(errorFilePath);
            }
        }
    }
}

