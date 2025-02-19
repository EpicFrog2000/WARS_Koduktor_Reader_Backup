namespace Excel_Data_Importer_WARS
{
    internal class Error_Logger
    {
        public Error_Logger(bool showmsg)
        {
            ShowErrorMessageOnWrite = showmsg;
        }

        private bool ShowErrorMessageOnWrite = true;

        // Plik excel na którym obecnie wykonwywane są operacje
        public string Nazwa_Pliku = "";

        // Zakładka na której wystąpił błąd
        public int Nr_Zakladki = 0;

        public string Nazwa_Zakladki = "";

        // Obecna wartość pola z błędem
        private string Wartosc_Pola = "";

        // Nazwa tego co powinno znaleźć się w tym polu
        private string Poprawna_Wartosc_Pola = "";

        // Scierzka do pliku w którym maja być zapisywane błędy
        public string ErrorFilePath = "";

        // Kolumna w której wystąpił błąd
        public int Kolumna = -1;

        // Kolumna w której wystąpił błąd
        public int Rzad = -1;

        // Dodatkowa wiadomośc na koncu errora w pliku
        private string OptionalMsg = "";

        // Osoba która ostatnio zedytowała dane
        public string Last_Mod_Osoba = "";

        public DateTime Last_Mod_Time = DateTime.Now;

        // Nazwa obecnie przetwarzanych plików
        public string Current_Processed_Files_Folder = "";

        public string Current_Bad_Files_Folder = "";

        /// <summary>
        /// Tworzy wiadomość z podanych parametrów i dodaje wiadomość o błędzie do pliku z errorami.
        /// </summary>
        public void New_Error(string? wartoscPola = "", string? nazwaPola = "", int kolumna = -1, int rzad = -1, string? optionalmsg = "")
        {
            Poprawna_Wartosc_Pola = nazwaPola!;
            Wartosc_Pola = wartoscPola!;
            Kolumna = kolumna;
            Rzad = rzad;
            OptionalMsg = optionalmsg!;
            Append_Error_To_File();
            if (ShowErrorMessageOnWrite)
            {
                Console.WriteLine(Get_Error_String());
            }
        }

        /// <summary>
        /// Zwraca wiadomość jaką wpisało by do pliku z errorami.
        /// </summary>
        /// <returns>Zwraca wiadomość jaką wpisało by do pliku z errorami.</returns>
        public string Get_Error_String()
        {
            string Wiadomosc = "-------------------------------------------------------------------------------";
            if (!string.IsNullOrEmpty(Nazwa_Pliku))
            {
                Wiadomosc += Environment.NewLine + "Plik: " + System.IO.Path.GetFileName(Nazwa_Pliku);
            }
            Wiadomosc += Environment.NewLine + "Nazwa zakladki: " + Nazwa_Zakladki;
            if (Nr_Zakladki != 0)
            {
                Wiadomosc += Environment.NewLine + "Nr Zakladki: " + Nr_Zakladki;
            }
            if (Kolumna != -1)
            {
                Wiadomosc += Environment.NewLine + "Kolumna: " + Kolumna;
            }
            if (Rzad != -1)
            {
                Wiadomosc += Environment.NewLine + "Rzad: " + Rzad;
            }
            Wiadomosc += Environment.NewLine + "Wartość w komórce: '" + Wartosc_Pola + "'";
            if (!string.IsNullOrEmpty(Poprawna_Wartosc_Pola))
            {
                Wiadomosc += Environment.NewLine + "Poprawna wartość jaka powinna być: " + Poprawna_Wartosc_Pola;
            }
            if (!string.IsNullOrEmpty(OptionalMsg))
            {
                Wiadomosc += Environment.NewLine + "Dodatkowa wiadomość: " + OptionalMsg;
            }
            Wiadomosc += Environment.NewLine + "-------------------------------------------------------------------------------" + Environment.NewLine;
            return Wiadomosc;
        }

        /// <summary>
        /// Wpisuje do pliku z errorami wiadomość z parametru.
        /// </summary>
        public void New_Custom_Error(string Error_Msg)
        {
            Error_Msg = "-------------------------------------------------------------------------------" + Environment.NewLine + Error_Msg + Environment.NewLine + "-------------------------------------------------------------------------------" + Environment.NewLine;
            Append_Error_To_File(Error_Msg);
            if (ShowErrorMessageOnWrite)
            {
                Console.WriteLine(Error_Msg);
            }
        }

        public void Set_Error_File_Path(string New_Error_File_Path)
        {
            ErrorFilePath = New_Error_File_Path;
        }

        private void Append_Error_To_File()
        {
            if (ErrorFilePath == "") { throw new Exception("ErrorLogger nie posiada właściwej scierzki do pliku Errors.txt"); }
            string ErrorsLogFile = Path.Combine(ErrorFilePath, "Errors.txt");
            if (!File.Exists(ErrorsLogFile))
            {
                File.Create(ErrorsLogFile).Dispose();
            }
            File.AppendAllText(ErrorsLogFile, Get_Error_String() + Environment.NewLine);
        }

        private void Append_Error_To_File(string Error_Msg)
        {
            if (ErrorFilePath == "") { throw new Exception("ErrorLogger nie posiada właściwej scierzki do pliku Errors.txt"); }
            string ErrorsLogFile = Path.Combine(ErrorFilePath, "Errors.txt");
            if (!File.Exists(ErrorsLogFile))
            {
                File.Create(ErrorsLogFile).Dispose();
            }
            File.AppendAllText(ErrorsLogFile, Error_Msg + Environment.NewLine);
        }
    }
}