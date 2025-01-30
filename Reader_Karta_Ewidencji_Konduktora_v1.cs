using ClosedXML.Excel;

namespace Konduktor_Reader
{
    internal static class Reader_Karta_Ewidencji_Konduktora_v1
    {
        // JA JEBE KURWA PRZECIEZ TO BĘDZIE ZMIENIANE Z 500 GORYLIONÓW RAZY WSZYSTKO PORA SIE ZAJEBAĆ
        private class Karta_Ewidencji
        {
            public int Rok = 0;
            public int Miesiac = 0;
            public Pracownik Pracownik = new();
            public List<Dane_Karty> Dane_Karty = [];
        }
        private class Dane_Karty
        {
            public Relacja Relacja = new();
            public List<Dane_Dnia> Dane_Dni_Relacji = [];
        }
        private class Dane_Dnia
        {
            public TimeSpan Godziny_Pracy_Od = TimeSpan.Zero;
            public TimeSpan Godziny_Pracy_Do = TimeSpan.Zero;
            public TimeSpan Godziny_Odpoczynku_Od = TimeSpan.Zero;
            public TimeSpan Godziny_Odpoczynku_Do = TimeSpan.Zero;
            public int Liczba_Godzin_Nadliczbowych_50 = 0;
            public int Liczba_Godzin_Nadliczbowych_100 = 0;
            public int Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 = 0;
            public int Liczba_Godzin_Nadliczbowych_Ryczalcie_100 = 0;
            public string Absencja_Nazwa = string.Empty;
            public int Liczba_Godzin_Absencji = 0;
        }
        public static void Process_Zakladka(IXLWorksheet Zakladka)
        {
            List<Karta_Ewidencji> Karty_Ewidencji = [];
            List<Helper.Current_Position> Pozycje = Helper.Find_Staring_Points_Tabele_Stawek(Zakladka, "Dzień miesiąca");
            foreach (Helper.Current_Position Pozycja in Pozycje)
            {
                Karta_Ewidencji Karta_Ewidencji = new();

                // TODO Get dane

                Karty_Ewidencji.Add(Karta_Ewidencji);
            }

            foreach (Karta_Ewidencji Karta_Ewidencji in Karty_Ewidencji)
            {
                // TODO Dodaj do bazy
            }
        }
        private static void Get_Dane_Naglowka()
        {
            // data, i pracownik
        }
    }
}
