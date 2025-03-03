# Excel Data Importer WARS
#### Do czego służy?
Program służy do automatycznego importowania danych z plików Excel do bazy danych. Dzięki niemu nie trzeba ręcznie przepisywać danych – wystarczy umieścić plik w odpowiednim folderze, a program zajmie się resztą. Jego zadaniem jest odczytanie danych z plików Excel umieszczonych w wyznaczonym folderze, zweryfikowanie ich poprawności oraz zapisanie ich w bazie danych lub obsłużenie ewentualnych błędów, które mogą wystąpić podczas przetwarzania plików. Jeśli napotka błędy, rejestruje je, aby można było je łatwo poprawić. Program obsługuje różne typy dokumentów związanych z ewidencją czasu pracy i wynagrodzeniami pracowników. 

#### Jakie są wyniki programu?
[Zobacz dokument](Dokumentacja.docx)


#### Jakie dane przetwarza?

Grafiki pracy pracowników – harmonogramy pracy, w tym zaplanowane absencje.

Karty ewidencji czasu pracy – rzeczywiste godziny pracy pracowników i konduktorów, w tym nadgodziny, niobecności oraz inne elementy związane z czasem pracy.

Tabele stawek wynagrodzeń – stawki płacowe obowiązujące dla różnych grup pracowników i okresów rozliczeniowych.

#### Funkcjonalności:

✅ Automatyczne wykrywanie nowych plików Excel w określonym folderze

✅ Obsługa różnych formatów arkuszy

✅ Sprawdza poprawność danych przed ich zapisaniem

✅ Integracja z bazą danych

✅ Rejestruje błędy, jeśli plik zawiera nieprawidłowe informacje

✅ Segregacja przetworzonych plików

✅ Łatwa zmiana opcji programu poprzez edycję pliku .json


## Użytkowanie
Plik konfiguracyjny `Config.json` znajduje się w tej samej lokalizacji co program. W tym pliku można skonfigurować działanie aplikacji poprzez ustawienie odpowiednich parametrów:

Przykładowy plik konfiguracyjny:
```
{
  "Files_Folders": [
    "C:\\Users\\norbert.tasarz\\Desktop\\pojedyncze pliki\\"
  ],
  "Optima_Conection_String": "Server=ITEGERNT;Database=CDN_Wars_prod_ITEGER_22012025;Encrypt=True;TrustServerCertificate=True;Integrated Security=True;",
  "Clear_Logs_On_Program_Restart": true,
  "Clear_Processed_Files_On_Restart": true,
  "Clear_Bad_Files_On_Restart": true,
  "Move_Files_To_Processed_Folder": false,
  "Clear_Good_Files_On_Restart": true
}
```
### Konfiguracja katalogów
- **`Files_Folders`** - lista folderów, z których program importuje pliki Excel.

### Połączenie z bazą danych
- **`Optima_Conection_String`** - ciąg połączeniowy do bazy danych.

### Zarządzanie plikami i logami
- **`Clear_Logs_On_Program_Restart`** - określa, czy logi z błędami powinny być kasowane po restarcie programu.
- **`Clear_Processed_Files_On_Restart`** - określa, czy przetworzone pliki powinny być kasowane po restarcie programu.
- **`Clear_Bad_Files_On_Restart`** - określa, czy niepoprawne pliki powinny być usuwane po restarcie programu.
- **`Clear_Good_Files_On_Restart`** - określa, czy poprawnie przetworzone pliki powinny być usuwane po restarcie programu.
- **`Move_Files_To_Processed_Folder`** - określa, czy pliki powinny być przenoszone do dedykowanego folderu po ich przetworzeniu.

### Obsługa błędów
Pliki do importu należy przekopiować do folderów, które zostały określone w pliku konfiguracyjnym (**`Files_Folders`**). 
Po uruchomieniu program utworzy nowe katalogi, np. `Errors\` i `Bad_Files\`.
- W folderze `Errors\` będzie znajdował się plik `Errors.txt` z opisanymi błędami, które wystąpiły w trakcie importu plików.
- W folderze `Bad_Files\` będą znajdowały się pliki, które nie zostały poprawnie wczytane (np. karty ewidencji, grafiki pracy itp.).
- W folderze `Good_Files\` znajdowały się pliki, które wczytały się w 100% bezbłędnie. 
- W folderze `Processed_Files\` będą znajdowały się pliki na których program przeprowadzał operacje.

Zgodnie z błędami opisanymi w pliku `Errors\Errors.txt` należy poprawić błędy w plikach znajdujących się w folderze `Bad_Files\`. 
Poprawione pliki można następnie przekopiować do folderów określonych w pliku konfiguracyjnym (**`Files_Folders`**) i ponownie uruchomić program.

- Błędy są opisywane na 2 sposoby:
  1. Strukturalny: opisane pola z wartościami między liniami np:
   ```
   -------------------------------------------------------------------------------
    Plik: Południe - zespół magazynowy- listopad.xlsx
    Nazwa zakladki: KATARZYNA FUNDAMENT
    Nr Zakladki: 13
    Kolumna: 5
    Rzad: 33
    Wartość w komórce: 'eqw45'
    Poprawna wartość jaka powinna być: Rodzaj Absencji
    Dodatkowa wiadomość: Nierozpoznany rodzaj absencji
    -------------------------------------------------------------------------------
  ```
  Oznacza to że:

  Wystąpił błąd w pliku o nazwie: `Południe - zespół magazynowy- listopad.xlsx`

  Konkretnie w zakładce o nazwie: `KATARZYNA FUNDAMENT`

  Lub numerze zakładki: `13`

  Błąd zlokalizowano w kolumnie `5` i rzędzie `33`

  Wartosć jaka jest obecnie w tym miejscu: `UŻ`

  Wartość oczekiwana przez program do wczytania: `Rodzaj Absencji`

  Oraz często znajdzie się dodatkowa wiadomość opisująca dlaczego wystąpił błąd. W tym przypadku: `Nierozpoznany rodzaj absencji`

  W takim przypadku należy zamienić wartość w tej komórce na poprawną wartość jaka powinna się tam znaleźć czyli kod absencji np: `PZ`
  
  2. Opisowy: W 1-2 Zdaniach opisany jest błąd który wystąpił w trakcie wczytywania danych:
  ```
  -------------------------------------------------------------------------------
  Zły format pliku, nie znaleniono żadnych grafików z pliku: Południe - zespół magazynowy- listopad.xlsx z zakladki: 1 nazwa zakladki: Plan1
  -------------------------------------------------------------------------------
  ```
  Oznacza to że:
  
  Zawartość zakładki nie jest zgodna z przewidzianym standardem.






  
# Kod

## Testowanie

Zalecam skompilowanie programu, a następnie przejście do folderu, w którym znajduje się skompilowana wersja. Uruchom program na 5 sekund, aby automatycznie wygenerował plik konfiguracyjny.

Następnie edytuj plik konfiguracyjny zgodnie z własnymi potrzebami, korzystając z powyższych wskazówek.

Do folderów określonych w polu `Files_Folders` dodaj odpowiednie pliki Excel.

I gotowe – maszyna w ruchu! Można odpalić program.


## Struktura Projektu
- `Program.cs`: Główny plik uruchamiający aplikację.
- `Config.cs`: Plik z klasą pobierającą i ustawiającą sposób działania programu zgodnie z plikiem konfiguracyjnym Config.json.
- `Error_Logger.cs`: Klasa do logowania błędów.
- `DbMaganer.cs`: Klasa przechowująca informacje potrzebne do operacji na bazie danych.
- `Helper.cs`: Klasa Helper zawiera metody pomocnicze do pracy z arkuszami Excel itp.
- `Reader_*.cs`: Klasy czytające pliki Excel oraz wkładające te dane do bazy.

# Dokumentacja Plików z kodem
### Program.cs
- `private static readonly bool LOG_TO_TERMINAL`: Określa czy w trakcie działania programu mają się pojawiać komunikaty w konsoli.
- `private static readonly bool Do_Stuff_In_loop`: Określa czy program ma wykonywać swoje założenia nieprzerwanie.
- `static class Pomiar`: Klasa do badania optymalizacji i wydajności czasowej programu. (Można usunąć)
- `Main()`: Główna metoda uruchamiająca aplikację.
- `Do_The_Thing()`: Metoda wykonująca główne operacje.
- `private static int Get_Typ_Zakladki(IXLWorksheet Worksheet)`: Zwraca typ zakładki. Każdy typ zakładki jest reprezentowany przez unikalny numer. Zakładka tzn. który typ danych jest w tem skoroszycie np. Grafiki pracy lub karty ewidencji konduktorów.
- `private static Task Usun_Ukryte_Karty(XLWorkbook workbook)`: Usuwa z pliku Excel zakładki które zostały ukryte przez użytkownika gdyż najprawdopodobniej nie mają one być importowane skoro ma ich nie widać xd. (Chce się zajebać)
- `private static (string, DateTime) Get_Metadane_Pliku(XLWorkbook Workbook, string File_Path)`: Zwraca ostatni czas oraz osobę modyfikacji danego pliku.
- `private static (XLWorkbook?, FileStream?) Open_Workbook(string File_Path)` otwiera plik Excel przez FileStream i zwraca handle na FileStream oraz XLWorkbook, które to należy zamknąć na koniec. (Zrobione jest to tak aby zoptymalizować użycie zasobów)
- `private static void Move_File(string filePath, int opcja)`: W zależności od opcji przenosi plik do odpowiednich folderów. (0 - Bad_Files\, 1 - Good_Files\, 2 - Processed_Files\)
- `private static async Task Copy_Bad_Sheet_To_Files_Folder(XLWorkbookProperties op, IXLWorksheet sheetToCopy, string filePath)`: Tworzy kopię przekazanej zakładki, która zawiera błąd, tę kopię dodaje do folderu Bad_Files\.
- `private static void Check_Base_Dirs(string path)`: Tworzy odpowiednie podfoldery jeśli nie instnieją i/lub czyści ich zawarość jeśli taka opcja została przekazana programowi.

Działanie programu: 

  0. Zmienna `LOG_TO_TERMINAL` mówi o tym czy w terminalu mają się pojawiać logi ale w końcowej fazie i tak program będzi ebez żądnego interfejsu wiec JD
  1. Wczytywane są opcje z pliku konfiguracyjnego: `config.GetConfigFromFile();`
  2. Testowane jest połączenie z bazą `DbManager.Valid_SQLConnection_String()`
  3. Tworzone są odpowiednie foldery w katalogach ustawionych z pliku konfiguracyjnego: `Check_Base_Dirs(Folder_Path)`
  4. Dla każdego pliku następuje `Process_Files(string File_Path)` w którym to:
  - Otwierany jest Arkusz
  - Usuwane są poukrywane zakładki: `Usun_Ukryte_Karty(Workbook)`
  - Tworzony jest error logger do wczytywania zakładek
  - Sprawdzany jest typ zakładki `Get_Typ_Zakladki(Zakladka)`
  - W zależności od typu zakładki wykonuje się odpowiedni proces wczytywania danych z arkusza: `Process_Zakladka()`
  - Plik i/lub odpowiednia zakładka jest przenoszona do odowiednich folderów w zależności od jej poprawności: `Move_File()`, `Copy_Bad_Sheet_To_Files_Folder()` 
  - Arkusz jes zamykany
  5. Operacje te wykonywane są w pętli z wależności od zmiennej `Do_Stuff_In_loop`

     
### Error_Logger.cs
- `public Error_Logger(bool showmsg)`: przy tworzeniu obiektu parametr showmsg określa czy w trakcie wykrycia nowego błędu program ma wypisać go w terminalu.
- `public void New_Error(string? wartoscPola = "", string? nazwaPola = "", int kolumna = -1, int rzad = -1, string? optionalmsg = "")`: Tworzy i zapisuje nowy błąd do pliku Errors.txt wg przekazanych parametrów.
- `public string Get_Error_String()`: Zwraca Wiadomość błędu utworzoną przez ostatnie użycie metody New_Error().
- `public void New_Custom_Error(string Error_Msg)`: Zapisuje przekazaną wiadomość jako nowy błąd do pliku Errors.txt.
- `public void Set_Error_File_Path(string New_Error_File_Path)`: Ustawia scierzkę do pliku do którego mają być zapisywanie błędy.
- `private void Append_Error_To_File()`: Dodaje tekst do pliku z błędami

### Config.cs
- `public void GetConfigFromFile()`: Jeśli nie istnieje plik konfiguracyjny to go tworzy a następnie pobiera z niego opcje i ustawia zmienne w programie.
- `public void Check_File()`: Tworzy plik konfiguracyjny jeśli nie egzystuje.

### DbManager.cs
- `public static bool Valid_SQLConnection_String()`: Zwraca czy połączenie z bazą jest prawidłowe z ustawonym connection stringiem.

### Helper.cs
- `public static bool Try_Get_Type_From_String<T>(string? value, ref T result)`: Ustawia wartość w podaej zmiennej a <T> jeśłi value jest <T>, w przeciwnm wypadku zwraca fałsz.
- `public static bool Try_Get_Type_From_String<T>(string? value, Action<T> method)`: Uzywa podajen metody (w zamyśle settera) do ustawienia zmiennej.
- `public static List<Current_Position> Find_Starting_Points(IXLWorksheet worksheet, string keyWord, bool CompareMode=true)`: Znajduje punkty odniesienia do czytywanych kart. Szuka w podajen zakładce komóek które mają w obie keyword i zwraca listę znalezionych komórek. CompareMode=t==rue -> Contains(keyword), CompareMode=false -> == keyword.
- `public static string Truncate(string? value, int maxLength)`: Przycina string do maxLength wartości jeśli jest większy.

### Absencja.cs
- `public static int Dodaj_Absencje_do_Optimy()`: Dodaje dane do bazy danych.
- `private static List<List<Absencja>> Podziel_Absencje_Na_Osobne(List<Absencja> Absencje)`: Grupuje nieobecności na podstawie ich kodu oraz czy występuję dzień po dniu.
- `private static int Ile_Dni_Roboczych(List<Absencja> Absencje)`: Zwraca ilość dni roboczych w trakcie nieobecności.
- `private static string Dopasuj_TBN_Nazwa(RodzajAbsencji rodzaj)`: zwraca nazwę nieobecności dopasowaną do podanego kodu z cdn.TypNieobec

### Pracownik.cs
- `public int Get_PraId(SqlConnection connection, SqlTransaction transaction)`: Pobiera Id pracownika z bazy. Może zrobić throw errora np. jeśli nie będzie żadnego znalezionego takiego pracownika w bazie.

### Relacja.cs
- `public static int Get_Relacja_Id(string Numer_Relacji, SqlConnection connection, SqlTransaction transaction)`: Zwraca Id Relacji lub throwuje błąd żenie znaleziono relacji o takim numerze.
- 'public void Insert_Relacja_Do_Optimy(Error_Logger Internal_Error_Logger, SqlConnection connection, SqlTransaction transaction)' Jeśli takier relacji nie ma w bazie to taka się wpierdoli do bazy.

### Reader_*.cs
- Wczytują dane z plików Excel i dodają dane do bazy, operacje w tych plikach wyglądają w miare podobnie ale i tak murzyńsko bo te Arkusze od WARSu są zjebane. Nie chce mi się tego opisywać.
  
  1. Każdy plik ma własny error logger `Internal_Error_Logger` który inicjuje się po wywołaniu funkcji `Process_Zakladka()` i z niego należy kożystać w tym pliku (w razie potrzeby na wielowątkowe przetwarzanie).
  2. Najpierw znajdywane są wszystkie punkty odniesienia do kart pracy/grafików itp. za pomocą metody `Find_Starting_Points()` a następnie dla każego z tych punktów wykonywane są operacjeszykania i wczytywania danych odnoszące się do owych punktów startowch.
  3. Po zebraniu danych metody `Dodaj_*_Do_Optimy()` robią inserty itp.  

## Autorzy
- **ITEGER** - 27.02.2025
