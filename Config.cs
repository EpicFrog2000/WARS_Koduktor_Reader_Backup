using System.Text.Json;

namespace Excel_Data_Importer_WARS
{
    internal class Config
    {
        public List<string> Files_Folders { get; set; } = [Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files")];
        public string Nazwa_Serwera { get; set; } = string.Empty;
        public string Nazwa_Bazy { get; set; } = string.Empty;
        public bool Clear_Logs_On_Program_Restart { get; set; } = false;
        public bool Clear_Bad_Files_On_Restart { get; set; } = false;
        public bool Clear_Processed_Files_On_Restart { get; set; } = false;
        public bool Move_Files_To_Processed_Folder { get; set; } = false;
        public bool Clear_Good_Files_On_Restart { get; set; } = false;
        public bool Tryb_Zapetlony { get; set; } = false;

        private readonly JsonSerializerOptions JsonSerializerOptions = new() { WriteIndented = true };
        public bool GetConfigFromFile()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json");
            bool existed = Check_File(filePath);
            if (!File.Exists(filePath))
            {
                var defaultConfig = new
                {
                    Files_Folders,
                    Nazwa_Serwera,
                    Nazwa_Bazy,
                    Clear_Processed_Files_On_Restart,
                    Clear_Bad_Files_On_Restart,
                    Clear_Logs_On_Program_Restart,
                    Move_Files_To_Processed_Folder,
                    Clear_Good_Files_On_Restart,
                    Tryb_Zapetlony
                };

                string defaultJson = JsonSerializer.Serialize(defaultConfig, JsonSerializerOptions);
                File.WriteAllText(filePath, defaultJson);
            }

            string json = File.ReadAllText(filePath);
            Config new_config = JsonSerializer.Deserialize<Config>(json)!;
            Files_Folders = new_config.Files_Folders;
            Nazwa_Serwera = new_config.Nazwa_Serwera ?? "";
            Nazwa_Bazy = new_config.Nazwa_Bazy ?? "";
            DbManager.Build_Connection_String(Nazwa_Serwera, Nazwa_Bazy);
            Clear_Logs_On_Program_Restart = new_config.Clear_Logs_On_Program_Restart;
            Clear_Bad_Files_On_Restart = new_config.Clear_Bad_Files_On_Restart;
            Clear_Processed_Files_On_Restart = new_config.Clear_Processed_Files_On_Restart;
            Move_Files_To_Processed_Folder = new_config.Move_Files_To_Processed_Folder;
            Clear_Good_Files_On_Restart = new_config.Clear_Good_Files_On_Restart;
            Tryb_Zapetlony = new_config.Tryb_Zapetlony;
            return existed;
        }
        public bool GetConfigFromFile(string Config_File_Path)
        {
            bool existed = Check_File(Config_File_Path);
            if (!File.Exists(Config_File_Path))
            {
                File.Create(Config_File_Path).Dispose();
                var defaultConfig = new
                {
                    Nazwa_Serwera,
                    Nazwa_Bazy,
                    Clear_Processed_Files_On_Restart,
                    Clear_Bad_Files_On_Restart,
                    Clear_Logs_On_Program_Restart,
                    Move_Files_To_Processed_Folder,
                    Clear_Good_Files_On_Restart,
                    Tryb_Zapetlony
                };
                File.WriteAllText(Config_File_Path, JsonSerializer.Serialize(defaultConfig, JsonSerializerOptions));
            }
            string json = File.ReadAllText(Config_File_Path);
            Config? config = JsonSerializer.Deserialize<Config>(json);
            if (config != null)
            {
                Files_Folders = config.Files_Folders;
                Nazwa_Serwera = config.Nazwa_Serwera;
                Nazwa_Bazy = config.Nazwa_Bazy;
                Clear_Logs_On_Program_Restart = config.Clear_Logs_On_Program_Restart;
                Clear_Bad_Files_On_Restart = config.Clear_Bad_Files_On_Restart;
                Clear_Processed_Files_On_Restart = config.Clear_Processed_Files_On_Restart;
                Move_Files_To_Processed_Folder = config.Move_Files_To_Processed_Folder;
                Clear_Good_Files_On_Restart = config.Clear_Good_Files_On_Restart;
                Tryb_Zapetlony = config.Tryb_Zapetlony;
            }
            DbManager.Build_Connection_String(Nazwa_Serwera, Nazwa_Bazy);
            return existed;
        }
        public bool Check_File()
        {
            string Config_File_Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json");
            if (!File.Exists(Config_File_Path))
            {
                File.Create(Config_File_Path).Dispose();
                var defaultConfig = new
                {
                    Files_Folders = new[] { Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files") },
                    Nazwa_Serwera,
                    Nazwa_Bazy,
                    Clear_Logs_On_Program_Restart,
                    Clear_Processed_Files_On_Restart,
                    Clear_Bad_Files_On_Restart,
                    Move_Files_To_Processed_Folder,
                    Clear_Good_Files_On_Restart,
                    Tryb_Zapetlony
                };
                File.WriteAllText(Config_File_Path, JsonSerializer.Serialize(defaultConfig, JsonSerializerOptions));
                return false;
            }
            else
            {
                string configFileContent = File.ReadAllText(Config_File_Path);
                Dictionary<string, object> currentConfig = JsonSerializer.Deserialize<Dictionary<string, object>>(configFileContent) ?? [];
                Dictionary<string, object> defaultConfig = new()
                {
                    { "Files_Folders", new[] { Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files") } },
                    { "Nazwa_Serwera", Nazwa_Serwera },
                    { "Nazwa_Bazy", Nazwa_Bazy },
                    { "Clear_Logs_On_Program_Restart", Clear_Logs_On_Program_Restart },
                    { "Clear_Processed_Files_On_Restart", Clear_Processed_Files_On_Restart },
                    { "Clear_Bad_Files_On_Restart", Clear_Bad_Files_On_Restart },
                    { "Move_Files_To_Processed_Folder", Move_Files_To_Processed_Folder },
                    { "Clear_Good_Files_On_Restart", Clear_Good_Files_On_Restart },
                    { "Tryb_Zapetlony", Tryb_Zapetlony }
                };

                foreach (var key in defaultConfig.Keys)
                {
                    if (!currentConfig.ContainsKey(key))
                    {
                        currentConfig[key] = defaultConfig[key];
                    }
                }
            }
            return true;
        }
        public bool Check_File(string Config_File_Path)
        {
            if (!File.Exists(Config_File_Path))
            {
                File.Create(Config_File_Path).Dispose();
                var defaultConfig = new
                {
                    Files_Folders = new[] { Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files") },
                    Nazwa_Serwera,
                    Nazwa_Bazy,
                    Clear_Logs_On_Program_Restart,
                    Clear_Processed_Files_On_Restart,
                    Clear_Bad_Files_On_Restart,
                    Move_Files_To_Processed_Folder,
                    Clear_Good_Files_On_Restart,
                    Tryb_Zapetlony
                };
                File.WriteAllText(Config_File_Path, JsonSerializer.Serialize(defaultConfig, JsonSerializerOptions));
                return false;
            }
            else
            {
                string configFileContent = File.ReadAllText(Config_File_Path);
                Dictionary<string, object> currentConfig = JsonSerializer.Deserialize<Dictionary<string, object>>(configFileContent) ?? [];
                Dictionary<string, object> defaultConfig = new()
                {
                    { "Files_Folders", new[] { Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files") } },
                    { "Nazwa_Serwera", Nazwa_Serwera },
                    { "Nazwa_Bazy", Nazwa_Bazy },
                    { "Clear_Logs_On_Program_Restart", Clear_Logs_On_Program_Restart },
                    { "Clear_Processed_Files_On_Restart", Clear_Processed_Files_On_Restart },
                    { "Clear_Bad_Files_On_Restart", Clear_Bad_Files_On_Restart },
                    { "Move_Files_To_Processed_Folder", Move_Files_To_Processed_Folder },
                    { "Clear_Good_Files_On_Restart", Clear_Good_Files_On_Restart },
                    { "Tryb_Zapetlony", Tryb_Zapetlony }
                };

                foreach (var key in defaultConfig.Keys)
                {
                    if (!currentConfig.ContainsKey(key))
                    {
                        currentConfig[key] = defaultConfig[key];
                    }
                }
            }
            return true;
        }
    }
}
