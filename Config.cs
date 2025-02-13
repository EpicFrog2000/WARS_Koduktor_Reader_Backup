using System.Text.Json;

namespace Konduktor_Reader
{
    internal class Config
    {
        public List<string> Files_Folders { get; set; } = [Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files")];
        public string Optima_Conection_String { get; set; } = "Server=ITEGERNT;Database=CDN_Wars_prod_ITEGER_22012025;Encrypt=True;TrustServerCertificate=True;Integrated Security=True;";
        public bool Clear_Logs_On_Program_Restart { get; set; } = false;
        public bool Clear_Bad_Files_On_Restart { get; set; } = false;
        public bool Clear_Processed_Files_On_Restart { get; set; } = false;
        public bool Move_Files_To_Processed_Folder { get; set; } = false;
        public void GetConfigFromFile()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json");
            Check_File(filePath);
            if (!File.Exists(filePath))
            {
                File.Create(filePath).Dispose();
                var defaultConfig = new
                {
                    Optima_Conection_String,
                    Clear_Processed_Files_On_Restart,
                    Clear_Bad_Files_On_Restart,
                    Clear_Logs_On_Program_Restart,
                    Move_Files_To_Processed_Folder,
                };
                File.WriteAllText(filePath, JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
            }
            string json = File.ReadAllText(filePath);
            Config? config = JsonSerializer.Deserialize<Config>(json);
            if (config != null)
            {
                Files_Folders = config.Files_Folders;
                Optima_Conection_String = config.Optima_Conection_String;
                Clear_Logs_On_Program_Restart = config.Clear_Logs_On_Program_Restart;
                Clear_Bad_Files_On_Restart = config.Clear_Bad_Files_On_Restart;
                Clear_Processed_Files_On_Restart = config.Clear_Processed_Files_On_Restart;
                Move_Files_To_Processed_Folder = config.Move_Files_To_Processed_Folder;
            }
        }
        public void GetConfigFromFile(string Config_File_Path)
        {
            Check_File(Config_File_Path);
            if (!File.Exists(Config_File_Path))
            {
                File.Create(Config_File_Path).Dispose();
                var defaultConfig = new
                {
                    Optima_Conection_String,
                    Clear_Processed_Files_On_Restart,
                    Clear_Bad_Files_On_Restart,
                    Clear_Logs_On_Program_Restart,
                    Move_Files_To_Processed_Folder,
                };
                File.WriteAllText(Config_File_Path, JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
            }
            string json = File.ReadAllText(Config_File_Path);
            Config? config = JsonSerializer.Deserialize<Config>(json);
            if (config != null)
            {
                Files_Folders = config.Files_Folders;
                Optima_Conection_String = config.Optima_Conection_String;
                Clear_Logs_On_Program_Restart = config.Clear_Logs_On_Program_Restart;
                Clear_Bad_Files_On_Restart = config.Clear_Bad_Files_On_Restart;
                Clear_Processed_Files_On_Restart = config.Clear_Processed_Files_On_Restart;
                Move_Files_To_Processed_Folder = config.Move_Files_To_Processed_Folder;
            }
        }
        public void Check_File()
        {
            string Config_File_Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json");
            if (!File.Exists(Config_File_Path))
            {
                File.Create(Config_File_Path).Dispose();
                var defaultConfig = new
                {
                    Files_Folders = new[] { Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files") },
                    Optima_Conection_String,
                    Clear_Logs_On_Program_Restart,
                    Clear_Processed_Files_On_Restart,
                    Clear_Bad_Files_On_Restart,
                    Move_Files_To_Processed_Folder,
                };
                File.WriteAllText(Config_File_Path, JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
            }
        }
        public void Check_File(string Config_File_Path)
        {
            if (!File.Exists(Config_File_Path))
            {
                File.Create(Config_File_Path).Dispose();
                var defaultConfig = new
                {
                    Files_Folders = new[] { Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files") },
                    Optima_Conection_String,
                    Clear_Logs_On_Program_Restart,
                    Clear_Processed_Files_On_Restart,
                    Clear_Bad_Files_On_Restart,
                    Move_Files_To_Processed_Folder,
                };
                File.WriteAllText(Config_File_Path, JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
            }
        }
    }
}
