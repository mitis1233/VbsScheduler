using System;
using System.IO;

public class AppConfig
{
    public int IntervalDays { get; set; } = 30;
    public DateTime LastRunDate { get; set; } = DateTime.MinValue;
    public string VbsPath { get; set; } = @"C:\請在此處填寫您的腳本路徑\script.vbs";
}

public static class ConfigManager
{
    private const string ConfigFilePath = "run_config.ini";

    public static AppConfig LoadConfig()
    {
        var config = new AppConfig();

        if (!File.Exists(ConfigFilePath))
        {
            // 檔案不存在，直接儲存預設設定並回傳
            SaveConfig(config);
            return config;
        }

        var lines = File.ReadAllLines(ConfigFilePath);
        bool intervalDaysFound = false;
        bool vbsPathFound = false;

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith('#'))
                continue;

            var parts = line.Split('=', 2);
            if (parts.Length != 2)
                continue;

            var key = parts[0].Trim();
            var value = parts[1].Trim();

            switch (key.ToLower())
            {
                case "intervaldays":
                    if (int.TryParse(value, out int days) && days > 0)
                    {
                        config.IntervalDays = days;
                    }
                    intervalDaysFound = true;
                    break;
                case "lastrundate":
                    if (DateTime.TryParse(value, out DateTime date))
                    {
                        config.LastRunDate = date;
                    }
                    break;
                case "vbspath":
                    config.VbsPath = value;
                    vbsPathFound = true;
                    break;
            }
        }

        // 如果現有設定檔缺少項目，則補全並重新儲存
        if (!intervalDaysFound || !vbsPathFound)
        {
            SaveConfig(config);
        }

        return config;
    }

    public static void SaveConfig(AppConfig config)
    {
        string content = $"IntervalDays={config.IntervalDays}\n" +
                         $"LastRunDate={config.LastRunDate:o}\n" +
                         $"VbsPath={config.VbsPath}\n";
        File.WriteAllText(ConfigFilePath, content);
    }
}
