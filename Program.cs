using System;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;

public class Program
{
    public static void Main(string[] args)
    {
        // 檢查管理員權限，若無則嘗試提權
        if (!IsAdministrator())
        {
            Logger.Log("未偵測到系統管理員權限，正在嘗試以系統管理員身分重新啟動。");
            RestartAsAdmin(args);
            return;
        }

        // 主排程邏輯
        try
        {
            Logger.Log("程式已使用系統管理員權限啟動。");
            RunScheduler();
            Logger.Log("排程器處理完畢。");
        }
        catch (Exception ex)
        {
            Logger.LogError("主排程器發生未預期的錯誤。", ex);
            // 發生錯誤時，暫停執行以便查看日誌
            Console.WriteLine("發生錯誤，按任意鍵結束...");
            Console.ReadKey();
        }
    }

    private static void RunScheduler()
    {
        // 1. 載入設定
        AppConfig config = ConfigManager.LoadConfig();
        Logger.Log($"設定已載入。執行間隔天數: {config.IntervalDays}, 上次執行日期: {config.LastRunDate:yyyy-MM-dd}");

        // 2. 驗證設定
        if (string.IsNullOrEmpty(config.VbsPath) || config.VbsPath.EndsWith("Path\\To\\Your\\Script.vbs"))
        {
            Logger.LogError("VBS 腳本路徑尚未設定。請編輯 run_config.ini 檔案。");
            return;
        }

        // 3. 檢查執行條件
        DateTime today = DateTime.Today;
        if ((today - config.LastRunDate).TotalDays < config.IntervalDays)
        {
            Logger.Log($"尚未到達 {config.IntervalDays} 天的執行間隔，程式即將關閉。");
            return;
        }

        // 4. 執行腳本
        Logger.Log($"已到達執行間隔，正在嘗試執行 VBS 腳本: {config.VbsPath}");
        try
        {
            ExecuteVbs(config.VbsPath);
            Logger.Log("VBS 腳本執行成功。");

            // 5. 更新設定檔
            config.LastRunDate = today;
            ConfigManager.SaveConfig(config);
            Logger.Log($"成功更新上次執行日期為 {today:yyyy-MM-dd}。");
        }
        catch (Exception ex)
        {
            Logger.LogError($"執行 VBS 腳本失敗。", ex);
        }
    }

    private static void ExecuteVbs(string scriptPath)
    {
        if (!File.Exists(scriptPath))
        {
            throw new FileNotFoundException("找不到 VBS 腳本檔案。", scriptPath);
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = "cscript.exe",
            Arguments = $"//Nologo \"{scriptPath}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using (var process = Process.Start(startInfo))
        {
            process.WaitForExit();
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            if (process.ExitCode != 0)
            {
                throw new Exception($"腳本執行失敗，結束代碼: {process.ExitCode}。\n輸出: {output}\n錯誤: {error}");
            }
            if (!string.IsNullOrWhiteSpace(output))
            {
                Logger.Log($"腳本輸出:\n{output}");
            }
        }
    }

    private static bool IsAdministrator()
    {
        using (var identity = WindowsIdentity.GetCurrent())
        {
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }

    private static void RestartAsAdmin(string[] args)
    {
        var exeName = Process.GetCurrentProcess().MainModule?.FileName;
        if (string.IsNullOrEmpty(exeName))
        {
            Logger.LogError("無法取得目前執行檔的路徑，無法重新啟動。");
            return;
        }

        var startInfo = new ProcessStartInfo(exeName)
        {
            Verb = "runas",
            Arguments = string.Join(" ", args),
            UseShellExecute = true
        };

        try
        {
            Process.Start(startInfo);
        }
        catch (System.ComponentModel.Win32Exception)
        {
            Logger.Log("使用者已取消 UAC 提示。程式需要系統管理員權限才能執行。");
        }
    }
}