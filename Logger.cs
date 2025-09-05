using System;
using System.IO;

public static class Logger
{
    private const string LogFilePath = "VbsScheduler.log";

    public static void Log(string message)
    {
        try
        {
            string logMessage = $"[{DateTime.Now:yyyy-MM-dd}] {message}";
            
            // 寫入主控台
            Console.WriteLine(logMessage);
            
            // 附加到日誌檔案 (已停用)
            // File.AppendAllText(LogFilePath, logMessage + Environment.NewLine);
        }
        catch (Exception ex)
        {
            // 如果連日誌都寫入失敗，至少在主控台顯示錯誤
            Console.WriteLine($"致命錯誤：無法寫入日誌檔案。錯誤訊息: {ex.Message}");
        }
    }

    public static void LogError(string message, Exception? ex = null)
    {
        string errorMessage = $"錯誤: {message}";
        if (ex != null)
        {
            errorMessage += $"\n詳細資訊: {ex.ToString()}";
        }
        Log(errorMessage);
    }
}
