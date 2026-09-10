using System;
using System.IO;

namespace WordMcpServer.Logging
{
    /// <summary>
    /// 文件日志实现：默认写入 %APPDATA%\WordMcpServer\Logs\wordmcp.log，
    /// 可通过 mcp.config.json 的 logDirectory 覆盖。
    /// 单文件超过 5MB 时滚动为 .bak（进程内互斥，避免多线程并发写损坏）。
    /// </summary>
    public sealed class FileMcpLogger : IMcpLogger
    {
        private const long MaxLogFileBytes = 5L * 1024 * 1024;

        private readonly object _sync = new object();
        private readonly string _filePath;

        public FileMcpLogger() : this(null)
        {
        }

        public FileMcpLogger(string logDirectory)
        {
            if (string.IsNullOrWhiteSpace(logDirectory))
            {
                logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "WordMcpServer", "Logs");
            }

            Directory.CreateDirectory(logDirectory);
            _filePath = Path.Combine(logDirectory, "wordmcp.log");
        }

        public string FilePath => _filePath;

        public void Debug(string message) => Write("DEBUG", message, null);

        public void Info(string message) => Write("INFO", message, null);

        public void Warn(string message) => Write("WARN", message, null);

        public void Error(string message, Exception exception = null) => Write("ERROR", message, exception);

        private void Write(string level, string message, Exception exception)
        {
            lock (_sync)
            {
                try
                {
                    RotateIfNeeded();
                    var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}";
                    if (exception != null)
                    {
                        line += Environment.NewLine + exception;
                    }

                    File.AppendAllText(_filePath, line + Environment.NewLine);
                }
                catch
                {
                    // 日志失败绝不能影响插件主流程
                }
            }
        }

        private void RotateIfNeeded()
        {
            var fi = new FileInfo(_filePath);
            if (fi.Exists && fi.Length > MaxLogFileBytes)
            {
                var backup = _filePath + ".bak";
                try
                {
                    if (File.Exists(backup))
                    {
                        File.Delete(backup);
                    }

                    File.Move(_filePath, backup);
                }
                catch
                {
                    // 滚动失败时继续写原文件
                }
            }
        }
    }
}
