namespace WordMcpServer.Logging
{
    /// <summary>
    /// MCP 服务器内部日志接口。
    /// Word 是 GUI 进程，控制台不可见，因此默认实现写入文件；
    /// 后续如需接入 Microsoft.Extensions.Logging 可在此增加适配器。
    /// </summary>
    public interface IMcpLogger
    {
        void Debug(string message);

        void Info(string message);

        void Warn(string message);

        void Error(string message, System.Exception exception = null);
    }
}
