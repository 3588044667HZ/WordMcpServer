using System;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using WordMcpServer.Logging;

namespace WordMcpServer.Mcp.Threading
{
    /// <summary>
    /// Word 主线程封送器。
    ///
    /// 线程模型背景：
    /// 1. Word 是 STA 原生应用，其对象模型（PIA 调用）必须在 Word 主 UI 线程
    ///    上执行，否则会抛出"调用被拒绝"（RPC_E_WRONG_THREAD）等 COM 错误。
    /// 2. Kestrel 收到 MCP 请求后在线程池线程上处理，这些线程不属于 UI 线程。
    /// 3. 因此所有 Word 对象模型操作必须先封送到 UI 线程再执行。
    ///
    /// 实现方式：
    /// - VSTO 环境下 SynchronizationContext.Current 通常为 null（Office 未安装 WinForms
    ///   消息泵上下文），这里在启动阶段（此时正处于 Word 主线程）手工安装一个
    ///   WindowsFormsSynchronizationContext 并创建隐藏控件锚点，确保有消息循环可投递。
    /// - Send 调用在 UI 线程同步执行（阻塞调用方线程，符合请求/响应语义）；
    ///   若调用方本身已在 UI 线程则直接执行，避免死锁。
    /// </summary>
    public sealed class OfficeThreadMarshaller : IDisposable
    {
        private readonly SynchronizationContext _uiContext;
        private readonly Control _anchor;
        private readonly IMcpLogger _logger;

        private OfficeThreadMarshaller(SynchronizationContext uiContext, Control anchor, IMcpLogger logger)
        {
            _uiContext = uiContext;
            _anchor = anchor;
            _logger = logger;
        }

        /// <summary>
        /// 必须在 Word 主线程（ThisAddIn_Startup 中）调用：
        /// 安装同步上下文并固定消息循环锚点。
        /// </summary>
        public static OfficeThreadMarshaller Install(IMcpLogger logger)
        {
            // VSTO 启动时当前线程即 Word 主 UI 线程
            if (SynchronizationContext.Current == null)
            {
                var ctx = new WindowsFormsSynchronizationContext();
                SynchronizationContext.SetSynchronizationContext(ctx);
            }

            // 创建隐藏控件锚点：WinForms 同步上下文通过控件句柄投递回调，
            // 保持引用可防止句柄/上下文被提前回收；Word 自身的消息泵会处理投递过来的消息。
            var anchor = new Control { Visible = false };
            anchor.CreateControl();

            return new OfficeThreadMarshaller(SynchronizationContext.Current, anchor, logger);
        }

        /// <summary>当前线程是否就是 Word 主线程。</summary>
        public bool IsOnUiThread => SynchronizationContext.Current == _uiContext;

        public void Invoke(Action action) => Invoke<object>(() => { action(); return null; });

        /// <summary>
        /// 把委托封送到 Word 主线程同步执行，返回其返回值。
        /// 调用方线程（Kestrel 线程池线程）会被阻塞，直到 UI 线程执行完成；
        /// 异常通过 ExceptionDispatchInfo 原样重新抛出（保留原始堆栈与类型，如 COMException）。
        /// </summary>
        public TResult Invoke<TResult>(Func<TResult> func)
        {
            if (IsOnUiThread)
            {
                return func();
            }

            TResult result = default;
            ExceptionDispatchInfo captured = null;
            try
            {
                // Send 阻塞调用方线程，直到 UI 线程上的回调完成
                _uiContext.Send(_ =>
                {
                    try
                    {
                        result = func();
                    }
                    catch (Exception ex)
                    {
                        captured = ExceptionDispatchInfo.Capture(ex);
                    }
                }, null);
            }
            catch (Exception ex)
            {
                // UI 线程已退出（如 Word 正在关闭）时 Send 会抛异常
                throw new InvalidOperationException("无法封送到 Word 主线程：UI 线程可能已退出。", ex);
            }

            if (captured != null)
            {
                captured.Throw();
            }

            return result;
        }

        public void Dispose()
        {
            try
            {
                _anchor?.Dispose();
            }
            catch
            {
                // 关闭阶段的清理失败不应影响卸载流程
            }
        }
    }
}
