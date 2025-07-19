namespace DingExcelTool.Utils
{
    using System;
    using Avalonia.Threading;
    using DingExcelTool;
    using Data;
    
    internal static class LogMessageHandler
    {
        public static MainWindow Window { get; private set; }

        public static void Init(MainWindow window) => Window = window;

        public static void AddInfo(string info)
        {
            Dispatcher.UIThread.Post(() =>
            {
                Window.AddLogItem(info, LogType.Info);
                Window.MoveLogListScroll();
            });
        }

        public static void AddWarn(string warn)
        {
            Dispatcher.UIThread.Post(() =>
            {
                Window.AddLogItem(warn, LogType.Warn);
                Window.MoveLogListScroll();
            });
        }

        public static void AddError(string error)
        {
            Dispatcher.UIThread.Post(() =>
            {
                Window.AddLogItem(error, LogType.Error);
                Window.MoveLogListScroll();
            });
        }

        public static void LogException(Exception ex)
        {
            AddError($"{ex.Message}\n{ex.StackTrace}");

            if (ex.InnerException != null)
            {
                AddError("Inner Exception:");
                LogException(ex.InnerException);
            }
        }
    }
}
