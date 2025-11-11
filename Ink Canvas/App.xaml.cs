using Ink_Canvas.Helpers;
using iNKORE.UI.WPF.Modern.Controls;
using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace Ink_Canvas
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        System.Threading.Mutex mutex;

        public static string[] StartArgs = null;
        public static string RootPath = Environment.GetEnvironmentVariable("APPDATA") + "\\Ink Canvas\\";
        
        private static int crashCount = 0;
        private static DateTime lastCrashTime = DateTime.MinValue;
        private const int MaxCrashCount = 3;
        private const int CrashCountResetMinutes = 10;

        public App()
        {
            this.Startup += new StartupEventHandler(App_Startup);
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LogHelper.WriteLogToFile($"崩溃异常捕获: {e.Exception}", LogHelper.LogType.Error);
            HandleCrash(e.Exception, "主线程异常");
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            LogHelper.WriteLogToFile($"应用域异常捕获: {ex}", LogHelper.LogType.Error);
            HandleCrash(ex, "应用域异常");
        }

        private void HandleCrash(Exception exception, string crashType)
        {
            try
            {
                // 重置崩溃计数（如果距离上次崩溃超过指定时间）
                if (DateTime.Now.Subtract(lastCrashTime).TotalMinutes > CrashCountResetMinutes)
                {
                    crashCount = 0;
                }

                crashCount++;
                lastCrashTime = DateTime.Now;

                LogHelper.WriteLogToFile($"[{crashType}] 崩溃计数: {crashCount}/{MaxCrashCount}", LogHelper.LogType.Error);

                if (crashCount >= MaxCrashCount)
                {
                    // 达到最大崩溃次数，放弃重启
                    string message = $"应用已连续崩溃 {MaxCrashCount} 次，无法自动恢复。\n\n错误信息:\n{exception?.Message}\n\n请检查您的系统环境或联系开发者。";
                    MessageBox.Show(message, "Ink Canvas - 严重错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    LogHelper.WriteLogToFile("达到最大崩溃次数，程序退出", LogHelper.LogType.Error);
                    Environment.Exit(1);
                    return;
                }

                // 显示崩溃通知并准备重启
                string restartMessage = $"Ink Canvas 遇到了一个严重错误，将在 3 秒后自动重启。\n\n错误类型: {crashType}\n错误信息: {exception?.Message}";
                Ink_Canvas.MainWindow.ShowNewMessage(restartMessage, true);

                // 延迟后重启应用
                Task.Delay(3000).ContinueWith(_ =>
                {
                    RestartApplication();
                });
            }
            catch (Exception ex)
            {
                LogHelper.WriteLogToFile($"处理崩溃时出错: {ex}", LogHelper.LogType.Error);
                try
                {
                    RestartApplication();
                }
                catch { }
            }
        }

        private static void RestartApplication()
        {
            try
            {
                LogHelper.WriteLogToFile("开始重启应用...", LogHelper.LogType.Event);
                
                // 获取当前应用的可执行文件路径
                string exePath = Process.GetCurrentProcess().MainModule.FileName;
                
                // 启动新进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = string.Join(" ", StartArgs ?? Array.Empty<string>()),
                    UseShellExecute = true
                };
                
                Process.Start(psi);
                LogHelper.WriteLogToFile("新进程已启动，退出当前应用", LogHelper.LogType.Event);
                
                // 关闭当前应用
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLogToFile($"应用重启失败: {ex}", LogHelper.LogType.Error);
                Environment.Exit(1);
            }
        }

        void App_Startup(object sender, StartupEventArgs e)
        {
            /*if (!StoreHelper.IsStoreApp) */RootPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            LogHelper.NewLog(string.Format("Ink Canvas Starting (Version: {0})", Assembly.GetExecutingAssembly().GetName().Version.ToString()));

            bool ret;
            mutex = new System.Threading.Mutex(true, "Ink_Canvas_Artistry", out ret);

            if (!ret && !e.Args.Contains("-m")) //-m multiple
            {
                LogHelper.NewLog("Detected existing instance");
                MessageBox.Show("已有一个程序实例正在运行");
                LogHelper.NewLog("Ink Canvas automatically closed");
                Environment.Exit(0);
            }

            StartArgs = e.Args;
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            try
            {
                if (System.Windows.Forms.SystemInformation.MouseWheelScrollLines == -1)
                    e.Handled = false;
                else
                    try
                    {
                        ScrollViewerEx SenderScrollViewer = (ScrollViewerEx)sender;
                        SenderScrollViewer.ScrollToVerticalOffset(SenderScrollViewer.VerticalOffset - e.Delta * 10 * System.Windows.Forms.SystemInformation.MouseWheelScrollLines / (double)120);
                        e.Handled = true;
                    }
                    catch {  }
            }
            catch {  }
        }
    }
}
