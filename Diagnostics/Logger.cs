using System.Collections.Concurrent;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowPilot.Diagnostics;

/// <summary>
/// 日志级别
/// </summary>
public enum LogLevel
{
    Trace   = 0,
    Debug   = 1,
    Info    = 2,
    Warning = 3,
    Error   = 4,
}

/// <summary>
/// WindowPilot 日志系统
/// ─────────────────────────────────────────────────────
/// · 彩色控制台输出（Console.Write）
/// · 每次运行生成独立日志文件，存放于 Logs/ 目录
/// · 线程安全：后台写文件线程，不阻塞调用方
/// · 调用方通过 [CallerMemberName] / [CallerFilePath] 自动附带来源信息
/// </summary>
public static class Logger
{
    // ── Win32（WinExe 无控制台，需手动申请） ─────────────────
    [DllImport("kernel32.dll")]
    private static extern bool AllocConsole();

    // ── 配置 ──────────────────────────────────────────────
    private static LogLevel _minConsoleLevel = LogLevel.Debug;
    private static LogLevel _minFileLevel    = LogLevel.Trace;

    /// <summary>最低控制台输出级别（默认 Debug）</summary>
    public static LogLevel MinConsoleLevel
    {
        get => _minConsoleLevel;
        set => _minConsoleLevel = value;
    }

    /// <summary>最低文件写入级别（默认 Trace）</summary>
    public static LogLevel MinFileLevel
    {
        get => _minFileLevel;
        set => _minFileLevel = value;
    }

    // ── 运行时状态 ────────────────────────────────────────
    private static string?              _logFilePath;
    private static StreamWriter?        _fileWriter;
    private static readonly object      _fileLock      = new();
    private static readonly BlockingCollection<string> _writeQueue = new(4096);
    private static Thread?              _writeThread;
    private static bool                 _initialized;
    private static readonly DateTime    _sessionStart  = DateTime.Now;
    private static long                 _entryCount;

    // ── 初始化 ────────────────────────────────────────────

    /// <summary>
    /// 初始化日志系统。应在程序启动时（App.OnStartup）调用一次。
    /// </summary>
    /// <param name="logDirectory">日志目录，默认为可执行文件旁的 Logs/</param>
    /// <param name="minConsoleLevel">控制台最低级别</param>
    /// <param name="minFileLevel">文件最低级别</param>
    public static void Initialize(
        string?  logDirectory    = null,
        LogLevel minConsoleLevel = LogLevel.Debug,
        LogLevel minFileLevel    = LogLevel.Trace)
    {
        if (_initialized) return;
        _initialized = true;

        // WPF (WinExe) 启动时没有控制台，必须先 AllocConsole() 申请控制台窗口，
        // 再设置编码；否则 Console.OutputEncoding = UTF8 会抛 IOException。
        AllocConsole();
        Console.OutputEncoding = Encoding.UTF8;
        Console.InputEncoding  = Encoding.UTF8;

        _minConsoleLevel = minConsoleLevel;
        _minFileLevel    = minFileLevel;

        // 确定日志目录
        logDirectory ??= Path.Combine(
            AppContext.BaseDirectory, "Logs");

        try
        {
            Directory.CreateDirectory(logDirectory);

            string timestamp  = _sessionStart.ToString("yyyy-MM-dd_HH-mm-ss");
            _logFilePath      = Path.Combine(logDirectory, $"WindowPilot_{timestamp}.log");

            _fileWriter = new StreamWriter(_logFilePath, append: false, Encoding.UTF8)
            {
                AutoFlush = false
            };

            // 启动后台写文件线程
            _writeThread = new Thread(FileWriteLoop)
            {
                IsBackground = true,
                Name         = "LogFileWriter",
                Priority     = ThreadPriority.BelowNormal
            };
            _writeThread.Start();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine($"[Logger] 无法初始化日志文件: {ex.Message}");
            Console.ResetColor();
        }

        // 写入会话头
        WriteSessionHeader();
    }

    /// <summary>
    /// 关闭日志系统，刷新并关闭文件。应在程序退出时调用。
    /// </summary>
    public static void Shutdown()
    {
        WriteSessionFooter();

        _writeQueue.CompleteAdding();
        _writeThread?.Join(TimeSpan.FromSeconds(3));

        lock (_fileLock)
        {
            _fileWriter?.Flush();
            _fileWriter?.Dispose();
            _fileWriter = null;
        }
    }

    // ── 公开日志方法 ──────────────────────────────────────

    public static void Trace(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Trace, message, category, member, filePath, line);

    public static void Debug(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Debug, message, category, member, filePath, line);

    public static void Info(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Info, message, category, member, filePath, line);

    public static void Warning(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Warning, message, category, member, filePath, line);

    public static void Error(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Error, message, category, member, filePath, line);

    public static void Error(string message, Exception ex,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Error,
            $"{message} | Exception: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}",
            category, member, filePath, line);

    /// <summary>
    /// 写入水平分隔线，用于在日志中划分阶段
    /// </summary>
    public static void Separator(string? label = null)
    {
        string line = label is null
            ? new string('─', 80)
            : $"── {label} {new string('─', Math.Max(0, 76 - label.Length))}";

        if (LogLevel.Info >= _minConsoleLevel)
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine(line);
            Console.ResetColor();
        }
        EnqueueFile(line);
    }

    // ── 核心写入逻辑 ──────────────────────────────────────

    private static void Write(
        LogLevel level,
        string   message,
        string   category,
        string   member,
        string   filePath,
        int      line)
    {
        long id        = Interlocked.Increment(ref _entryCount);
        var  now       = DateTime.Now;
        string elapsed = (now - _sessionStart).ToString(@"hh\:mm\:ss\.fff");
        string srcFile = Path.GetFileNameWithoutExtension(filePath);
        string tag     = string.IsNullOrWhiteSpace(category) ? srcFile : category;

        // ── 控制台 ──
        if (level >= _minConsoleLevel)
        {
            lock (Console.Out) // 防止多线程混色
            {
                // 时间戳（灰色）
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.Write($"[{elapsed}] ");

                // 级别（彩色）
                Console.ForegroundColor = LevelColor(level);
                Console.Write($"[{LevelTag(level)}] ");

                // 分类（青色）
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write($"[{tag,-20}] ");

                // 成员名（灰色）
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.Write($"{member}:{line,-4} ");

                // 消息（正常色）
                Console.ForegroundColor = MessageColor(level);
                Console.WriteLine(message);

                Console.ResetColor();
            }
        }

        // ── 文件（交由后台线程写） ──
        if (level >= _minFileLevel)
        {
            string fileEntry =
                $"{now:yyyy-MM-dd HH:mm:ss.fff} | {LevelTag(level),-7} | {tag,-20} | {srcFile}.{member}:{line,-4} | {message}";
            EnqueueFile(fileEntry);
        }
    }

    // ── 后台文件写入 ──────────────────────────────────────

    private static void EnqueueFile(string line)
    {
        if (_writeQueue.IsAddingCompleted) return;
        try { _writeQueue.TryAdd(line, millisecondsTimeout: 0); }
        catch { /* 队列已满时丢弃 */ }
    }

    private static void FileWriteLoop()
    {
        try
        {
            foreach (string line in _writeQueue.GetConsumingEnumerable())
            {
                lock (_fileLock)
                {
                    _fileWriter?.WriteLine(line);
                    // 每 50 行刷一次，平衡性能与及时性
                    if (_entryCount % 50 == 0)
                        _fileWriter?.Flush();
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[Logger.FileWriteLoop] 异常: {ex.Message}");
        }
        finally
        {
            lock (_fileLock) { _fileWriter?.Flush(); }
        }
    }

    // ── 会话头尾 ──────────────────────────────────────────

    private static void WriteSessionHeader()
    {
        string header =
            $"""
            ╔══════════════════════════════════════════════════════════════════════════════╗
            ║                          WindowPilot  Debug Session                         ║
            ╠══════════════════════════════════════════════════════════════════════════════╣
            ║  Started   : {_sessionStart:yyyy-MM-dd HH:mm:ss}                                       ║
            ║  Log File  : {Path.GetFileName(_logFilePath ?? "N/A"),-62}║
            ║  Host OS   : {Environment.OSVersion,-62}║
            ║  CLR       : {Environment.Version,-62}║
            ╚══════════════════════════════════════════════════════════════════════════════╝
            """;

        if (LogLevel.Info >= _minConsoleLevel)
        {
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine(header);
            Console.ResetColor();
        }
        EnqueueFile(header);
    }

    private static void WriteSessionFooter()
    {
        var duration = DateTime.Now - _sessionStart;
        string footer =
            $"""

            ╔══════════════════════════════════════════════════════════════════════════════╗
            ║                        WindowPilot  Session  End                            ║
            ╠══════════════════════════════════════════════════════════════════════════════╣
            ║  Ended     : {DateTime.Now:yyyy-MM-dd HH:mm:ss}                                       ║
            ║  Duration  : {duration:hh\:mm\:ss\.fff}                                                ║
            ║  Log Lines : {_entryCount,-62}║
            ╚══════════════════════════════════════════════════════════════════════════════╝
            """;

        if (LogLevel.Info >= _minConsoleLevel)
        {
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine(footer);
            Console.ResetColor();
        }
        EnqueueFile(footer);
    }

    // ── 辅助 ──────────────────────────────────────────────

    private static string LevelTag(LogLevel level) => level switch
    {
        LogLevel.Trace   => "TRACE",
        LogLevel.Debug   => "DEBUG",
        LogLevel.Info    => "INFO ",
        LogLevel.Warning => "WARN ",
        LogLevel.Error   => "ERROR",
        _                => "?????"
    };

    private static ConsoleColor LevelColor(LogLevel level) => level switch
    {
        LogLevel.Trace   => ConsoleColor.DarkGray,
        LogLevel.Debug   => ConsoleColor.Gray,
        LogLevel.Info    => ConsoleColor.Green,
        LogLevel.Warning => ConsoleColor.Yellow,
        LogLevel.Error   => ConsoleColor.Red,
        _                => ConsoleColor.White
    };

    private static ConsoleColor MessageColor(LogLevel level) => level switch
    {
        LogLevel.Trace   => ConsoleColor.DarkGray,
        LogLevel.Warning => ConsoleColor.Yellow,
        LogLevel.Error   => ConsoleColor.Red,
        _                => ConsoleColor.White
    };

    /// <summary>当前日志文件完整路径（可用于状态栏提示）</summary>
    public static string? CurrentLogFilePath => _logFilePath;
}