using System.Collections.Concurrent;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowPilot.Diagnostics;

/// <summary>
/// 日志级别枚举，数值越大表示越严重。
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
/// 线程安全的静态日志系统，支持彩色控制台输出和异步文件写入。
/// 需要在程序启动时调用 <see cref="Initialize"/> 并在退出前调用 <see cref="Shutdown"/>。
/// </summary>
public static class Logger
{
    // WPF 应用默认没有控制台窗口，必须通过此 API 手动申请
    [DllImport("kernel32.dll")]
    private static extern bool AllocConsole();

    // 控制台与文件的独立过滤级别
    private static LogLevel _minConsoleLevel = LogLevel.Debug;
    private static LogLevel _minFileLevel    = LogLevel.Trace;

    public static LogLevel MinConsoleLevel
    {
        get => _minConsoleLevel;
        set => _minConsoleLevel = value;
    }

    public static LogLevel MinFileLevel
    {
        get => _minFileLevel;
        set => _minFileLevel = value;
    }

    // 日志文件路径
    private static string?       _logFilePath;
    // 文件流写入器
    private static StreamWriter? _fileWriter;
    // 文件写入互斥锁，防止多线程同时写入
    private static readonly object _fileLock = new();
    // 异步写入队列，避免日志写入阻塞主线程
    private static readonly BlockingCollection<string> _writeQueue = new(4096);
    // 后台文件写入线程
    private static Thread?       _writeThread;
    // 防止重复初始化的标志
    private static bool          _initialized;
    // 会话开始时间，用于计算相对时间戳
    private static readonly DateTime _sessionStart = DateTime.Now;
    // 累计日志条目数，用于控制刷新频率
    private static long _entryCount;

    /// <summary>
    /// 初始化日志系统，在 <c>App.OnStartup</c> 中最先调用。
    /// </summary>
    /// <param name="logDirectory">日志文件存放目录，默认为可执行文件旁的 <c>Logs/</c> 子目录。</param>
    /// <param name="minConsoleLevel">控制台输出的最低级别，低于此级别的日志不显示。</param>
    /// <param name="minFileLevel">文件写入的最低级别，低于此级别的日志不写入文件。</param>
    public static void Initialize(
        string?  logDirectory    = null,
        LogLevel minConsoleLevel = LogLevel.Debug,
        LogLevel minFileLevel    = LogLevel.Trace)
    {
        // 防止重复初始化
        if (_initialized) return;
        _initialized = true;

        // WPF (WinExe) 启动时没有控制台，必须先申请控制台窗口，再设置编码
        // 若顺序颠倒，设置编码会抛 IOException
        AllocConsole();
        Console.OutputEncoding = Encoding.UTF8;
        Console.InputEncoding  = Encoding.UTF8;

        _minConsoleLevel = minConsoleLevel;
        _minFileLevel    = minFileLevel;

        // 未指定目录时默认使用可执行文件旁的 Logs 子目录
        logDirectory ??= Path.Combine(AppContext.BaseDirectory, "Logs");

        try
        {
            Directory.CreateDirectory(logDirectory);

            // 以启动时间戳为文件名，避免多次运行的日志互相覆盖
            string timestamp = _sessionStart.ToString("yyyy-MM-dd_HH-mm-ss");
            _logFilePath     = Path.Combine(logDirectory, $"WindowPilot_{timestamp}.log");

            // AutoFlush 关闭，由后台线程按批次手动 Flush 以提升性能
            _fileWriter = new StreamWriter(_logFilePath, append: false, Encoding.UTF8)
            {
                AutoFlush = false
            };

            // 后台写入线程优先级设为 BelowNormal，不与主线程争抢 CPU
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

        WriteSessionHeader();
    }

    // 关闭日志系统，刷新并关闭文件，在程序退出时调用
    public static void Shutdown()
    {
        WriteSessionFooter();

        // 标记队列完成，后台线程将处理完剩余条目后自然退出
        _writeQueue.CompleteAdding();
        // 等待后台线程写完，最多等待 3 秒
        _writeThread?.Join(TimeSpan.FromSeconds(3));

        lock (_fileLock)
        {
            _fileWriter?.Flush();
            _fileWriter?.Dispose();
            _fileWriter = null;
        }
    }

    // 公开日志方法，CallerMemberName 等编译器属性由运行时自动填充，无需手动传入

    /// <summary>
    /// 写入 Trace 级别日志，用于高频追踪信息。
    /// </summary>
    /// <param name="message">日志内容。</param>
    /// <param name="category">分类标签，通常使用调用模块的名称。</param>
    /// <param name="member">调用方法名，由编译器自动填充，无需手动传入。</param>
    /// <param name="filePath">调用文件路径，由编译器自动填充，无需手动传入。</param>
    /// <param name="line">调用行号，由编译器自动填充，无需手动传入。</param>
    public static void Trace(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Trace, message, category, member, filePath, line);

    /// <summary>
    /// 写入 Debug 级别日志，用于开发调试信息。
    /// </summary>
    /// <param name="message">日志内容。</param>
    /// <param name="category">分类标签。</param>
    /// <param name="member">调用方法名，由编译器自动填充。</param>
    /// <param name="filePath">调用文件路径，由编译器自动填充。</param>
    /// <param name="line">调用行号，由编译器自动填充。</param>
    public static void Debug(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Debug, message, category, member, filePath, line);

    /// <summary>
    /// 写入 Info 级别日志，用于关键业务流程节点。
    /// </summary>
    /// <param name="message">日志内容。</param>
    /// <param name="category">分类标签。</param>
    /// <param name="member">调用方法名，由编译器自动填充。</param>
    /// <param name="filePath">调用文件路径，由编译器自动填充。</param>
    /// <param name="line">调用行号，由编译器自动填充。</param>
    public static void Info(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Info, message, category, member, filePath, line);

    /// <summary>
    /// 写入 Warning 级别日志，用于非致命的异常情况。
    /// </summary>
    /// <param name="message">日志内容。</param>
    /// <param name="category">分类标签。</param>
    /// <param name="member">调用方法名，由编译器自动填充。</param>
    /// <param name="filePath">调用文件路径，由编译器自动填充。</param>
    /// <param name="line">调用行号，由编译器自动填充。</param>
    public static void Warning(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Warning, message, category, member, filePath, line);

    /// <summary>
    /// 写入 Error 级别日志，用于功能失败或不可恢复的错误。
    /// </summary>
    /// <param name="message">日志内容。</param>
    /// <param name="category">分类标签。</param>
    /// <param name="member">调用方法名，由编译器自动填充。</param>
    /// <param name="filePath">调用文件路径，由编译器自动填充。</param>
    /// <param name="line">调用行号，由编译器自动填充。</param>
    public static void Error(string message,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Error, message, category, member, filePath, line);

    /// <summary>
    /// 写入 Error 级别日志，附带异常详情和堆栈信息。
    /// </summary>
    /// <param name="message">描述发生了什么的上下文信息。</param>
    /// <param name="ex">捕获到的异常对象，类型名、消息和堆栈都会被记录。</param>
    /// <param name="category">分类标签。</param>
    /// <param name="member">调用方法名，由编译器自动填充。</param>
    /// <param name="filePath">调用文件路径，由编译器自动填充。</param>
    /// <param name="line">调用行号，由编译器自动填充。</param>
    public static void Error(string message, Exception ex,
        string category = "",
        [CallerMemberName] string member   = "",
        [CallerFilePath]   string filePath = "",
        [CallerLineNumber] int    line     = 0)
        => Write(LogLevel.Error,
            $"{message} | Exception: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}",
            category, member, filePath, line);

    /// <summary>
    /// 向日志中插入一条带可选标签的水平分隔线，用于划分不同阶段。
    /// </summary>
    /// <param name="label">嵌入在分隔线中的文字标签，为 null 时输出纯分隔线。</param>
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

    // 核心写入逻辑，同时输出到控制台和文件队列
    private static void Write(
        LogLevel level,
        string   message,
        string   category,
        string   member,
        string   filePath,
        int      line)
    {
        long id      = Interlocked.Increment(ref _entryCount);
        var  now     = DateTime.Now;
        // 计算相对于会话开始的经过时间，便于快速定位耗时
        string elapsed = (now - _sessionStart).ToString(@"hh\:mm\:ss\.fff");
        // 取文件名（不含扩展名）作为来源标识
        string srcFile = Path.GetFileNameWithoutExtension(filePath);
        // category 为空时回退到源文件名
        string tag = string.IsNullOrWhiteSpace(category) ? srcFile : category;

        if (level >= _minConsoleLevel)
        {
            // 加锁防止多线程并发输出时颜色混乱
            lock (Console.Out)
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.Write($"[{elapsed}] ");

                Console.ForegroundColor = LevelColor(level);
                Console.Write($"[{LevelTag(level)}] ");

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write($"[{tag,-20}] ");

                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.Write($"{member}:{line,-4} ");

                Console.ForegroundColor = MessageColor(level);
                Console.WriteLine(message);

                Console.ResetColor();
            }
        }

        if (level >= _minFileLevel)
        {
            string fileEntry = $"{now:yyyy-MM-dd HH:mm:ss.fff} | {LevelTag(level),-7} | {tag,-20} | {srcFile}.{member}:{line,-4} | {message}";
            // 放入队列由后台线程异步写入，避免阻塞调用方
            EnqueueFile(fileEntry);
        }
    }

    // 将日志行放入异步写入队列，队列满时直接丢弃而不阻塞
    private static void EnqueueFile(string line)
    {
        if (_writeQueue.IsAddingCompleted) return;
        try { _writeQueue.TryAdd(line, millisecondsTimeout: 0); }
        catch { /* 队列已满，丢弃该条日志 */ }
    }

    // 后台线程循环，从队列中取日志行写入文件
    private static void FileWriteLoop()
    {
        try
        {
            foreach (string line in _writeQueue.GetConsumingEnumerable())
            {
                lock (_fileLock)
                {
                    _fileWriter?.WriteLine(line);
                    // 每累积 50 条记录做一次 Flush，平衡性能与数据安全
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
            // 退出前确保所有缓冲内容都写入磁盘
            lock (_fileLock) { _fileWriter?.Flush(); }
        }
    }

    // 写入会话开始标头
    private static void WriteSessionHeader()
    {
        string header =
            $"""
            ---------------------------DEBUG模式启动---------------------------
            启动时间: {_sessionStart:yyyy-MM-dd HH:mm:ss}
            日志文件路径: {Path.GetFileName(_logFilePath ?? "N/A"),-62}
            Host OS: {Environment.OSVersion,-62}
            CLR: {Environment.Version,-62}
            ------------------------------------------------------------------
            """;

        if (LogLevel.Info >= _minConsoleLevel)
        {
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine(header);
            Console.ResetColor();
        }
        EnqueueFile(header);
    }

    // 写入会话结束标脚
    private static void WriteSessionFooter()
    {
        var duration = DateTime.Now - _sessionStart;
        string footer =
            $"""
            ------------------------------------------------------------------
            结束时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}
            DEBUG时长: {duration:hh\:mm\:ss\.fff}
            日志长度: {_entryCount}
            ---------------------------DEBUG模式结束----------------------------
            """;

        if (LogLevel.Info >= _minConsoleLevel)
        {
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine(footer);
            Console.ResetColor();
        }
        EnqueueFile(footer);
    }

    // 辅助方法

    // 返回级别对应的短标签字符串
    private static string LevelTag(LogLevel level) => level switch
    {
        LogLevel.Trace   => "TRACE",
        LogLevel.Debug   => "DEBUG",
        LogLevel.Info    => "INFO ",
        LogLevel.Warning => "WARN ",
        LogLevel.Error   => "ERROR",
        _                => "?????"
    };

    // 返回级别对应的控制台前景色
    private static ConsoleColor LevelColor(LogLevel level) => level switch
    {
        LogLevel.Trace   => ConsoleColor.DarkGray,
        LogLevel.Debug   => ConsoleColor.Gray,
        LogLevel.Info    => ConsoleColor.Green,
        LogLevel.Warning => ConsoleColor.Yellow,
        LogLevel.Error   => ConsoleColor.Red,
        _                => ConsoleColor.White
    };

    // 返回消息正文对应的控制台前景色
    private static ConsoleColor MessageColor(LogLevel level) => level switch
    {
        LogLevel.Trace   => ConsoleColor.DarkGray,
        LogLevel.Warning => ConsoleColor.Yellow,
        LogLevel.Error   => ConsoleColor.Red,
        _                => ConsoleColor.White
    };

    /// <summary>当前日志文件的完整路径，可用于状态栏显示或用户导航。</summary>
    public static string? CurrentLogFilePath => _logFilePath;
}