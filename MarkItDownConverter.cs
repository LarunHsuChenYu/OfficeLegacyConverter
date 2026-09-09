using System.Diagnostics;
using System.Text.RegularExpressions;

namespace OfficeLegacyConverter;

internal static class MarkItDownConverter
{
    private static readonly string[] PythonCommands = OperatingSystem.IsWindows()
        ? ["py -3", "python3", "python"]
        : ["python3", "python"];

    private static string? _resolvedPythonCommand;
    private static string? _environmentStatus;

    public static string? ResolvedPythonCommand => _resolvedPythonCommand;
    public static string EnvironmentStatus => _environmentStatus ?? "尚未檢查";

    public static bool IsEnvironmentReady { get; private set; }

    public static void CheckEnvironment()
    {
        _resolvedPythonCommand = null;
        _environmentStatus = null;
        IsEnvironmentReady = false;

        foreach (var command in PythonCommands)
        {
            if (!TryRunCommand(command, "--version", out var versionOutput, out var versionError))
            {
                if (!string.IsNullOrWhiteSpace(versionError))
                    _environmentStatus = $"嘗試 {command} 失敗：{versionError.Trim()}";
                continue;
            }

            if (!TryParsePythonVersion(versionOutput, out var version) || version < new Version(3, 10))
            {
                _environmentStatus = $"找到 {command}，但版本不符合需求（需要 Python 3.10+）：{versionOutput.Trim()}";
                continue;
            }

            if (IsWindowsStorePythonStub(command))
            {
                _environmentStatus = $"略過 Windows Store Python 占位：{command}";
                continue;
            }

            if (!TryRunCommand(command, "-m markitdown --help", out _, out var markitdownError)
                && !TryRunCommand(command, "-m markitdown -h", out _, out markitdownError))
            {
                _environmentStatus = $"找到 {command} {versionOutput.Trim()}，但 markitdown 模組不可用：{markitdownError.Trim()}";
                continue;
            }

            _resolvedPythonCommand = command;
            IsEnvironmentReady = true;
            _environmentStatus = $"就緒：{command}（{versionOutput.Trim()}），markitdown 可用";
            return;
        }

        _environmentStatus ??= "找不到可用的 Python 3.10+ 或 markitdown 模組。請確認已安裝：pip install markitdown";
    }

    public static void ConvertFile(string inputPath, string outputMdPath, CancellationToken cancellationToken)
    {
        if (!IsEnvironmentReady || _resolvedPythonCommand is null)
            throw new InvalidOperationException(EnvironmentStatus);

        var outputDir = Path.GetDirectoryName(outputMdPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        var arguments = $"-m markitdown \"{inputPath}\" -o \"{outputMdPath}\"";
        RunPython(arguments, cancellationToken);
    }

    private static void RunPython(string arguments, CancellationToken cancellationToken)
    {
        var command = _resolvedPythonCommand
            ?? throw new InvalidOperationException(EnvironmentStatus);

        var (fileName, args) = SplitCommand(command, arguments);

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8
            }
        };

        process.Start();

        using var registration = cancellationToken.Register(() =>
        {
            try
            {
                if (!process.HasExited)
                    process.Kill(entireProcessTree: true);
            }
            catch { /* ignore */ }
        });

        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();

        cancellationToken.ThrowIfCancellationRequested();

        if (process.ExitCode != 0)
        {
            var detail = BuildProcessErrorDetail(stdout, stderr, process.ExitCode);
            throw new InvalidOperationException(detail);
        }
    }

    private static bool TryRunCommand(string command, string arguments, out string stdout, out string stderr)
    {
        stdout = "";
        stderr = "";

        try
        {
            var (fileName, args) = SplitCommand(command, arguments);
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            });

            if (process is null)
                return false;

            stdout = process.StandardOutput.ReadToEnd();
            stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();
            return process.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    private static (string FileName, string Arguments) SplitCommand(string command, string arguments)
    {
        var parts = command.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);
        var fileName = parts[0];
        var commandArgs = parts.Length > 1 ? parts[1] + " " : "";
        return (fileName, commandArgs + arguments);
    }

    private static bool TryParsePythonVersion(string output, out Version version)
    {
        version = new Version(0, 0);
        var match = Regex.Match(output, @"Python\s+(\d+)\.(\d+)");
        if (!match.Success)
            return false;

        version = new Version(int.Parse(match.Groups[1].Value), int.Parse(match.Groups[2].Value));
        return true;
    }

    private static bool IsWindowsStorePythonStub(string command)
    {
        if (!OperatingSystem.IsWindows())
            return false;

        if (!TryRunCommand(command, "-c \"import sys; print(sys.executable)\"", out var executablePath, out _))
            return false;

        var normalized = executablePath.Trim().Replace('/', '\\');
        return normalized.Contains(@"\Microsoft\WindowsApps\python", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildProcessErrorDetail(string stdout, string stderr, int exitCode)
    {
        var parts = new List<string> { $"markitdown 結束代碼 {exitCode}" };
        if (!string.IsNullOrWhiteSpace(stderr))
            parts.Add($"stderr:\n{stderr.Trim()}");
        if (!string.IsNullOrWhiteSpace(stdout))
            parts.Add($"stdout:\n{stdout.Trim()}");
        return parts.Count > 1 ? string.Join(Environment.NewLine + Environment.NewLine, parts) : parts[0];
    }
}
