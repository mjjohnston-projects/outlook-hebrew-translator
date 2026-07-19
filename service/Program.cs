using System.Diagnostics;
using System.ServiceProcess;

var options = ServiceOptions.Parse(args);

if (Environment.UserInteractive)
{
    Console.Error.WriteLine("This executable is managed by Windows Services. Use services.msc to start or stop it.");
    return;
}

ServiceBase.Run(new NodeServerService(options));

sealed class NodeServerService : ServiceBase
{
    private const string ServiceIdentifier = "OutlookEmailLanguageAssistant";
    private readonly ServiceOptions options;
    private Process? nodeProcess;
    private readonly string logDirectory;

    public NodeServerService(ServiceOptions options)
    {
        this.options = options;
        logDirectory = Path.Combine(options.AppRoot, "logs");
        ServiceName = ServiceIdentifier;
        CanStop = true;
        AutoLog = true;
    }

    protected override void OnStart(string[] args)
    {
        Directory.CreateDirectory(logDirectory);
        var startInfo = new ProcessStartInfo
        {
            FileName = options.NodePath,
            Arguments = "server.mjs",
            WorkingDirectory = options.AppRoot,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.Environment["TLS_CERT_PATH"] = options.CertificatePath;
        startInfo.Environment["TLS_KEY_PATH"] = options.KeyPath;

        nodeProcess = new Process { StartInfo = startInfo, EnableRaisingEvents = true };
        nodeProcess.OutputDataReceived += (_, eventArgs) => WriteLog("service.log", eventArgs.Data);
        nodeProcess.ErrorDataReceived += (_, eventArgs) => WriteLog("service.error.log", eventArgs.Data);
        nodeProcess.Start();
        nodeProcess.BeginOutputReadLine();
        nodeProcess.BeginErrorReadLine();
    }

    protected override void OnStop()
    {
        if (nodeProcess is { HasExited: false })
        {
            nodeProcess.Kill(entireProcessTree: true);
            nodeProcess.WaitForExit(10_000);
        }
        nodeProcess?.Dispose();
        nodeProcess = null;
    }

    private void WriteLog(string fileName, string? message)
    {
        if (string.IsNullOrWhiteSpace(message)) return;
        File.AppendAllText(Path.Combine(logDirectory, fileName), $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}");
    }
}

sealed record ServiceOptions(string AppRoot, string NodePath, string CertificatePath, string KeyPath)
{
    public static ServiceOptions Parse(string[] args)
    {
        string? GetOption(string name)
        {
            var index = Array.FindIndex(args, item => string.Equals(item, name, StringComparison.OrdinalIgnoreCase));
            return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
        }

        var appRoot = GetOption("--app-root") ?? throw new ArgumentException("Missing --app-root.");
        var nodePath = GetOption("--node-path") ?? throw new ArgumentException("Missing --node-path.");
        var certificatePath = GetOption("--certificate") ?? throw new ArgumentException("Missing --certificate.");
        var keyPath = GetOption("--key") ?? throw new ArgumentException("Missing --key.");
        return new ServiceOptions(appRoot, nodePath, certificatePath, keyPath);
    }
}
