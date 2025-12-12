using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using Newtonsoft.Json;

namespace NexusAgent
{
    public class NexusAgentService : ServiceBase
    {
        private Timer _pollTimer;
        private Timer _heartbeatTimer;
        private HttpClient _httpClient;
        private AgentConfig _config;
        private string _agentId;
        private bool _consoleMode;

        public NexusAgentService()
        {
            ServiceName = "NEXUS Agent";
            CanStop = true;
            CanPauseAndContinue = false;
            AutoLog = true;
            _consoleMode = false;
        }

        public void StartConsoleMode()
        {
            _consoleMode = true;
            OnStart(null);
        }

        public void StopConsoleMode()
        {
            OnStop();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                LogEvent("Starting NEXUS Agent...");
                _config = ExtractEmbeddedConfig();
                
                if (_config == null)
                {
                    LogEvent("ERROR: Failed to extract configuration");
                    return;
                }
                
                LogEvent("Connected to deployment: " + _config.deployment_id);

                _httpClient = new HttpClient();
                _httpClient.DefaultRequestHeaders.Add("apikey", _config.supabase_key);
                _httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + _config.supabase_key);

                _agentId = RegisterAgent();
                if (string.IsNullOrEmpty(_agentId))
                {
                    LogEvent("ERROR: Failed to register agent");
                    return;
                }

                LogEvent("Agent registered with ID: " + _agentId);

                SendHeartbeat();

                _pollTimer = new Timer(30000);
                _pollTimer.Elapsed += PollForCommands;
                _pollTimer.Start();

                _heartbeatTimer = new Timer(300000);
                _heartbeatTimer.Elapsed += HeartbeatTimerElapsed;
                _heartbeatTimer.Start();

                LogEvent("NEXUS Agent started successfully");
            }
            catch (Exception ex)
            {
                LogEvent("Failed to start: " + ex.Message);
                throw;
            }
        }

        private void HeartbeatTimerElapsed(object sender, ElapsedEventArgs e)
        {
            SendHeartbeat();
        }

        protected override void OnStop()
        {
            if (_pollTimer != null)
            {
                _pollTimer.Stop();
                _pollTimer.Dispose();
            }
            if (_heartbeatTimer != null)
            {
                _heartbeatTimer.Stop();
                _heartbeatTimer.Dispose();
            }
            if (_httpClient != null)
            {
                _httpClient.Dispose();
            }
            LogEvent("NEXUS Agent stopped");
        }

        private AgentConfig ExtractEmbeddedConfig()
        {
            try
            {
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                
                if (string.IsNullOrEmpty(exePath) || !File.Exists(exePath))
                {
                    LogEvent("Cannot find executable path");
                    return null;
                }

                byte[] exeBytes = File.ReadAllBytes(exePath);
                string fileContent = Encoding.UTF8.GetString(exeBytes);

                int startIndex = -1;
                for (int i = 0; i < exeBytes.Length - 20; i++)
                {
                    if (exeBytes[i] == (byte)'{' && exeBytes[i + 1] == (byte)'"')
                    {
                        int remaining = exeBytes.Length - i;
                        int checkLength = remaining < 50 ? remaining : 50;
                        string testStr = Encoding.UTF8.GetString(exeBytes, i, checkLength);
                        if (testStr.StartsWith("{\"supabase_url\":"))
                        {
                            startIndex = i;
                            break;
                        }
                    }
                }

                if (startIndex == -1)
                {
                    if (fileContent.Contains("<<<NEXUS_CREDENTIALS_PLACEHOLDER>>>"))
                    {
                        LogEvent("Agent has not been configured - placeholder still present");
                    }
                    else
                    {
                        LogEvent("Could not find configuration in executable");
                    }
                    return null;
                }

                int endIndex = startIndex;
                int braceCount = 0;
                for (int i = startIndex; i < exeBytes.Length; i++)
                {
                    if (exeBytes[i] == (byte)'{')
                    {
                        braceCount++;
                    }
                    if (exeBytes[i] == (byte)'}')
                    {
                        braceCount--;
                    }
                    if (braceCount == 0)
                    {
                        endIndex = i + 1;
                        break;
                    }
                }

                string jsonStr = Encoding.UTF8.GetString(exeBytes, startIndex, endIndex - startIndex);
                int logLength = jsonStr.Length < 80 ? jsonStr.Length : 80;
                LogEvent("Found config: " + jsonStr.Substring(0, logLength) + "...");

                AgentConfig config = JsonConvert.DeserializeObject<AgentConfig>(jsonStr);
                return config;
            }
            catch (Exception ex)
            {
                LogEvent("Failed to extract config: " + ex.Message);
                return null;
            }
        }

        private string RegisterAgent()
        {
            try
            {
                string hostname = Environment.MachineName;
                string baseUrl = _config.supabase_url.TrimEnd('/');

                string selectUrl = baseUrl + "/rest/v1/agent_installations?agent_key=eq." + _config.secret_key + "&select=id";
                
                HttpResponseMessage response = _httpClient.GetAsync(selectUrl).Result;
                string content = response.Content.ReadAsStringAsync().Result;

                if (response.IsSuccessStatusCode && content != "[]")
                {
                    List<AgentRecord> records = JsonConvert.DeserializeObject<List<AgentRecord>>(content);
                    if (records != null && records.Count > 0)
                    {
                        string updateUrl = baseUrl + "/rest/v1/agent_installations?agent_key=eq." + _config.secret_key;
                        Dictionary<string, object> updateData = new Dictionary<string, object>();
                        updateData.Add("hostname", hostname);
                        updateData.Add("status", "connected");
                        updateData.Add("last_heartbeat", DateTime.UtcNow.ToString("o"));
                        updateData.Add("agent_version", "1.0.0");

                        StringContent updateContent = new StringContent(
                            JsonConvert.SerializeObject(updateData),
                            Encoding.UTF8,
                            "application/json"
                        );

                        HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl);
                        request.Content = updateContent;

                        _httpClient.SendAsync(request).Wait();
                        LogEvent("Updated agent registration for " + hostname);
                        return records[0].id;
                    }
                }

                LogEvent("Warning: No existing agent registration found");
                return null;
            }
            catch (Exception ex)
            {
                LogEvent("Error registering agent: " + ex.Message);
                return null;
            }
        }

        private void SendHeartbeat()
        {
            try
            {
                string baseUrl = _config.supabase_url.TrimEnd('/');
                string updateUrl = baseUrl + "/rest/v1/agent_installations?agent_key=eq." + _config.secret_key;

                Dictionary<string, object> updateData = new Dictionary<string, object>();
                updateData.Add("last_heartbeat", DateTime.UtcNow.ToString("o"));
                updateData.Add("status", "connected");
                updateData.Add("hostname", Environment.MachineName);

                StringContent content = new StringContent(
                    JsonConvert.SerializeObject(updateData),
                    Encoding.UTF8,
                    "application/json"
                );

                HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl);
                request.Content = content;

                HttpResponseMessage response = _httpClient.SendAsync(request).Result;
                
                if (response.IsSuccessStatusCode)
                {
                    LogEvent("Heartbeat sent");
                }
                else
                {
                    string error = response.Content.ReadAsStringAsync().Result;
                    LogEvent("Heartbeat failed: " + response.StatusCode + " - " + error);
                }
            }
            catch (Exception ex)
            {
                LogEvent("Error sending heartbeat: " + ex.Message);
            }
        }

        private void PollForCommands(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_agentId))
                {
                    return;
                }

                string baseUrl = _config.supabase_url.TrimEnd('/');
                string url = baseUrl + "/rest/v1/agent_commands?agent_id=eq." + _agentId + "&status=eq.pending&select=*";
                
                HttpResponseMessage response = _httpClient.GetAsync(url).Result;

                if (!response.IsSuccessStatusCode)
                {
                    return;
                }

                string json = response.Content.ReadAsStringAsync().Result;
                List<AgentCommand> commands = JsonConvert.DeserializeObject<List<AgentCommand>>(json);

                if (commands != null && commands.Count > 0)
                {
                    LogEvent("Found " + commands.Count + " pending command(s)");
                    foreach (AgentCommand command in commands)
                    {
                        ExecuteCommand(command);
                    }
                }
            }
            catch (Exception ex)
            {
                LogEvent("Error polling commands: " + ex.Message);
            }
        }

        private void ExecuteCommand(AgentCommand command)
        {
            LogEvent("Executing command: " + command.command_type);
            
            UpdateCommandStatus(command.id, "running", null, null);

            string output = null;
            string error = null;

            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddScript(command.script);
                    System.Collections.ObjectModel.Collection<PSObject> results = ps.Invoke();

                    StringBuilder outputBuilder = new StringBuilder();
                    foreach (PSObject item in results)
                    {
                        if (item != null)
                        {
                            outputBuilder.AppendLine(item.ToString());
                        }
                        else
                        {
                            outputBuilder.AppendLine("");
                        }
                    }
                    output = outputBuilder.ToString();

                    if (ps.HadErrors)
                    {
                        StringBuilder errorBuilder = new StringBuilder();
                        foreach (ErrorRecord err in ps.Streams.Error)
                        {
                            errorBuilder.AppendLine(err.ToString());
                        }
                        error = errorBuilder.ToString();
                    }
                }

                UpdateCommandStatus(command.id, "completed", output, error);
                LogEvent("Command completed: " + command.command_type);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                UpdateCommandStatus(command.id, "failed", output, error);
                LogEvent("Command failed: " + ex.Message);
            }
        }

        private void UpdateCommandStatus(string commandId, string status, string output, string error)
        {
            try
            {
                string baseUrl = _config.supabase_url.TrimEnd('/');
                string updateUrl = baseUrl + "/rest/v1/agent_commands?id=eq." + commandId;

                Dictionary<string, object> updateData = new Dictionary<string, object>();
                updateData.Add("status", status);

                if (output != null)
                {
                    updateData.Add("output", output);
                }
                if (error != null)
                {
                    updateData.Add("error", error);
                }

                if (status == "running")
                {
                    updateData.Add("started_at", DateTime.UtcNow.ToString("o"));
                }
                else if (status == "completed" || status == "failed")
                {
                    updateData.Add("completed_at", DateTime.UtcNow.ToString("o"));
                }

                StringContent content = new StringContent(
                    JsonConvert.SerializeObject(updateData),
                    Encoding.UTF8,
                    "application/json"
                );

                HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl);
                request.Content = content;

                _httpClient.SendAsync(request).Wait();
            }
            catch (Exception ex)
            {
                LogEvent("Error updating command status: " + ex.Message);
            }
        }

        private void LogEvent(string message)
        {
            string logMessage = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + message;
            
            try
            {
                string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "NEXUS Agent");
                Directory.CreateDirectory(logDir);
                string logFile = Path.Combine(logDir, "agent.log");
                
                if (File.Exists(logFile))
                {
                    FileInfo fileInfo = new FileInfo(logFile);
                    if (fileInfo.Length > 10 * 1024 * 1024)
                    {
                        File.Delete(logFile);
                    }
                }
                
                File.AppendAllText(logFile, logMessage + Environment.NewLine);
            }
            catch
            {
                // Ignore logging errors
            }

            if (_consoleMode || Environment.UserInteractive)
            {
                Console.WriteLine(logMessage);
            }
        }
    }

    public class AgentConfig
    {
        public string supabase_url { get; set; }
        public string supabase_key { get; set; }
        public string deployment_id { get; set; }
        public string secret_key { get; set; }
    }

    public class AgentRecord
    {
        public string id { get; set; }
    }

    public class AgentCommand
    {
        public string id { get; set; }
        public string command_type { get; set; }
        public string script { get; set; }
    }
}
