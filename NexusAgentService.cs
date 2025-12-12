using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NexusAgent
{
    public class NexusAgentService : ServiceBase
    {
        private CancellationTokenSource _cancellationTokenSource;
        private Task _workerTask;
        private AgentConfig _config;
        private HttpClient _httpClient;
        private string _agentId;

        public NexusAgentService()
        {
            ServiceName = Program.ServiceName;
            _httpClient = new HttpClient();
        }

        protected override void OnStart(string[] args)
        {
            _cancellationTokenSource = new CancellationTokenSource();
            _workerTask = RunAgentAsync(_cancellationTokenSource.Token);
        }

        protected override void OnStop()
        {
            if (_cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
            }
            if (_workerTask != null)
            {
                _workerTask.Wait(TimeSpan.FromSeconds(30));
            }
        }

        public void StartConsoleMode()
        {
            _cancellationTokenSource = new CancellationTokenSource();
            _workerTask = RunAgentAsync(_cancellationTokenSource.Token);
        }

        public void StopConsoleMode()
        {
            if (_cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
            }
            try 
            { 
                if (_workerTask != null)
                {
                    _workerTask.Wait(TimeSpan.FromSeconds(10)); 
                }
            } 
            catch { }
        }

        private async Task RunAgentAsync(CancellationToken cancellationToken)
        {
            try
            {
                LogMessage("NEXUS Agent starting...");

                _config = ExtractConfig();
                if (_config == null)
                {
                    LogMessage("ERROR: Failed to extract configuration. Agent cannot start.");
                    return;
                }

                LogMessage("Configuration loaded for deployment: " + _config.deployment_id);

                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("apikey", _config.supabase_key);
                _httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + _config.supabase_key);

                _agentId = await RegisterAgentAsync();
                if (string.IsNullOrEmpty(_agentId))
                {
                    LogMessage("ERROR: Failed to register agent.");
                    return;
                }

                LogMessage("Agent registered with ID: " + _agentId);

                while (!cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        await SendHeartbeatAsync();
                        await PollAndExecuteCommandsAsync();
                    }
                    catch (Exception ex)
                    {
                        LogMessage("Error in main loop: " + ex.Message);
                    }

                    try
                    {
                        await Task.Delay(TimeSpan.FromSeconds(30), cancellationToken);
                    }
                    catch (TaskCanceledException)
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage("Fatal error: " + ex.Message);
            }
            
            LogMessage("Agent stopped.");
        }

        private AgentConfig ExtractConfig()
        {
            try
            {
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

                if (!File.Exists(exePath))
                {
                    LogMessage("Could not find executable path");
                    return null;
                }

                byte[] fileBytes = File.ReadAllBytes(exePath);
                string fileContent = Encoding.UTF8.GetString(fileBytes);

                // Look for JSON config pattern
                int startIndex = -1;
                for (int i = 0; i < fileBytes.Length - 20; i++)
                {
                    if (fileBytes[i] == (byte)'{' && fileBytes[i + 1] == (byte)'"' && fileBytes[i + 2] == (byte)'s')
                    {
                        string testStr = Encoding.UTF8.GetString(fileBytes, i, Math.Min(50, fileBytes.Length - i));
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
                        LogMessage("Agent has not been configured - placeholder still present");
                    }
                    else
                    {
                        LogMessage("Could not find configuration in executable");
                    }
                    return null;
                }

                // Find end of JSON
                int endIndex = startIndex;
                int braceCount = 0;
                for (int i = startIndex; i < fileBytes.Length; i++)
                {
                    if (fileBytes[i] == (byte)'{') braceCount++;
                    if (fileBytes[i] == (byte)'}') braceCount--;
                    if (braceCount == 0)
                    {
                        endIndex = i + 1;
                        break;
                    }
                }

                string jsonStr = Encoding.UTF8.GetString(fileBytes, startIndex, endIndex - startIndex);
                LogMessage("Found config, length: " + jsonStr.Length);

                AgentConfig config = JsonConvert.DeserializeObject<AgentConfig>(jsonStr);
                return config;
            }
            catch (Exception ex)
            {
                LogMessage("Error extracting config: " + ex.Message);
                return null;
            }
        }

        private async Task<string> RegisterAgentAsync()
        {
            try
            {
                string hostname = Environment.MachineName;
                string baseUrl = _config.supabase_url.TrimEnd('/');

                // Find existing registration by secret key
                string selectUrl = baseUrl + "/rest/v1/agent_installations?agent_key=eq." + _config.secret_key + "&select=id";
                
                HttpResponseMessage response = await _httpClient.GetAsync(selectUrl);
                string content = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode && content != "[]")
                {
                    List<AgentRecord> records = JsonConvert.DeserializeObject<List<AgentRecord>>(content);
                    if (records != null && records.Count > 0)
                    {
                        // Update existing record
                        string updateUrl = baseUrl + "/rest/v1/agent_installations?agent_key=eq." + _config.secret_key;
                        
                        var updateData = new Dictionary<string, object>();
                        updateData["hostname"] = hostname;
                        updateData["status"] = "connected";
                        updateData["last_heartbeat"] = DateTime.UtcNow.ToString("o");
                        updateData["agent_version"] = "1.0.0";
                        updateData["is_active"] = true;

                        StringContent updateContent = new StringContent(
                            JsonConvert.SerializeObject(updateData),
                            Encoding.UTF8,
                            "application/json"
                        );

                        HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl);
                        request.Content = updateContent;

                        await _httpClient.SendAsync(request);
                        LogMessage("Updated existing agent registration for " + hostname);
                        return records[0].id;
                    }
                }

                LogMessage("Warning: No existing agent registration found for this secret key");
                return null;
            }
            catch (Exception ex)
            {
                LogMessage("Error registering agent: " + ex.Message);
                return null;
            }
        }

        private async Task SendHeartbeatAsync()
        {
            try
            {
                string baseUrl = _config.supabase_url.TrimEnd('/');
                string updateUrl = baseUrl + "/rest/v1/agent_installations?agent_key=eq." + _config.secret_key;

                var updateData = new Dictionary<string, object>();
                updateData["last_heartbeat"] = DateTime.UtcNow.ToString("o");
                updateData["status"] = "connected";
                updateData["hostname"] = Environment.MachineName;

                StringContent content = new StringContent(
                    JsonConvert.SerializeObject(updateData),
                    Encoding.UTF8,
                    "application/json"
                );

                HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl);
                request.Content = content;

                HttpResponseMessage response = await _httpClient.SendAsync(request);
                
                if (response.IsSuccessStatusCode)
                {
                    LogMessage("Heartbeat sent successfully");
                }
                else
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    LogMessage("Heartbeat failed: " + response.StatusCode + " - " + errorContent);
                }
            }
            catch (Exception ex)
            {
                LogMessage("Error sending heartbeat: " + ex.Message);
            }
        }

        private async Task PollAndExecuteCommandsAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_agentId)) return;

                string baseUrl = _config.supabase_url.TrimEnd('/');
                string selectUrl = baseUrl + "/rest/v1/agent_commands?agent_id=eq." + _agentId + "&status=eq.pending&select=*";

                HttpResponseMessage response = await _httpClient.GetAsync(selectUrl);
                if (!response.IsSuccessStatusCode) return;

                string content = await response.Content.ReadAsStringAsync();
                List<AgentCommand> commands = JsonConvert.DeserializeObject<List<AgentCommand>>(content);

                if (commands == null || commands.Count == 0) return;

                LogMessage("Found " + commands.Count + " pending command(s)");

                foreach (AgentCommand command in commands)
                {
                    await ExecuteCommandAsync(command);
                }
            }
            catch (Exception ex)
            {
                LogMessage("Error polling commands: " + ex.Message);
            }
        }

        private async Task ExecuteCommandAsync(AgentCommand command)
        {
            LogMessage("Executing command: " + command.command_type);
            
            await UpdateCommandStatusAsync(command.id, "running", null, null);

            string output = null;
            string error = null;

            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddScript(command.script);

                    var results = ps.Invoke();

                    StringBuilder outputBuilder = new StringBuilder();
                    foreach (var item in results)
                    {
                        outputBuilder.AppendLine(item != null ? item.ToString() : "");
                    }
                    output = outputBuilder.ToString();

                    if (ps.HadErrors)
                    {
                        StringBuilder errorBuilder = new StringBuilder();
                        foreach (var err in ps.Streams.Error)
                        {
                            errorBuilder.AppendLine(err.ToString());
                        }
                        error = errorBuilder.ToString();
                    }
                }

                await UpdateCommandStatusAsync(command.id, "completed", output, error);
                LogMessage("Command completed: " + command.command_type);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                await UpdateCommandStatusAsync(command.id, "failed", output, error);
                LogMessage("Command failed: " + ex.Message);
            }
        }

        private async Task UpdateCommandStatusAsync(string commandId, string status, string output, string error)
        {
            try
            {
                string baseUrl = _config.supabase_url.TrimEnd('/');
                string updateUrl = baseUrl + "/rest/v1/agent_commands?id=eq." + commandId;

                var updateData = new Dictionary<string, object>();
                updateData["status"] = status;

                if (output != null) updateData["output"] = output;
                if (error != null) updateData["error"] = error;

                if (status == "running")
                {
                    updateData["started_at"] = DateTime.UtcNow.ToString("o");
                }
                else if (status == "completed" || status == "failed")
                {
                    updateData["completed_at"] = DateTime.UtcNow.ToString("o");
                }

                StringContent content = new StringContent(
                    JsonConvert.SerializeObject(updateData),
                    Encoding.UTF8,
                    "application/json"
                );

                HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl);
                request.Content = content;

                await _httpClient.SendAsync(request);
            }
            catch (Exception ex)
            {
                LogMessage("Error updating command status: " + ex.Message);
            }
        }

        private void LogMessage(string message)
        {
            string logMessage = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + message;
            
            try
            {
                string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "NEXUS Agent");
                Directory.CreateDirectory(logDir);
                string logFile = Path.Combine(logDir, "agent.log");
                
                FileInfo fi = new FileInfo(logFile);
                if (fi.Exists && fi.Length > 10 * 1024 * 1024)
                {
                    File.Delete(logFile);
                }
                
                File.AppendAllText(logFile, logMessage + Environment.NewLine);
            }
            catch { }

            if (Environment.UserInteractive)
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
