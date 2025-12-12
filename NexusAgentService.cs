using System;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Newtonsoft.Json;

namespace NexusAgent
{
    public partial class NexusAgentService : ServiceBase
    {
        private Timer _heartbeatTimer;
        private Timer _commandTimer;
        private HttpClient _httpClient;
        private AgentCredentials _credentials;
        private string _agentId;
        private const string PLACEHOLDER = "<<<NEXUS_CREDENTIALS_PLACEHOLDER>>>";
        private const int CREDENTIALS_SIZE = 500;

        public NexusAgentService()
        {
            ServiceName = "NexusAgent";
            CanStop = true;
            CanPauseAndContinue = false;
            AutoLog = true;
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                WriteLog("NEXUS Agent starting...");

                // Load embedded credentials
                _credentials = LoadCredentials();
                if (_credentials == null)
                {
                    WriteLog("ERROR: Failed to load credentials. Agent cannot start.");
                    Stop();
                    return;
                }

                WriteLog($"Credentials loaded for deployment: {_credentials.deployment_id}");

                // Initialize HTTP client for Supabase
                _httpClient = new HttpClient();
                _httpClient.DefaultRequestHeaders.Add("apikey", _credentials.supabase_key);
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_credentials.supabase_key}");

                // Register agent and get agent ID
                Task.Run(async () => await RegisterAgent()).Wait();

                // Start heartbeat timer (every 5 minutes)
                _heartbeatTimer = new Timer(300000); // 5 minutes
                _heartbeatTimer.Elapsed += async (s, e) => await SendHeartbeat();
                _heartbeatTimer.AutoReset = true;
                _heartbeatTimer.Start();

                // Start command polling timer (every 30 seconds)
                _commandTimer = new Timer(30000);
                _commandTimer.Elapsed += async (s, e) => await PollCommands();
                _commandTimer.AutoReset = true;
                _commandTimer.Start();

                // Send initial heartbeat
                Task.Run(async () => await SendHeartbeat()).Wait();

                WriteLog("NEXUS Agent started successfully.");
            }
            catch (Exception ex)
            {
                WriteLog($"ERROR starting agent: {ex.Message}");
                Stop();
            }
        }

        protected override void OnStop()
        {
            WriteLog("NEXUS Agent stopping...");
            _heartbeatTimer?.Stop();
            _heartbeatTimer?.Dispose();
            _commandTimer?.Stop();
            _commandTimer?.Dispose();
            _httpClient?.Dispose();
            WriteLog("NEXUS Agent stopped.");
        }

        private AgentCredentials LoadCredentials()
        {
            try
            {
                // Get executable path (.NET Framework 4.8 compatible)
                string exePath = Assembly.GetExecutingAssembly().Location;
                WriteLog($"Loading credentials from: {exePath}");

                byte[] exeBytes = File.ReadAllBytes(exePath);
                WriteLog($"Executable size: {exeBytes.Length} bytes");

                // Find the placeholder in the binary
                string exeContent = Encoding.ASCII.GetString(exeBytes);
                int placeholderIndex = exeContent.IndexOf(PLACEHOLDER);

                if (placeholderIndex >= 0)
                {
                    WriteLog($"Found placeholder at offset {placeholderIndex} - this is an unpatched template");
                    return null;
                }

                // Read the last CREDENTIALS_SIZE bytes where credentials are patched
                if (exeBytes.Length < CREDENTIALS_SIZE)
                {
                    WriteLog("Executable too small to contain credentials");
                    return null;
                }

                // The credentials are at the position where the placeholder was originally
                // We need to find the JSON by scanning the end of the file
                // Try reading the last 500 bytes and parse as JSON
                byte[] credentialBytes = new byte[CREDENTIALS_SIZE];
                Array.Copy(exeBytes, exeBytes.Length - CREDENTIALS_SIZE, credentialBytes, 0, CREDENTIALS_SIZE);
                string credentialsJson = Encoding.UTF8.GetString(credentialBytes).Trim();

                // If that doesn't work, scan the file for JSON pattern
                if (!credentialsJson.StartsWith("{"))
                {
                    // Scan backwards for JSON
                    for (int i = exeBytes.Length - 1; i >= Math.Max(0, exeBytes.Length - 10000); i--)
                    {
                        if (exeBytes[i] == '{')
                        {
                            string potential = Encoding.UTF8.GetString(exeBytes, i, Math.Min(CREDENTIALS_SIZE, exeBytes.Length - i)).Trim();
                            if (potential.Contains("supabase_url") && potential.Contains("deployment_id"))
                            {
                                // Find the end of JSON
                                int endBrace = potential.IndexOf('}');
                                if (endBrace > 0)
                                {
                                    credentialsJson = potential.Substring(0, endBrace + 1);
                                    break;
                                }
                            }
                        }
                    }
                }

                WriteLog($"Credentials JSON: {credentialsJson.Substring(0, Math.Min(100, credentialsJson.Length))}...");

                var credentials = JsonConvert.DeserializeObject<AgentCredentials>(credentialsJson);
                if (credentials == null || string.IsNullOrEmpty(credentials.supabase_url))
                {
                    WriteLog("Failed to parse credentials JSON");
                    return null;
                }

                return credentials;
            }
            catch (Exception ex)
            {
                WriteLog($"ERROR loading credentials: {ex.Message}");
                return null;
            }
        }

        private async Task RegisterAgent()
        {
            try
            {
                string hostname = Environment.MachineName;
                string url = $"{_credentials.supabase_url}/rest/v1/agent_installations?agent_key=eq.{_credentials.secret_key}";

                var response = await _httpClient.GetAsync(url);
                var content = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode && content != "[]")
                {
                    // Agent already registered, get the ID
                    var agents = JsonConvert.DeserializeObject<AgentInstallation[]>(content);
                    if (agents != null && agents.Length > 0)
                    {
                        _agentId = agents[0].id;
                        WriteLog($"Agent already registered with ID: {_agentId}");

                        // Update hostname and status
                        await UpdateAgentInfo(hostname);
                        return;
                    }
                }

                WriteLog($"Agent registered successfully.");
            }
            catch (Exception ex)
            {
                WriteLog($"ERROR registering agent: {ex.Message}");
            }
        }

        private async Task UpdateAgentInfo(string hostname)
        {
            try
            {
                string url = $"{_credentials.supabase_url}/rest/v1/agent_installations?agent_key=eq.{_credentials.secret_key}";
                var updateData = new
                {
                    hostname = hostname,
                    status = "active",
                    agent_version = "1.0.0",
                    last_heartbeat = DateTime.UtcNow.ToString("o")
                };

                var jsonContent = new StringContent(JsonConvert.SerializeObject(updateData), Encoding.UTF8, "application/json");
                _httpClient.DefaultRequestHeaders.Remove("Prefer");
                _httpClient.DefaultRequestHeaders.Add("Prefer", "return=representation");

                var response = await _httpClient.PatchAsync(url, jsonContent);
                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    WriteLog($"ERROR updating agent info: {error}");
                }
            }
            catch (Exception ex)
            {
                WriteLog($"ERROR updating agent info: {ex.Message}");
            }
        }

        private async Task SendHeartbeat()
        {
            try
            {
                string url = $"{_credentials.supabase_url}/rest/v1/agent_installations?agent_key=eq.{_credentials.secret_key}";
                var updateData = new
                {
                    last_heartbeat = DateTime.UtcNow.ToString("o"),
                    status = "active"
                };

                var jsonContent = new StringContent(JsonConvert.SerializeObject(updateData), Encoding.UTF8, "application/json");
                
                var request = new HttpRequestMessage(new HttpMethod("PATCH"), url);
                request.Content = jsonContent;
                
                var response = await _httpClient.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    WriteLog("Heartbeat sent successfully.");
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    WriteLog($"ERROR sending heartbeat: {error}");
                }
            }
            catch (Exception ex)
            {
                WriteLog($"ERROR sending heartbeat: {ex.Message}");
            }
        }

        private async Task PollCommands()
        {
            try
            {
                string url = $"{_credentials.supabase_url}/rest/v1/agent_commands?status=eq.pending&agent_id=eq.{_agentId}&order=created_at.asc&limit=1";

                var response = await _httpClient.GetAsync(url);
                if (!response.IsSuccessStatusCode) return;

                var content = await response.Content.ReadAsStringAsync();
                if (content == "[]") return;

                var commands = JsonConvert.DeserializeObject<AgentCommand[]>(content);
                if (commands == null || commands.Length == 0) return;

                var command = commands[0];
                WriteLog($"Executing command: {command.command_type}");

                // Execute command (PowerShell, etc.)
                await ExecuteCommand(command);
            }
            catch (Exception ex)
            {
                WriteLog($"ERROR polling commands: {ex.Message}");
            }
        }

        private async Task ExecuteCommand(AgentCommand command)
        {
            try
            {
                // Mark as running
                await UpdateCommandStatus(command.id, "running", null, null);

                string output = "";
                string error = "";

                // Execute based on command type
                switch (command.command_type.ToLower())
                {
                    case "powershell":
                        (output, error) = ExecutePowerShell(command.script);
                        break;
                    default:
                        error = $"Unknown command type: {command.command_type}";
                        break;
                }

                // Mark as completed
                string status = string.IsNullOrEmpty(error) ? "completed" : "failed";
                await UpdateCommandStatus(command.id, status, output, error);
            }
            catch (Exception ex)
            {
                await UpdateCommandStatus(command.id, "failed", null, ex.Message);
            }
        }

        private (string output, string error) ExecutePowerShell(string script)
        {
            try
            {
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{script}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = System.Diagnostics.Process.Start(startInfo))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();
                    return (output, error);
                }
            }
            catch (Exception ex)
            {
                return ("", ex.Message);
            }
        }

        private async Task UpdateCommandStatus(string commandId, string status, string output, string error)
        {
            try
            {
                string url = $"{_credentials.supabase_url}/rest/v1/agent_commands?id=eq.{commandId}";
                var updateData = new
                {
                    status = status,
                    output = output,
                    error = error,
                    completed_at = (status == "completed" || status == "failed") ? DateTime.UtcNow.ToString("o") : null
                };

                var jsonContent = new StringContent(JsonConvert.SerializeObject(updateData), Encoding.UTF8, "application/json");
                var request = new HttpRequestMessage(new HttpMethod("PATCH"), url);
                request.Content = jsonContent;

                await _httpClient.SendAsync(request);
            }
            catch (Exception ex)
            {
                WriteLog($"ERROR updating command status: {ex.Message}");
            }
        }

        private void WriteLog(string message)
        {
            try
            {
                string logPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "NexusAgent",
                    "agent.log"
                );
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n");
            }
            catch { }
        }
    }

    // Data classes
    public class AgentCredentials
    {
        public string supabase_url { get; set; }
        public string supabase_key { get; set; }
        public string deployment_id { get; set; }
        public string secret_key { get; set; }
        public string generated_at { get; set; }
    }

    public class AgentInstallation
    {
        public string id { get; set; }
        public string deployment_id { get; set; }
        public string hostname { get; set; }
        public string status { get; set; }
    }

    public class AgentCommand
    {
        public string id { get; set; }
        public string command_type { get; set; }
        public string script { get; set; }
        public string status { get; set; }
    }
}
