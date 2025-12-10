using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.ServiceProcess;
using System.Timers;
using System.Management.Automation;

namespace NexusAgent
{
    public partial class NexusAgentService : ServiceBase
    {
        private Timer _pollTimer;
        private Timer _heartbeatTimer;
        private HttpClient _httpClient;
        private AgentConfig _config;

        public NexusAgentService()
        {
            if (OperatingSystem.IsWindows())
            {
                ServiceName = "NEXUS Agent";
                CanStop = true;
                CanPauseAndContinue = false;
                AutoLog = true;
            }
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                LogEvent("Starting NEXUS Agent...");
                _config = ExtractEmbeddedConfig();
                LogEvent($"Connected to: {_config.deployment_id}");
                
                _httpClient = new HttpClient();
                _httpClient.DefaultRequestHeaders.Add("apikey", _config.supabase_key);
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_config.supabase_key}");

                SendHeartbeat();

                _pollTimer = new Timer(30000);
                _pollTimer.Elapsed += PollForCommands;
                _pollTimer.Start();

                _heartbeatTimer = new Timer(300000);
                _heartbeatTimer.Elapsed += (s, e) => SendHeartbeat();
                _heartbeatTimer.Start();

                LogEvent("NEXUS Agent started successfully");
            }
            catch (Exception ex)
            {
                LogEvent($"Failed to start: {ex.Message}");
                throw;
            }
        }

        protected override void OnStop()
        {
            _pollTimer?.Stop();
            _heartbeatTimer?.Stop();
            _httpClient?.Dispose();
            LogEvent("NEXUS Agent stopped");
        }

        private AgentConfig ExtractEmbeddedConfig()
        {
            try
            {
                string exePath = Environment.ProcessPath;
                if (string.IsNullOrEmpty(exePath))
                {
                    throw new Exception("Cannot find executable path");
                }

                byte[] exeBytes = File.ReadAllBytes(exePath);
                
                string placeholder = "<<<NEXUS_CREDENTIALS_PLACEHOLDER>>>";
                byte[] placeholderBytes = Encoding.UTF8.GetBytes(placeholder);
                
                int index = FindPattern(exeBytes, placeholderBytes);
                if (index == -1)
                {
                    throw new Exception("Config placeholder not found");
                }

                byte[] configBytes = new byte[500];
                Array.Copy(exeBytes, index, configBytes, 0, 500);
                
                string configJson = Encoding.UTF8.GetString(configBytes).TrimEnd(' ', '\0');
                var config = JsonSerializer.Deserialize<AgentConfig>(configJson);
                
                if (config == null)
                {
                    throw new Exception("Failed to deserialize config");
                }
                
                return config;
            }
            catch (Exception ex)
            {
                LogEvent($"Failed to extract config: {ex.Message}");
                throw;
            }
        }

        private void PollForCommands(object sender, ElapsedEventArgs e)
        {
            try
            {
                string url = $"{_config.supabase_url}/rest/v1/agent_commands?status=eq.pending&select=*";
                var response = _httpClient.GetAsync(url).Result;
                
                if (!response.IsSuccessStatusCode) return;

                string json = response.Content.ReadAsStringAsync().Result;
                var commands = JsonSerializer.Deserialize<AgentCommand[]>(json);

                if (commands != null)
                {
                    foreach (var command in commands)
                    {
                        ExecuteCommand(command);
                    }
                }
            }
            catch (Exception ex)
            {
                LogEvent($"Poll error: {ex.Message}");
            }
        }

        private void ExecuteCommand(AgentCommand command)
        {
            try
            {
                LogEvent($"Executing: {command.command_type}");
                UpdateCommandStatus(command.id, "running", "", "");

                string output = "";
                string error = "";

                if (command.command_type == "powershell")
                {
                    (output, error) = ExecutePowerShell(command.script ?? "");
                }
                else
                {
                    error = $"Unknown command: {command.command_type}";
                }

                bool success = string.IsNullOrEmpty(error);
                UpdateCommandStatus(command.id, success ? "completed" : "failed", output, error);
            }
            catch (Exception ex)
            {
                UpdateCommandStatus(command.id, "failed", "", ex.Message);
            }
        }

        private (string output, string error) ExecutePowerShell(string script)
        {
            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddScript(script);
                    var results = ps.Invoke();
                    
                    var output = new StringBuilder();
                    foreach (var result in results)
                    {
                        output.AppendLine(result?.ToString() ?? "");
                    }

                    var errors = new StringBuilder();
                    foreach (var error in ps.Streams.Error)
                    {
                        errors.AppendLine(error.ToString());
                    }

                    return (output.ToString(), errors.ToString());
                }
            }
            catch (Exception ex)
            {
                return ("", ex.Message);
            }
        }

        private void UpdateCommandStatus(string commandId, string status, string output, string error)
        {
            try
            {
                var update = new
                {
                    status = status,
                    output = output,
                    error = error,
                    executed_at = DateTime.UtcNow.ToString("o")
                };

                string json = JsonSerializer.Serialize(update);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                string url = $"{_config.supabase_url}/rest/v1/agent_commands?id=eq.{commandId}";
                _httpClient.PatchAsync(url, content).Wait();
            }
            catch { }
        }

        private void SendHeartbeat()
        {
            try
            {
                var heartbeat = new
                {
                    last_seen = DateTime.UtcNow.ToString("o"),
                    hostname = Environment.MachineName,
                    agent_version = "1.0.0"
                };

                string json = JsonSerializer.Serialize(heartbeat);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                string url = $"{_config.supabase_url}/rest/v1/agent_keys?secret_key=eq.{_config.secret_key}&deployment_id=eq.{_config.deployment_id}";
                _httpClient.PatchAsync(url, content).Wait();
            }
            catch { }
        }

        private int FindPattern(byte[] data, byte[] pattern)
        {
            for (int i = 0; i < data.Length - pattern.Length; i++)
            {
                bool found = true;
                for (int j = 0; j < pattern.Length; j++)
                {
                    if (data[i + j] != pattern[j])
                    {
                        found = false;
                        break;
                    }
                }
                if (found) return i;
            }
            return -1;
        }

        private void LogEvent(string message)
        {
            try
            {
                string logPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "NEXUS",
                    "agent.log"
                );
                
                string directory = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                
                File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n");
            }
            catch { }
        }
    }

    public class AgentConfig
    {
        public string supabase_url { get; set; } = "";
        public string supabase_key { get; set; } = "";
        public string deployment_id { get; set; } = "";
        public string secret_key { get; set; } = "";
    }

    public class AgentCommand
    {
        public string id { get; set; } = "";
        public string command_type { get; set; } = "";
        public string script { get; set; } = "";
    }
}
