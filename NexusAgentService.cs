using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace NexusAgent
{
    public class NexusAgentService : ServiceBase
    {
        private CancellationTokenSource? _cancellationTokenSource;
        private Task? _workerTask;
        private AgentConfig? _config;
        private readonly HttpClient _httpClient;
        private string? _agentId;

        public NexusAgentService()
        {
            ServiceName = Program.ServiceName;
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("apikey", "");
            _httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer ");
        }

        protected override void OnStart(string[] args)
        {
            _cancellationTokenSource = new CancellationTokenSource();
            _workerTask = RunAgentAsync(_cancellationTokenSource.Token);
        }

        protected override void OnStop()
        {
            _cancellationTokenSource?.Cancel();
            _workerTask?.Wait(TimeSpan.FromSeconds(30));
        }

        // For console mode debugging
        public void StartConsoleMode()
        {
            _cancellationTokenSource = new CancellationTokenSource();
            _workerTask = RunAgentAsync(_cancellationTokenSource.Token);
        }

        public void StopConsoleMode()
        {
            _cancellationTokenSource?.Cancel();
            _workerTask?.Wait(TimeSpan.FromSeconds(10));
        }

        private async Task RunAgentAsync(CancellationToken cancellationToken)
        {
            try
            {
                LogMessage("NEXUS Agent starting...");

                // Extract embedded configuration
                _config = ExtractConfig();
                if (_config == null)
                {
                    LogMessage("ERROR: Failed to extract configuration. Agent cannot start.");
                    return;
                }

                LogMessage($"Configuration loaded for deployment: {_config.DeploymentId}");

                // Set up HTTP client with credentials
                _httpClient.DefaultRequestHeaders.Remove("apikey");
                _httpClient.DefaultRequestHeaders.Remove("Authorization");
                _httpClient.DefaultRequestHeaders.Add("apikey", _config.SupabaseKey);
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_config.SupabaseKey}");

                // Register/update agent and get agent ID
                _agentId = await RegisterAgentAsync();
                if (string.IsNullOrEmpty(_agentId))
                {
                    LogMessage("ERROR: Failed to register agent.");
                    return;
                }

                LogMessage($"Agent registered with ID: {_agentId}");

                // Main loop
                while (!cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        // Send heartbeat
                        await SendHeartbeatAsync();

                        // Poll for commands
                        await PollAndExecuteCommandsAsync();
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Error in main loop: {ex.Message}");
                    }

                    // Wait before next poll (30 seconds)
                    await Task.Delay(TimeSpan.FromSeconds(30), cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
                LogMessage("Agent stopping...");
            }
            catch (Exception ex)
            {
                LogMessage($"Fatal error: {ex.Message}");
            }
        }

        private AgentConfig? ExtractConfig()
        {
            try
            {
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                
                // For single-file published apps, use the process path
                if (string.IsNullOrEmpty(exePath) || !File.Exists(exePath))
                {
                    exePath = Environment.ProcessPath ?? "";
                }

                if (!File.Exists(exePath))
                {
                    LogMessage("Could not find executable path");
                    return null;
                }

                byte[] fileBytes = File.ReadAllBytes(exePath);
                string fileContent = Encoding.UTF8.GetString(fileBytes);

                // Look for the JSON config that replaced the placeholder
                string marker = "<<<NEXUS_CREDENTIALS_PLACEHOLDER>>>";
                
                // The config should be JSON, look for the pattern
                int startIndex = -1;
                for (int i = 0; i < fileBytes.Length - 20; i++)
                {
                    // Look for {"supabase_url":
                    if (fileBytes[i] == '{' && fileBytes[i + 1] == '"' && fileBytes[i + 2] == 's')
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
                    // Check if placeholder still exists (template not configured)
                    if (fileContent.Contains(marker))
                    {
                        LogMessage("Agent has not been configured - placeholder still present");
                    }
                    else
                    {
                        LogMessage("Could not find configuration in executable");
                    }
                    return null;
                }

                // Find the end of the JSON (look for the closing brace followed by padding or nulls)
                int endIndex = startIndex;
                int braceCount = 0;
                for (int i = startIndex; i < fileBytes.Length; i++)
                {
                    if (fileBytes[i] == '{') braceCount++;
                    if (fileBytes[i] == '}') braceCount--;
                    if (braceCount == 0)
                    {
                        endIndex = i + 1;
                        break;
                    }
                }

                string jsonStr = Encoding.UTF8.GetString(fileBytes, startIndex, endIndex - startIndex);
                LogMessage($"Found config JSON: {jsonStr.Substring(0, Math.Min(100, jsonStr.Length))}...");

                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var config = JsonSerializer.Deserialize<AgentConfig>(jsonStr, options);

                return config;
            }
            catch (Exception ex)
            {
                LogMessage($"Error extracting config: {ex.Message}");
                return null;
            }
        }

        private async Task<string?> RegisterAgentAsync()
        {
            try
            {
                string hostname = Environment.MachineName;
                string baseUrl = _config!.SupabaseUrl.TrimEnd('/');

                // First, try to find existing registration by secret key
                string selectUrl = $"{baseUrl}/rest/v1/agent_installations?agent_key=eq.{_config.SecretKey}&select=id";
                
                var response = await _httpClient.GetAsync(selectUrl);
                var content = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode && content != "[]")
                {
                    // Parse existing record
                    var records = JsonSerializer.Deserialize<List<AgentRecord>>(content);
                    if (records != null && records.Count > 0)
                    {
                        // Update existing record
                        string updateUrl = $"{baseUrl}/rest/v1/agent_installations?agent_key=eq.{_config.SecretKey}";
                        var updateData = new
                        {
                            hostname = hostname,
                            status = "connected",
                            last_heartbeat = DateTime.UtcNow.ToString("o"),
                            agent_version = "1.0.0",
                            is_active = true
                        };

                        var updateContent = new StringContent(
                            JsonSerializer.Serialize(updateData),
                            Encoding.UTF8,
                            "application/json"
                        );

                        // Add Prefer header for upsert
                        var request = new HttpRequestMessage(HttpMethod.Patch, updateUrl)
                        {
                            Content = updateContent
                        };

                        await _httpClient.SendAsync(request);
                        LogMessage($"Updated existing agent registration for {hostname}");
                        return records[0].Id;
                    }
                }

                // No existing record - this shouldn't happen as download-agent creates the record
                // But just in case, log it
                LogMessage("Warning: No existing agent registration found for this secret key");
                return null;
            }
            catch (Exception ex)
            {
                LogMessage($"Error registering agent: {ex.Message}");
                return null;
            }
        }

        private async Task SendHeartbeatAsync()
        {
            try
            {
                string baseUrl = _config!.SupabaseUrl.TrimEnd('/');
                string updateUrl = $"{baseUrl}/rest/v1/agent_installations?agent_key=eq.{_config.SecretKey}";

                var updateData = new
                {
                    last_heartbeat = DateTime.UtcNow.ToString("o"),
                    status = "connected",
                    hostname = Environment.MachineName
                };

                var content = new StringContent(
                    JsonSerializer.Serialize(updateData),
                    Encoding.UTF8,
                    "application/json"
                );

                var request = new HttpRequestMessage(HttpMethod.Patch, updateUrl)
                {
                    Content = content
                };

                var response = await _httpClient.SendAsync(request);
                
                if (response.IsSuccessStatusCode)
                {
                    LogMessage("Heartbeat sent successfully");
                }
                else
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    LogMessage($"Heartbeat failed: {response.StatusCode} - {errorContent}");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error sending heartbeat: {ex.Message}");
            }
        }

        private async Task PollAndExecuteCommandsAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_agentId)) return;

                string baseUrl = _config!.SupabaseUrl.TrimEnd('/');
                string selectUrl = $"{baseUrl}/rest/v1/agent_commands?agent_id=eq.{_agentId}&status=eq.pending&select=*";

                var response = await _httpClient.GetAsync(selectUrl);
                if (!response.IsSuccessStatusCode) return;

                var content = await response.Content.ReadAsStringAsync();
                var commands = JsonSerializer.Deserialize<List<AgentCommand>>(content);

                if (commands == null || commands.Count == 0) return;

                LogMessage($"Found {commands.Count} pending command(s)");

                foreach (var command in commands)
                {
                    await ExecuteCommandAsync(command);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error polling commands: {ex.Message}");
            }
        }

        private async Task ExecuteCommandAsync(AgentCommand command)
        {
            LogMessage($"Executing command: {command.CommandType}");
            
            // Update status to running
            await UpdateCommandStatusAsync(command.Id, "running", null, null);

            string? output = null;
            string? error = null;
            object? result = null;

            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddScript(command.Script);

                    var results = await Task.Run(() => ps.Invoke());

                    var outputBuilder = new StringBuilder();
                    foreach (var item in results)
                    {
                        outputBuilder.AppendLine(item?.ToString() ?? "");
                    }
                    output = outputBuilder.ToString();

                    if (ps.HadErrors)
                    {
                        var errorBuilder = new StringBuilder();
                        foreach (var err in ps.Streams.Error)
                        {
                            errorBuilder.AppendLine(err.ToString());
                        }
                        error = errorBuilder.ToString();
                    }

                    // Try to serialize results
                    if (results.Count > 0)
                    {
                        try
                        {
                            result = results.Count == 1 
                                ? results[0]?.BaseObject 
                                : results.Select(r => r?.BaseObject).ToList();
                        }
                        catch
                        {
                            result = output;
                        }
                    }
                }

                await UpdateCommandStatusAsync(command.Id, "completed", output, error, result);
                LogMessage($"Command completed: {command.CommandType}");
            }
            catch (Exception ex)
            {
                error = ex.Message;
                await UpdateCommandStatusAsync(command.Id, "failed", output, error);
                LogMessage($"Command failed: {ex.Message}");
            }
        }

        private async Task UpdateCommandStatusAsync(string commandId, string status, string? output, string? error, object? result = null)
        {
            try
            {
                string baseUrl = _config!.SupabaseUrl.TrimEnd('/');
                string updateUrl = $"{baseUrl}/rest/v1/agent_commands?id=eq.{commandId}";

                var updateData = new Dictionary<string, object?>
                {
                    { "status", status },
                    { "output", output },
                    { "error", error }
                };

                if (status == "running")
                {
                    updateData["started_at"] = DateTime.UtcNow.ToString("o");
                }
                else if (status == "completed" || status == "failed")
                {
                    updateData["completed_at"] = DateTime.UtcNow.ToString("o");
                    if (result != null)
                    {
                        updateData["result"] = result;
                    }
                }

                var content = new StringContent(
                    JsonSerializer.Serialize(updateData),
                    Encoding.UTF8,
                    "application/json"
                );

                var request = new HttpRequestMessage(HttpMethod.Patch, updateUrl)
                {
                    Content = content
                };

                await _httpClient.SendAsync(request);
            }
            catch (Exception ex)
            {
                LogMessage($"Error updating command status: {ex.Message}");
            }
        }

        private void LogMessage(string message)
        {
            string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            
            // Log to file
            try
            {
                string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "NEXUS Agent");
                Directory.CreateDirectory(logDir);
                string logFile = Path.Combine(logDir, "agent.log");
                
                // Keep log file from growing too large
                if (File.Exists(logFile) && new FileInfo(logFile).Length > 10 * 1024 * 1024) // 10MB
                {
                    File.Delete(logFile);
                }
                
                File.AppendAllText(logFile, logMessage + Environment.NewLine);
            }
            catch
            {
                // Ignore logging errors
            }

            // Also log to console if running interactively
            if (Environment.UserInteractive)
            {
                Console.WriteLine(logMessage);
            }
        }

        private class AgentRecord
        {
            public string Id { get; set; } = "";
        }
    }

    public class AgentConfig
    {
        public string SupabaseUrl { get; set; } = "";
        public string SupabaseKey { get; set; } = "";
        public string DeploymentId { get; set; } = "";
        public string SecretKey { get; set; } = "";

        // JSON property name mapping
        [System.Text.Json.Serialization.JsonPropertyName("supabase_url")]
        public string SupabaseUrlAlt { set { SupabaseUrl = value; } }
        
        [System.Text.Json.Serialization.JsonPropertyName("supabase_key")]
        public string SupabaseKeyAlt { set { SupabaseKey = value; } }
        
        [System.Text.Json.Serialization.JsonPropertyName("deployment_id")]
        public string DeploymentIdAlt { set { DeploymentId = value; } }
        
        [System.Text.Json.Serialization.JsonPropertyName("secret_key")]
        public string SecretKeyAlt { set { SecretKey = value; } }
    }

    public class AgentCommand
    {
        [System.Text.Json.Serialization.JsonPropertyName("id")]
        public string Id { get; set; } = "";
        
        [System.Text.Json.Serialization.JsonPropertyName("command_type")]
        public string CommandType { get; set; } = "";
        
        [System.Text.Json.Serialization.JsonPropertyName("script")]
        public string Script { get; set; } = "";
        
        [System.Text.Json.Serialization.JsonPropertyName("parameters")]
        public JsonElement? Parameters { get; set; }
    }
}
