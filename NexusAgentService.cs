using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NexusAgent
{
    public class NexusAgentService : ServiceBase
    {
        private CancellationTokenSource _cancellationTokenSource;
        private Task _mainTask;
        private HttpClient _httpClient;
        private Timer _heartbeatTimer;
        private Timer _commandTimer;
        
        // Credentials read from embedded data
        private string _supabaseUrl;
        private string _supabaseKey;
        private string _deploymentId;
        private string _secretKey;
        
        // Installation ID for this agent instance
        private string _installationId;
        
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
            _cancellationTokenSource = new CancellationTokenSource();
            _mainTask = Task.Run(() => RunAgent(_cancellationTokenSource.Token));
        }

        protected override void OnStop()
        {
            _cancellationTokenSource?.Cancel();
            _heartbeatTimer?.Dispose();
            _commandTimer?.Dispose();
            _httpClient?.Dispose();
            
            try
            {
                _mainTask?.Wait(TimeSpan.FromSeconds(10));
            }
            catch { }
        }

        private async Task RunAgent(CancellationToken cancellationToken)
        {
            try
            {
                // Read embedded credentials from the executable
                if (!ReadEmbeddedCredentials())
                {
                    LogError("Failed to read embedded credentials");
                    return;
                }

                LogInfo($"NEXUS Agent starting for deployment: {_deploymentId}");
                
                // Initialize HTTP client
                _httpClient = new HttpClient();
                _httpClient.DefaultRequestHeaders.Add("apikey", _supabaseKey);
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_supabaseKey}");
                
                // Register or update this agent installation
                await RegisterAgent();
                
                // Start heartbeat timer (every 60 seconds)
                _heartbeatTimer = new Timer(
                    async _ => await SendHeartbeat(),
                    null,
                    TimeSpan.Zero,
                    TimeSpan.FromSeconds(60)
                );
                
                // Start command polling timer (every 10 seconds)
                _commandTimer = new Timer(
                    async _ => await PollForCommands(),
                    null,
                    TimeSpan.FromSeconds(5),
                    TimeSpan.FromSeconds(10)
                );
                
                LogInfo("NEXUS Agent started successfully");
                
                // Keep running until cancelled
                while (!cancellationToken.IsCancellationRequested)
                {
                    await Task.Delay(1000, cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
                LogInfo("NEXUS Agent stopping...");
            }
            catch (Exception ex)
            {
                LogError($"Agent error: {ex.Message}");
            }
        }

        private bool ReadEmbeddedCredentials()
        {
            try
            {
                // Get the path to the currently running executable
                string exePath = Assembly.GetExecutingAssembly().Location;
                
                LogInfo($"Reading credentials from: {exePath}");
                
                byte[] exeBytes = File.ReadAllBytes(exePath);
                
                if (exeBytes.Length < CREDENTIALS_SIZE)
                {
                    LogError("Executable too small to contain credentials");
                    return false;
                }
                
                // Read the last CREDENTIALS_SIZE bytes
                byte[] credentialsBytes = new byte[CREDENTIALS_SIZE];
                Array.Copy(exeBytes, exeBytes.Length - CREDENTIALS_SIZE, credentialsBytes, 0, CREDENTIALS_SIZE);
                
                // Convert to string and trim null bytes
                string credentialsJson = Encoding.UTF8.GetString(credentialsBytes).TrimEnd('\0', ' ');
                
                LogInfo($"Raw credentials length: {credentialsJson.Length}");
                
                // Check if it's still the placeholder
                if (credentialsJson.Contains("<<<NEXUS_CREDENTIALS_PLACEHOLDER>>>"))
                {
                    LogError("Agent executable has not been configured with credentials");
                    return false;
                }
                
                // Parse JSON
                var credentials = JObject.Parse(credentialsJson);
                
                _supabaseUrl = credentials["supabase_url"]?.ToString();
                _supabaseKey = credentials["supabase_key"]?.ToString();
                _deploymentId = credentials["deployment_id"]?.ToString();
                _secretKey = credentials["secret_key"]?.ToString();
                
                if (string.IsNullOrEmpty(_supabaseUrl) || 
                    string.IsNullOrEmpty(_supabaseKey) || 
                    string.IsNullOrEmpty(_deploymentId) || 
                    string.IsNullOrEmpty(_secretKey))
                {
                    LogError("Missing required credential fields");
                    return false;
                }
                
                LogInfo($"Credentials loaded successfully for deployment: {_deploymentId}");
                return true;
            }
            catch (Exception ex)
            {
                LogError($"Failed to read credentials: {ex.Message}");
                return false;
            }
        }

        private async Task RegisterAgent()
        {
            try
            {
                string hostname = Environment.MachineName;
                string url = $"{_supabaseUrl}/rest/v1/agent_installations?secret_key=eq.{_secretKey}";
                
                var response = await _httpClient.GetAsync(url);
                var content = await response.Content.ReadAsStringAsync();
                var installations = JArray.Parse(content);
                
                if (installations.Count > 0)
                {
                    // Update existing installation
                    _installationId = installations[0]["id"]?.ToString();
                    
                    var updateData = new
                    {
                        hostname = hostname,
                        last_heartbeat = DateTime.UtcNow.ToString("o"),
                        status = "connected",
                        agent_version = "1.0.0"
                    };
                    
                    string updateUrl = $"{_supabaseUrl}/rest/v1/agent_installations?id=eq.{_installationId}";
                    var updateContent = new StringContent(
                        JsonConvert.SerializeObject(updateData),
                        Encoding.UTF8,
                        "application/json"
                    );
                    
                    // Use SendAsync with PATCH method (PatchAsync not available in .NET Framework 4.8)
                    var request = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl);
                    request.Content = updateContent;
                    await _httpClient.SendAsync(request);
                    
                    LogInfo($"Updated agent registration: {_installationId}");
                }
                else
                {
                    LogError("No matching agent installation found for secret key");
                }
            }
            catch (Exception ex)
            {
                LogError($"Failed to register agent: {ex.Message}");
            }
        }

        private async Task SendHeartbeat()
        {
            try
            {
                if (string.IsNullOrEmpty(_installationId)) return;
                
                var updateData = new
                {
                    last_heartbeat = DateTime.UtcNow.ToString("o"),
                    status = "connected"
                };
                
                string url = $"{_supabaseUrl}/rest/v1/agent_installations?id=eq.{_installationId}";
                var content = new StringContent(
                    JsonConvert.SerializeObject(updateData),
                    Encoding.UTF8,
                    "application/json"
                );
                
                // Use SendAsync with PATCH method (PatchAsync not available in .NET Framework 4.8)
                var request = new HttpRequestMessage(new HttpMethod("PATCH"), url);
                request.Content = content;
                await _httpClient.SendAsync(request);
                
                LogInfo("Heartbeat sent");
            }
            catch (Exception ex)
            {
                LogError($"Heartbeat failed: {ex.Message}");
            }
        }

        private async Task PollForCommands()
        {
            try
            {
                if (string.IsNullOrEmpty(_installationId)) return;
                
                string url = $"{_supabaseUrl}/rest/v1/agent_commands?installation_id=eq.{_installationId}&status=eq.pending&order=created_at.asc&limit=1";
                
                var response = await _httpClient.GetAsync(url);
                var content = await response.Content.ReadAsStringAsync();
                var commands = JArray.Parse(content);
                
                if (commands.Count > 0)
                {
                    var command = commands[0];
                    string commandId = command["id"]?.ToString();
                    string commandType = command["command_type"]?.ToString();
                    string commandData = command["command_data"]?.ToString();
                    
                    LogInfo($"Executing command: {commandType}");
                    
                    // Mark command as running
                    await UpdateCommandStatus(commandId, "running", null);
                    
                    // Execute the command
                    string result = await ExecuteCommand(commandType, commandData);
                    
                    // Mark command as completed
                    await UpdateCommandStatus(commandId, "completed", result);
                    
                    LogInfo($"Command completed: {commandType}");
                }
            }
            catch (Exception ex)
            {
                LogError($"Command polling failed: {ex.Message}");
            }
        }

        private async Task<string> ExecuteCommand(string commandType, string commandData)
        {
            try
            {
                switch (commandType?.ToLower())
                {
                    case "powershell":
                        return await ExecutePowerShell(commandData);
                    
                    case "ping":
                        return $"Pong from {Environment.MachineName} at {DateTime.UtcNow:o}";
                    
                    case "info":
                        return JsonConvert.SerializeObject(new
                        {
                            hostname = Environment.MachineName,
                            os = Environment.OSVersion.ToString(),
                            dotnet = Environment.Version.ToString(),
                            processors = Environment.ProcessorCount,
                            uptime = Environment.TickCount / 1000
                        });
                    
                    default:
                        return $"Unknown command type: {commandType}";
                }
            }
            catch (Exception ex)
            {
                return $"Error executing command: {ex.Message}";
            }
        }

        private async Task<string> ExecutePowerShell(string script)
        {
            try
            {
                var processInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{script}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                
                using (var process = System.Diagnostics.Process.Start(processInfo))
                {
                    string output = await process.StandardOutput.ReadToEndAsync();
                    string error = await process.StandardError.ReadToEndAsync();
                    
                    process.WaitForExit();
                    
                    if (!string.IsNullOrEmpty(error))
                    {
                        return $"Error: {error}\nOutput: {output}";
                    }
                    
                    return output;
                }
            }
            catch (Exception ex)
            {
                return $"PowerShell execution failed: {ex.Message}";
            }
        }

        private async Task UpdateCommandStatus(string commandId, string status, string result)
        {
            try
            {
                var updateData = new
                {
                    status = status,
                    result = result,
                    executed_at = status == "completed" ? DateTime.UtcNow.ToString("o") : null
                };
                
                string url = $"{_supabaseUrl}/rest/v1/agent_commands?id=eq.{commandId}";
                var content = new StringContent(
                    JsonConvert.SerializeObject(updateData),
                    Encoding.UTF8,
                    "application/json"
                );
                
                // Use SendAsync with PATCH method (PatchAsync not available in .NET Framework 4.8)
                var request = new HttpRequestMessage(new HttpMethod("PATCH"), url);
                request.Content = content;
                await _httpClient.SendAsync(request);
            }
            catch (Exception ex)
            {
                LogError($"Failed to update command status: {ex.Message}");
            }
        }

        private void LogInfo(string message)
        {
            try
            {
                string logPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "NexusAgent",
                    "agent.log"
                );
                
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [INFO] {message}{Environment.NewLine}";
                File.AppendAllText(logPath, logMessage);
            }
            catch { }
        }

        private void LogError(string message)
        {
            try
            {
                string logPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "NexusAgent",
                    "agent.log"
                );
                
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [ERROR] {message}{Environment.NewLine}";
                File.AppendAllText(logPath, logMessage);
            }
            catch { }
        }
    }
}
