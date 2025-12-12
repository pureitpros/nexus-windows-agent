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
        private Task _heartbeatTask;
        private readonly HttpClient _httpClient;
        
        private string _supabaseUrl;
        private string _supabaseKey;
        private string _secretKey;
        private string _deploymentId;
        
        private const int CREDENTIALS_SIZE = 500;

        public NexusAgentService()
        {
            ServiceName = "NexusAgent";
            CanStop = true;
            CanPauseAndContinue = false;
            AutoLog = true;
            
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(30);
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                WriteLog("NEXUS Agent starting...");
                
                if (!LoadCredentials())
                {
                    WriteLog("Failed to load credentials. Service cannot start.");
                    Stop();
                    return;
                }
                
                WriteLog($"Credentials loaded. Deployment: {_deploymentId}");
                WriteLog($"Supabase URL: {_supabaseUrl}");
                
                // Set up HTTP client headers
                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("apikey", _supabaseKey);
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_supabaseKey}");
                
                _cancellationTokenSource = new CancellationTokenSource();
                _heartbeatTask = Task.Run(() => HeartbeatLoop(_cancellationTokenSource.Token));
                
                WriteLog("NEXUS Agent started successfully.");
            }
            catch (Exception ex)
            {
                WriteLog($"Error starting service: {ex.Message}");
                throw;
            }
        }

        protected override void OnStop()
        {
            WriteLog("NEXUS Agent stopping...");
            
            try
            {
                _cancellationTokenSource?.Cancel();
                
                if (_heartbeatTask != null)
                {
                    _heartbeatTask.Wait(TimeSpan.FromSeconds(5));
                }
            }
            catch (Exception ex)
            {
                WriteLog($"Error during shutdown: {ex.Message}");
            }
            
            _httpClient?.Dispose();
            WriteLog("NEXUS Agent stopped.");
        }

        private bool LoadCredentials()
        {
            try
            {
                // Get the path to this executable
                string exePath = Assembly.GetExecutingAssembly().Location;
                WriteLog($"Loading credentials from: {exePath}");
                
                byte[] exeBytes = File.ReadAllBytes(exePath);
                WriteLog($"Executable size: {exeBytes.Length} bytes");
                
                if (exeBytes.Length < CREDENTIALS_SIZE)
                {
                    WriteLog("Executable too small to contain credentials");
                    return false;
                }
                
                // Read the last CREDENTIALS_SIZE bytes
                byte[] credentialsBytes = new byte[CREDENTIALS_SIZE];
                Array.Copy(exeBytes, exeBytes.Length - CREDENTIALS_SIZE, credentialsBytes, 0, CREDENTIALS_SIZE);
                
                // Convert to string and trim null bytes
                string credentialsJson = Encoding.UTF8.GetString(credentialsBytes).TrimEnd('\0').Trim();
                WriteLog($"Raw credentials string (first 100 chars): {credentialsJson.Substring(0, Math.Min(100, credentialsJson.Length))}");
                
                // Find the start of the JSON object
                int jsonStart = credentialsJson.IndexOf('{');
                if (jsonStart < 0)
                {
                    WriteLog("No JSON object found in credentials");
                    return false;
                }
                
                credentialsJson = credentialsJson.Substring(jsonStart);
                
                // Find the end of the JSON object
                int braceCount = 0;
                int jsonEnd = -1;
                for (int i = 0; i < credentialsJson.Length; i++)
                {
                    if (credentialsJson[i] == '{') braceCount++;
                    if (credentialsJson[i] == '}') braceCount--;
                    if (braceCount == 0)
                    {
                        jsonEnd = i + 1;
                        break;
                    }
                }
                
                if (jsonEnd > 0)
                {
                    credentialsJson = credentialsJson.Substring(0, jsonEnd);
                }
                
                WriteLog($"Parsed JSON: {credentialsJson}");
                
                var credentials = JObject.Parse(credentialsJson);
                
                _supabaseUrl = credentials["supabase_url"]?.ToString();
                _supabaseKey = credentials["supabase_key"]?.ToString();
                _secretKey = credentials["secret_key"]?.ToString();
                _deploymentId = credentials["deployment_id"]?.ToString();
                
                if (string.IsNullOrEmpty(_supabaseUrl) || string.IsNullOrEmpty(_supabaseKey) || 
                    string.IsNullOrEmpty(_secretKey) || string.IsNullOrEmpty(_deploymentId))
                {
                    WriteLog("One or more credentials are missing");
                    WriteLog($"URL: {!string.IsNullOrEmpty(_supabaseUrl)}, Key: {!string.IsNullOrEmpty(_supabaseKey)}, Secret: {!string.IsNullOrEmpty(_secretKey)}, Deployment: {!string.IsNullOrEmpty(_deploymentId)}");
                    return false;
                }
                
                return true;
            }
            catch (Exception ex)
            {
                WriteLog($"Error loading credentials: {ex.Message}");
                WriteLog($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        private async Task HeartbeatLoop(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    await SendHeartbeat();
                    await CheckForCommands();
                }
                catch (Exception ex)
                {
                    WriteLog($"Heartbeat error: {ex.Message}");
                }
                
                // Wait 30 seconds before next heartbeat
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

        private async Task SendHeartbeat()
        {
            try
            {
                string hostname = Environment.MachineName;
                string url = $"{_supabaseUrl}/rest/v1/agent_installations?deployment_id=eq.{_deploymentId}&secret_key=eq.{_secretKey}";
                
                var payload = new
                {
                    last_heartbeat = DateTime.UtcNow.ToString("o"),
                    hostname = hostname,
                    agent_version = "1.0.0",
                    status = "connected"
                };
                
                string jsonPayload = JsonConvert.SerializeObject(payload);
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                content.Headers.Add("Prefer", "return=minimal");
                
                var response = await _httpClient.PatchAsync(url, content);
                
                if (!response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    WriteLog($"Heartbeat failed: {response.StatusCode} - {responseBody}");
                }
                else
                {
                    WriteLog("Heartbeat sent successfully");
                }
            }
            catch (Exception ex)
            {
                WriteLog($"Error sending heartbeat: {ex.Message}");
            }
        }

        private async Task CheckForCommands()
        {
            try
            {
                string url = $"{_supabaseUrl}/rest/v1/agent_commands?deployment_id=eq.{_deploymentId}&status=eq.pending&order=created_at.asc&limit=1";
                
                var response = await _httpClient.GetAsync(url);
                
                if (response.IsSuccessStatusCode)
                {
                    string body = await response.Content.ReadAsStringAsync();
                    var commands = JArray.Parse(body);
                    
                    if (commands.Count > 0)
                    {
                        var command = commands[0];
                        string commandId = command["id"]?.ToString();
                        string commandType = command["command_type"]?.ToString();
                        string commandPayload = command["payload"]?.ToString();
                        
                        WriteLog($"Received command: {commandType}");
                        
                        await ExecuteCommand(commandId, commandType, commandPayload);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog($"Error checking commands: {ex.Message}");
            }
        }

        private async Task ExecuteCommand(string commandId, string commandType, string payload)
        {
            string result = "";
            string status = "completed";
            
            try
            {
                switch (commandType?.ToLower())
                {
                    case "ping":
                        result = "pong";
                        break;
                    case "powershell":
                        result = await ExecutePowerShell(payload);
                        break;
                    default:
                        result = $"Unknown command type: {commandType}";
                        status = "failed";
                        break;
                }
            }
            catch (Exception ex)
            {
                result = $"Error: {ex.Message}";
                status = "failed";
            }
            
            // Update command status
            await UpdateCommandStatus(commandId, status, result);
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
                    
                    process.WaitForExit(60000); // 60 second timeout
                    
                    if (!string.IsNullOrEmpty(error))
                    {
                        return $"Output:\n{output}\n\nErrors:\n{error}";
                    }
                    
                    return output;
                }
            }
            catch (Exception ex)
            {
                return $"PowerShell execution error: {ex.Message}";
            }
        }

        private async Task UpdateCommandStatus(string commandId, string status, string result)
        {
            try
            {
                string url = $"{_supabaseUrl}/rest/v1/agent_commands?id=eq.{commandId}";
                
                var payload = new
                {
                    status = status,
                    result = result,
                    completed_at = DateTime.UtcNow.ToString("o")
                };
                
                string jsonPayload = JsonConvert.SerializeObject(payload);
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                content.Headers.Add("Prefer", "return=minimal");
                
                await _httpClient.PatchAsync(url, content);
            }
            catch (Exception ex)
            {
                WriteLog($"Error updating command status: {ex.Message}");
            }
        }

        private void WriteLog(string message)
        {
            try
            {
                string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "NexusAgent");
                Directory.CreateDirectory(logDir);
                
                string logFile = Path.Combine(logDir, "nexus-agent.log");
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n";
                
                File.AppendAllText(logFile, logEntry);
            }
            catch
            {
                // Ignore logging errors
            }
        }
    }
}
