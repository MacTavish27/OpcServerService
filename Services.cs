using System;
using System.IO;
using Opc;

namespace opc_bridge
{
    internal class Services
    {
        public static class HttpLogger
        {
            public static event Action<string> LogReceived;

            private static readonly string LogFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs.txt");

            public static void Log(string message)
            {
                string timestamped = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                LogReceived?.Invoke(timestamped);

                try
                {
                    File.AppendAllText(LogFilePath, timestamped + Environment.NewLine);
                }
                catch (Exception ex)
                {
                    HttpLogger.Log($"[Logger Error] Failed to write log: {ex.Message}");
                }
            }
        }
        public static void ConnectToServer(Opc.Da.Server server, string selectedHost, string opcServerName)
        {

            try
            {
                server.Connect(new URL($"opcda://{selectedHost}/{opcServerName}"), new ConnectData(new System.Net.NetworkCredential()));
                HttpLogger.Log($"Connected to OPC Server: {opcServerName}");
            }

            catch (Exception ex)
            {
                HttpLogger.Log("[ERROR] Failed to connect to OPC Server: " + ex.Message);
            }

        }
        public static void DisconnectServer(Opc.Da.Server server)
        {
            try
            {
                if (server != null && server.IsConnected)
                    server.Disconnect();

                HttpLogger.Log($"{server.Name} is successfully disconnected");
            }

            catch (Exception ex)
            {
                HttpLogger.Log("[ERROR] Error in disconnecting the server: " + ex.Message);
            }
        }

    }
}
