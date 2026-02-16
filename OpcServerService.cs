using System;
using System.Collections.Concurrent;
using System.Linq;
using System.ServiceProcess;
using System.Threading;
using System.Web.Http;
using Microsoft.Owin.Hosting;
using Opc.Da;
using Owin;
using static opc_bridge.Services;

namespace opc_bridge
{
    public partial class OpcServerService : ServiceBase
    {
        private IDisposable _webApp;
        private Thread _opcThread;
        private readonly BlockingCollection<Action> _opcQueue = new BlockingCollection<Action>();
        private Opc.Da.Server opcServer;
        private OpcCom.Factory factory;
        public static OpcServerService Instance { get; private set; }


        private const string selectedHost = "localhost";
        private const string opcServerName = "Matrikon.OPC.Simulation.1";
        public OpcServerService()
        {
            InitializeComponent();
            Instance = this;
            this.ServiceName = "OPC Server Service";
        }

        protected override void OnStart(string[] args)
        {
            const string baseAddress = "http://*:8080";
            StartOpcThread();

            try
            {
                _webApp = WebApp.Start(baseAddress, appBuilder =>
                {
                    var config = new HttpConfiguration();
                    OpcWebApiConfig.Register(config);

                    appBuilder.UseWebApi(config);
                });
                HttpLogger.Log("Web API started at " + baseAddress);
            }

            catch (Exception ex)
            {
                HttpLogger.Log("[ERROR] Error in starting web server: " + ex.Message);
            }

        }

        protected override void OnStop()
        {
            HttpLogger.Log("Service is stopped");
            _opcQueue.CompleteAdding();
            _opcThread.Join();
            DisconnectServer(opcServer);
            _webApp?.Dispose();
        }

        public ItemValueResult ReadTag(string tagId)
        {
            ItemValueResult result = null;
            var waitHandle = new ManualResetEvent(false);
            Exception exception = null;
            _opcQueue.Add(() =>
            {
                try
                {
                    var item = new Item { ItemName = tagId };
                    var values = opcServer.Read(new[] { item });
                    result = values.FirstOrDefault();
                }
                catch (Exception ex)
                {
                    exception = ex;
                    HttpLogger.Log("[ERROR] Reading a value failed: " + ex.Message);
                }
                finally
                {
                    waitHandle.Set();
                }
            });

            waitHandle.WaitOne();

            if (exception != null) throw exception;
            return result;
        }


        public void WriteTag(string tagId, object value)
        {
            Exception exception = null;
            var waitHandle = new ManualResetEvent(false);

            _opcQueue.Add(() =>
            {
                try
                {
                    var item = new ItemValue { ItemName = tagId, Value = value };
                    opcServer.Write(new[] { item });
                }
                catch (Exception ex)
                {
                    exception = ex;
                    HttpLogger.Log("[ERROR] Writing a value failed: " + ex.Message);
                }
                finally
                {
                    waitHandle.Set();
                }
            });

            waitHandle.WaitOne();
            if (exception != null) throw exception;
        }


        private void StartOpcThread()
        {
            _opcThread = new Thread(() =>
            {
                factory = new OpcCom.Factory();
                opcServer = new Server(factory, null);

                try
                {
                    if (!opcServer.IsConnected)
                        ConnectToServer(opcServer, selectedHost, opcServerName);
                }
                catch (Exception ex)
                {
                    HttpLogger.Log("[ERROR] OPC connection error: " + ex.Message);
                }

                foreach (var action in _opcQueue.GetConsumingEnumerable())
                {
                    try
                    {
                        action();
                    }
                    catch (Exception ex)
                    {
                        HttpLogger.Log("[ERROR] OPC thread action error: " + ex.Message);
                    }
                }
            });

            _opcThread.SetApartmentState(ApartmentState.STA);
            _opcThread.IsBackground = true;
            _opcThread.Start();
        }

    }
}

