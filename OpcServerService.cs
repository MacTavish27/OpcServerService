using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
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

        private readonly List<Subscription> _subscriptions = new List<Subscription>();
        private Subscription subscription;
        public static OpcServerService Instance { get; private set; }


        private readonly string selectedHost;
        private readonly string opcServerName;
        private readonly string baseAddress;

        private DateTime _lastUpdateTime;
        private Timer latencyTimer;
        private const int _expectedInterval = 100;
        private int _updateCount = 0;
        private long totalLatencyTicks = 0;
        private long latencySamples = 0;
        private DateTime _fpsTime = DateTime.UtcNow;
        private readonly object _subLock = new object();
        public OpcServerService()
        {
            InitializeComponent();
            Instance = this;
            this.ServiceName = "OPC Server Service";

            selectedHost = ConfigurationManager.AppSettings["OpcHost"];
            opcServerName = ConfigurationManager.AppSettings["OpcServerName"];
            baseAddress = ConfigurationManager.AppSettings["BaseAddress"];
        }

        protected override void OnStart(string[] args)
        {
            StartOpcThread();
            StartLatencyMonitor();

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
            HttpLogger.Log("Service stopped");
            var wait = new ManualResetEvent(false);

            _opcQueue.Add(() =>
            {
                RemoveAllSubscriptions();
                DisconnectServer(opcServer);
                wait.Set();
            });

            wait.WaitOne();
            _opcQueue.CompleteAdding();
            _opcThread.Join();
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

        public void EnsureSubscription()
        {
            lock (_subLock)
            {
                if (subscription != null)
                    return;

                var state = new SubscriptionState
                {
                    Name = "MainSubscription",
                    Active = true,
                    UpdateRate = 100
                };

                subscription = (Subscription)opcServer.CreateSubscription(state);

                subscription.DataChanged += OnDataChanged;

                HttpLogger.Log("Subscription created");
            }
        }

        public void SubscribeTags(string[] tagIds)
        {

            var wait = new ManualResetEvent(false);

            _opcQueue.Add(() =>
            {
                try
                {
                    EnsureSubscription();

                    var items = tagIds.Select(tag => new Item
                    {
                        ItemName = tag
                    }).ToArray();


                    subscription.AddItems(items);

                    _subscriptions.Add(subscription);

                    HttpLogger.Log($"Subscribed {items.Length} tags");

                }
                finally
                {
                    wait.Set();
                }

            });

            wait.WaitOne();

        }

        private void RemoveAllSubscriptions()
        {
            foreach (var sub in _subscriptions)
            {
                try
                {
                    sub.DataChanged -= OnDataChanged;

                    if (sub.Items != null)
                    {
                        sub.RemoveItems(sub.Items);

                        sub.Dispose();

                        HttpLogger.Log("All subscriptions have been removed successfully");
                    }

                }
                catch (Exception ex)
                {
                    HttpLogger.Log("[ERROR] Subscription dispose failed: " + ex.Message);
                }
            }

            _subscriptions.Clear();
        }

        private void OnDataChanged(object subscriptionHandle, object requestHandle, ItemValueResult[] values)
        {
            Interlocked.Add(ref _updateCount, values.Length);

            var now = DateTime.UtcNow;

            var receiveTime = DateTime.Now;

            var actualInterval = (now - _lastUpdateTime).TotalMilliseconds;

            var jitter = actualInterval - _expectedInterval;

            if ((now - _fpsTime).TotalSeconds >= 1)
            {
                int fps = _updateCount;
                _updateCount = 0;
                _fpsTime = now;

                LogFPS(fps);
            }

            LogPerformance(actualInterval, jitter);


            _lastUpdateTime = now;


            foreach (var item in values)
            {
                var latency = (receiveTime - item.Timestamp);
                Interlocked.Add(ref totalLatencyTicks, latency.Ticks);
                Interlocked.Increment(ref latencySamples);
            }
        }

        private void StartLatencyMonitor()
        {
            latencyTimer = new Timer(_ =>
            {
                var samples = latencySamples;

                if (samples == 0)
                    return;

                var avgTicks = totalLatencyTicks / samples;

                var avgLatency = TimeSpan.FromTicks(avgTicks);

                LatencyLog(avgLatency, samples);

                totalLatencyTicks = 0;
                latencySamples = 0;

            }, null, 1000, 1000);
        }

    }
}
