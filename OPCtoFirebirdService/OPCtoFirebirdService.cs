using System;
using System.Diagnostics;
using System.ServiceProcess;

using System.Collections.Generic;



namespace OPCtoFirebirdService
{
    public partial class OPCtoFirebirdService : ServiceBase
    {
        private List<OPCClient> _clients;
        public OPCtoFirebirdService()
        {
            InitializeComponent();

            _clients = new List<OPCClient>();
        }

        private OPCClient client;
        protected override void OnStart(string[] args)
        {
            try
            {
                foreach (OPCObject obj in Settings.Objects)
                {
                    OPCClient client = new OPCClient(Settings.HostName, Settings.OpcServerVendor, Settings.UpdateRate, obj);
                    client.Connect();

                    _clients.Add(client);
                }
            }
            catch (Exception ex)
            {
                AddLog(ex.Message, EventLogEntryType.Error);
                this.Stop();
            }
        }

        protected override void OnStop()
        {
            foreach (OPCClient client in _clients)
            {
                if (client != null) { client.StopPolling = true; client.Stop(); }
            }
        }

        /// <summary>
        /// Добавляет запись в Журнал Windows
        /// </summary>
        /// <param name="log">Текст записи</param>
        public void AddLog(string text, EventLogEntryType eventType)
        {
            try
            {
                if (!EventLog.SourceExists("OPCtoFirebirdService"))
                {
                    EventLog.CreateEventSource("OPCtoFirebirdService", "OPCtoFirebirdService");
                }
                eventLog.Source = "OPCtoFirebirdService";
                eventLog.WriteEntry(text, eventType);
            }
            catch { this.Stop(); }
        }
    }
}
