using System;
using System.Collections.Generic;

using System.IO;
using System.Net;
using Opc.Da;
using System.Diagnostics;
using System.Threading;

namespace OPCtoFirebirdService
{
    class OPCClient
    {
        //тут все понятно
        private Server _opcServer;
        private string _hostName;
        private string _opcServerVendor;
        private int _updateRate;
        private OPCObject _opcObject;
        private List<Item> _scadaItems;
        private Subscription _scadaSubscription;
        private SubscriptionState _subscriptionState;

        //поток,в котором будет выполняться опрос контроллера
        Thread _readThread;
        //словари хранят значение тега и фиксируют изменение по каждому
        private Dictionary<string, double> _tagsValues;
        private Dictionary<string, bool> _tagsChanged;
        //бит остановки опроса, сервер disconect, dispose
        private bool _stop = false;
        //опрашиваемый топик
        private string _topic;
        //обеспечивает единоразовую синхронизацию времени
        private bool _timeSynced = false;
        //сюда будет писать метод Read() (не исп-ся)
        Opc.IRequest _request;

        public OPCClient(string hostName, string opcServerVendor, int updateRate, OPCObject obj)
        {
            //присваиваем значения переменных
            _hostName = hostName;
            _opcServerVendor = opcServerVendor;
            _updateRate = updateRate;
            _opcObject = obj;

            //инициализируем объект Server
            string scadaUrl = string.Format("opcda://{0}/{1}", hostName, opcServerVendor);
            _opcServer = new Opc.Da.Server(new OpcCom.Factory(), new Opc.URL(scadaUrl));

            //заполняем словари тегов значениями по умолчанию
            _tagsValues = new Dictionary<string, double>();
            _tagsChanged = new Dictionary<string, bool>();
            foreach (string tag in _opcObject.Tags)
            {
                try
                {
                    _tagsValues.Add(tag, 00.00f);
                    _tagsChanged.Add(tag, false);
                }
                catch (ArgumentException)
                {
                    throw new ArgumentException("Адреса тегов, хранящих время последней сводки (напр.: DTHourAddress и т.д.) не должны дублироваться в адресах раздела Tags");
                }
            }

            //создаем список тегов
            _scadaItems = new List<Opc.Da.Item>();
            foreach (string tag in _opcObject.Tags)
            {
                Opc.Da.Item item = new Opc.Da.Item()
                {
                    ItemName = tag,
                    Active = true,
                    ActiveSpecified = true
                };
                _scadaItems.Add(item);
            }

            //вытаскиваем топик, который опрашивает данный клиент
            string tag0 = _opcObject.Tags[0];
            _topic = tag0.Substring(tag0.IndexOf("[") + 1, tag0.IndexOf("]") - tag0.IndexOf("[") - 1);
        }

        #region Глобальные переменные
        /// <summary>
        /// Возвращает имя или ip-адрес машины, на которой установлен OPC-сервер
        /// </summary>
        public string HostName
        {
            get { return _hostName; }
        }
        /// <summary>
        /// Возвращает имя OPC-сервера
        /// </summary>
        public string OpcServerVendor
        {
            get { return _opcServerVendor; }
        }
        /// <summary>
        /// Возвращает период опроса в мс
        /// </summary>
        public int UpdateRate
        {
            get { return _updateRate; }
        }
        /// <summary>
        /// Возвращает массив с адресами OPC-тегов (напр.:[TOPIC]F14:0)
        /// </summary>
        public List<string> Tags
        {
            get { return _opcObject.Tags; }
        }
        /// <summary>
        /// Задает или возвращает значение бита, прекращающего опрос
        /// </summary>
        public bool StopPolling
        {
            get { return _stop; }
            set { _stop = value; }
        }
        /// <summary>
        /// Возвращает опрашиваемый топик
        /// </summary>
        public string Topic
        {
            get { return _topic; }
        }
        /// <summary>
        /// Возвращает OPC-сервер объекта
        /// </summary>
        public Server OPCServer
        {
            get { return _opcServer; }
        }
        #endregion

        public void Connect()
        {
            try
            {
                //подключаемся к серверу
                _opcServer.Connect(new Opc.ConnectData(new NetworkCredential()));
            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка подключения к серверу " + string.Format("opcda://{0}/{1}", _hostName, _opcServerVendor) + ": " + ex.Message);
            }

            if (!_opcServer.IsConnected)
            {
                throw new Exception("Подключение к серверу " + string.Format("opcda://{0}/{1} не удалось", _hostName, _opcServerVendor));
            }
            else
            {
                //создаем подписку
                _subscriptionState = new Opc.Da.SubscriptionState()
                {
                    Active = true,
                    UpdateRate = _updateRate,
                    Deadband = 0,
                    Name = "OPCtoFirebirdSubscription on " + _topic
                };
                _scadaSubscription = (Opc.Da.Subscription)_opcServer.CreateSubscription(_subscriptionState);

                //добавляем теги в подписку
                Opc.Da.ItemResult[] result = _scadaSubscription.AddItems(_scadaItems.ToArray());
                for (int i = 0; i < result.Length; i++)
                {
                    _scadaItems[i].ServerHandle = result[i].ServerHandle;
                }

                _scadaSubscription.State.Active = true;

                //запрашиваем данные пока не подана команда _stop
                _readThread = new Thread(new ThreadStart(ReadData));
                _readThread.Start();
            }
        }

        private void ReadData()
        {
            while (!_stop)
            {
                if (_opcObject.NeedTimeSync)
                {
                    //подпрограмма синхронизации времени
                    SyncTime(2, 30);
                }

                try { _scadaSubscription.Read(_scadaItems.ToArray(), 123, group_DataReadDone, out _request); }
                catch (Exception ex)
                {
                    _stop = true;
                    Stop();
                    Program.MyService.AddLog(ex.Message, EventLogEntryType.Error);
                    Program.MyService.Stop();
                }

                Thread.Sleep(_updateRate);
            }
        }

        private void group_DataReadDone(object clientHandle, ItemValueResult[] results)
        {
            try
            {
                //обходим каждое измененное значение
                foreach (ItemValueResult item in results)
                {
                    //записываем его в наш массив
                    _tagsValues[item.ItemName] = Math.Round(Convert.ToDouble(item.Value), 3);
                    //и фиксируем изменение
                    _tagsChanged[item.ItemName] = true;
                }

                //проверяем все ли теги зафиксированы
                bool allChanged = true;
                foreach (string TKey in _tagsChanged.Keys)
                {
                    if (_tagsChanged[TKey] == false)
                        allChanged = false;
                }

                //если да
                if (allChanged)
                {
                    try
                    {
                        //записываем значения в базу
                        FBWriter fbWriter = new FBWriter();
                        fbWriter.Write(_tagsValues, _opcObject);
                    }
                    catch (Exception ex)
                    {
                        Program.MyService.AddLog("Исключение при записи в БД Firebird: " + ex.Message, EventLogEntryType.Error);
                    }
                    //обнуляем изменения
                    foreach (string tag in _opcObject.Tags)
                    {
                        _tagsValues[tag] = 00.00f;
                        _tagsChanged[tag] = false;
                    }
                }
            }
            catch (Exception ex)
            {
                //тут пусто
                //чтобы опрос продолжался, пока контроллер опять не появится
            }
        }

        /// <summary>
        /// Синхронизирует время контроллера со временем машины в указанный период суток
        /// </summary>
        private void SyncTime(int hour, int minute)
        {
            if (DateTime.Now.Hour == hour && DateTime.Now.Minute == minute)
            {
                if (!_timeSynced)
                {
                    try
                    {
                        TimeSync sync = new TimeSync(this);
                        sync.SyncNow();
                        Program.MyService.AddLog("Время успешно синхронизированно на " + _topic, EventLogEntryType.SuccessAudit);
                    }
                    catch (Exception ex)
                    {
                        Program.MyService.AddLog(ex.Message, EventLogEntryType.Error);
                    }
                    finally
                    {
                        //не будем долбить больше 1го раза
                        //даже если не получилось
                        _timeSynced = true;
                    }
                }
            }
            else
            {
                _timeSynced = false;
            }
        }

        /// <summary>
        /// Прекращает опрос OPC-сервера
        /// </summary>
        public void Stop()
        {
            if (_opcServer != null)
            {
                try
                {
                    if (_opcServer.IsConnected)
                    {
                        foreach (Subscription sub in _opcServer.Subscriptions)
                        {
                            if (sub != null)
                                _opcServer.CancelSubscription(sub);
                        }
                        _opcServer.Disconnect();
                    }

                    if (_readThread.ThreadState == System.Threading.ThreadState.Running)
                        _readThread.Abort();

                    _opcServer.Dispose();
                }
                catch (Exception ex)
                {
                    Program.MyService.AddLog(ex.Message, EventLogEntryType.Error);
                }
            }
        }
    }
}
