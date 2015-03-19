using System;
using System.Collections.Generic;

using Opc.Da;

namespace OPCtoFirebirdService
{
    class TimeSync
    {
        private OPCClient _opcClient;
        private Subscription _writeGroup;

        public TimeSync(OPCClient opcClient)
        {
            _opcClient = opcClient;
        }

        public void SyncNow()
        {
            try
            {
                //создаем группу
                SubscriptionState groupState = new SubscriptionState()
                {
                    Name = "TymeSync with " + _opcClient.Topic,
                    Active = false
                };
                _writeGroup = (Subscription)_opcClient.OPCServer.CreateSubscription(groupState);

                //добавляем в группу адреса минут и секунд контроллера
                Item[] items = new Item[2];
                items[0] = new Item();
                items[0].ItemName = "[" + _opcClient.Topic + "]" + "S2:41";
                items[1] = new Item();
                items[1].ItemName = "[" + _opcClient.Topic + "]" + "S2:42";
                items = _writeGroup.AddItems(items);

                //записываем по адресу минуты и секунды компьютера
                ItemValue[] writeValues = new ItemValue[2];
                writeValues[0] = new ItemValue();
                writeValues[1] = new ItemValue();

                writeValues[0].ServerHandle = _writeGroup.Items[0].ServerHandle;
                writeValues[0].Value = DateTime.Now.Minute;
                writeValues[1].ServerHandle = _writeGroup.Items[1].ServerHandle;
                writeValues[1].Value = DateTime.Now.Second;

                Opc.IRequest req;
                _writeGroup.Write(writeValues, 321, new WriteCompleteEventHandler(WriteCompleteCallback), out req);

                //// тут еще один неизведанный способ считать данные
                //group.Read(group.Items, 123, new Opc.Da.ReadCompleteEventHandler(ReadCompleteCallback), out req);
                //Console.ReadLine();
            }
            catch (Exception ex)
            {
                _opcClient.OPCServer.CancelSubscription(_writeGroup);
                throw new Exception("Ошибка синхронизации времени с контроллером " + _opcClient.Topic + ": " + ex.Message);
            }
        }

        private void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {
            foreach (Opc.IdentifiedResult writeResult in results)
            {
                if (writeResult.ResultID != Opc.ResultID.S_OK)
                {
                    _opcClient.OPCServer.CancelSubscription(_writeGroup);
                    throw new Exception(writeResult.ItemName + " результат записи: " + writeResult.ResultID);
                }
            }
        }

        //private void ReadCompleteCallback(object clientHandle, Opc.Da.ItemValueResult[] results)
        //{
        //    foreach (Opc.Da.ItemValueResult readResult in results)
        //    {
        //        System.Windows.Forms.MessageBox.Show(String.Format("\t{0}\tval:{1}", readResult.ItemName, readResult.Value));
        //    }
        //}
    }
}
