using System;
using System.Collections.Generic;

using FirebirdSql.Data.FirebirdClient;

namespace OPCtoFirebirdService
{
    class FBWriter
    {
        private FbConnection fbConnection;
        public FBWriter()
        {
            //создаем экземпляр объекта FbConnection
            FbConnectionStringBuilder fbConStr = new FbConnectionStringBuilder();
            fbConStr.Charset = "WIN1251";
            fbConStr.UserID = Settings.UserName;
            fbConStr.Password = Settings.Password;
            fbConStr.Database = Settings.DatabasePath;
            fbConStr.ServerType = FbServerType.Default;

            fbConnection = new FbConnection(fbConStr.ToString());
        }

        public void Write(Dictionary<string, double> tagsValuesDic, OPCObject obj)
        {
            //Открываем подключение
            if (fbConnection.State == System.Data.ConnectionState.Closed)
            {
                try
                {
                    fbConnection.Open();
                }
                catch (Exception ex)
                {
                    throw new Exception("Ошибка подключения к базе данных: " + ex.Message);
                }
            }

            FbCommand procCommand = new FbCommand(obj.StoredProcedureName, fbConnection);
            procCommand.CommandType = System.Data.CommandType.StoredProcedure;

            FbTransaction fbTransaction = fbConnection.BeginTransaction();
            procCommand.Transaction = fbTransaction;

            //Формируем массив значений равный количеству параметров процедуры,
            //  указанному в файле settings и заполняем его значениями считанных параметров
            //  с адресами, указаннами в settings.
            //  Если количество адресов больше количества параметров процедуры,
            //  выдать ошибку, остановить службу
            int numberOfVars = 0;
            Int32.TryParse(obj.NumberOfVariablesInProcedure, out numberOfVars);

            if (tagsValuesDic.Count > numberOfVars)
            {
                throw new Exception(@"Количество добавленных адресов превышает количество параметров хранимой процедуры БД.
                        Запись данных в БД невозможна.");
            }
            double[] tagsValues = new double[numberOfVars];
            int j = 0;
            foreach (string Tkey in tagsValuesDic.Keys)
            {
                tagsValues[j] = tagsValuesDic[Tkey];
                j++;
            }
            for (; j < numberOfVars; j++)
            {
                tagsValues[j] = 00.00f;
            }

            //Формируем дату последней сводки
            DateTime last2HdateTime = new DateTime(
                Convert.ToInt32(tagsValuesDic[obj.DTYearAddress]),
                Convert.ToInt32(tagsValuesDic[obj.DTMonthAddress]),
                Convert.ToInt32(tagsValuesDic[obj.DTDayAddress]),
                Convert.ToInt32(tagsValuesDic[obj.DTHourAddress]),
                0, 0);

            //Добавляем параметры в команду
            FbParameter par0 = new FbParameter { ParameterName = "CPSID", FbDbType = FirebirdSql.Data.FirebirdClient.FbDbType.Integer, Value = 55 };
            procCommand.Parameters.Add(par0);
            FbParameter par1 = new FbParameter { ParameterName = "Date", FbDbType = FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp, Value = last2HdateTime };
            procCommand.Parameters.Add(par1);

            for (int i = 0; i < numberOfVars; i++)
            {
                FbParameter par = new FbParameter { ParameterName = "P" + i.ToString(), FbDbType = FirebirdSql.Data.FirebirdClient.FbDbType.Double, Value = tagsValues[i] };
                procCommand.Parameters.Add(par);
            }

            //Выполняем команду
            try
            {
                int res = procCommand.ExecuteNonQuery();
                fbTransaction.Commit();
            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка выполнения команды Insert: " + ex.Message);
            }
            finally
            {
                procCommand.Dispose();
                fbConnection.Close();
            }
        }
    }
}
