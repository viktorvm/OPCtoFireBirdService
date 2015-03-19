using System;
using System.Collections.Generic;

using System.Xml.Linq;

namespace OPCtoFirebirdService
{
    static class Settings
    {
        static Settings()
        {
            string filePath = "C:\\OPCtoFirebirdSettings.xml";
            try
            {
                XDocument xdoc = XDocument.Load(filePath);

                HostName = xdoc.Root.Element("HostName").Value;
                OpcServerVendor = xdoc.Root.Element("OpcServerVendor").Value;
                UpdateRate = Convert.ToInt32(xdoc.Root.Element("UpdateRate").Value);
                DatabasePath = xdoc.Root.Element("DatabasePath").Value;
                UserName = xdoc.Root.Element("UserName").Value;
                Password = xdoc.Root.Element("Password").Value;

                XElement objEl = xdoc.Root.Element("Objects");
                Objects = new List<OPCObject>();
                if (objEl.HasElements)
                {
                    foreach (XElement obj in objEl.Elements())
                    {
                        OPCObject opcObject = new OPCObject();

                        opcObject.TableName = obj.Element("TableName").Value;
                        opcObject.StoredProcedureName = obj.Element("StoredProcedureName").Value;
                        opcObject.NumberOfVariablesInProcedure = obj.Element("NumberOfVariablesInProcedure").Value;
                        opcObject.NeedTimeSync = Convert.ToBoolean(obj.Element("NeedTimeSync").Value);
                        opcObject.DTHourAddress = obj.Element("DTHourAddress").Value;
                        opcObject.DTDayAddress = obj.Element("DTDayAddress").Value;
                        opcObject.DTMonthAddress = obj.Element("DTMonthAddress").Value;
                        opcObject.DTYearAddress = obj.Element("DTYearAddress").Value;

                        XElement tagsEl = obj.Element("Tags");
                        opcObject.Tags = new List<string>();

                        //Добавляем адреса, хранящие значения времени
                        opcObject.Tags.Add(opcObject.DTHourAddress);
                        opcObject.Tags.Add(opcObject.DTDayAddress);
                        opcObject.Tags.Add(opcObject.DTMonthAddress);
                        opcObject.Tags.Add(opcObject.DTYearAddress);

                        //Добавляем адреса из ветки Tags
                        if (tagsEl.HasElements)
                        {
                            foreach (XElement el in tagsEl.Elements())
                            {
                                string fullAddress = el.Value.Trim();

                                //если адресс составной, разбираем его
                                if (fullAddress.IndexOf("{") != -1)
                                {
                                    string prefix = "";
                                    string sufix = "";
                                    short startElement = 0;
                                    short endElement = 0;
                                    prefix = fullAddress.Substring(0, fullAddress.IndexOf("{"));
                                    sufix = fullAddress.Substring(fullAddress.IndexOf("}") + 1, fullAddress.Length - fullAddress.IndexOf("}") - 1);
                                    Int16.TryParse(fullAddress.Substring(fullAddress.IndexOf("{") + 1, fullAddress.IndexOf("-") - fullAddress.IndexOf("{") - 1), out startElement);
                                    Int16.TryParse(fullAddress.Substring(fullAddress.IndexOf("-") + 1, fullAddress.IndexOf("}") - fullAddress.IndexOf("-") - 1), out endElement);

                                    if (endElement > 0)
                                    {
                                        for (short i = startElement; i < endElement + 1; i++)
                                        {
                                            opcObject.Tags.Add(prefix + i.ToString() + sufix);
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Ошибка парсинга, проверьте правильность написания адреса - " + fullAddress);
                                    }
                                }
                                //иначе просто добавляем к тегам
                                else
                                {
                                    opcObject.Tags.Add(el.Value);
                                }
                            }
                        }
                        Objects.Add(opcObject);
                    }
                }
            }
            catch (Exception ex)
            {
                HostName = "ERROR";
                OpcServerVendor = "ERROR";
                DatabasePath = "ERROR";
                UserName = "ERROR";
                Password = "ERROR";
                Objects = new List<OPCObject>();

                throw new Exception("Не удалось загрузить конфигурацию: " + ex.Message);
            }
        }

        public static string HostName { get; set; }
        public static string OpcServerVendor { get; set; }
        public static int UpdateRate { get; set; }
        public static string DatabasePath { get; set; }
        public static string UserName { get; set; }
        public static string Password { get; set; }
        public static List<OPCObject> Objects { get; set; }
    }

    public class OPCObject
    {
        public string TableName { get; set; }
        public string StoredProcedureName { get; set; }
        public string NumberOfVariablesInProcedure { get; set; }
        public bool NeedTimeSync { get; set; }
        public List<string> Tags { get; set; }
        public string DTHourAddress { get; set; }
        public string DTDayAddress { get; set; }
        public string DTMonthAddress { get; set; }
        public string DTYearAddress { get; set; }
    }
}
