using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinForm.ClassLibrary
{
    class LogEventInfoData
    {
        public Dictionary<string, string> LogEventDataDictionary { get; private set; }

        public string FileName
        {
            set
            {
                if (LogEventDataDictionary.ContainsKey("FileName"))
                    LogEventDataDictionary["FileName"] = value;
                else
                {
                    LogEventDataDictionary.Add("FileName", value);
                }
            }
        }

        public string StationId
        {
            set
            {
                if (LogEventDataDictionary.ContainsKey("StationID"))
                    LogEventDataDictionary["StationID"] = value;
                else
                {
                    LogEventDataDictionary.Add("StationID", value);
                }
            }
        }
        public LogEventInfoData(string fileName, string stationId)
        {
            FileName = fileName;
            StationId = stationId;
        }

    }
}
