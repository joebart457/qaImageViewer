using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Globalization;

namespace qaImageViewer.Service
{
    public static class LoggerService
    {
        private static bool _doLogDebug = true;
        public static bool DoLogDebug
        {
            get { return _doLogDebug; }
            set { _doLogDebug = value; }
        }
        private static bool _doLogWarning = true;
        public static bool DoLogWarning
        {
            get { return _doLogWarning; }
            set { _doLogWarning = value; }
        }
        private static bool _doLogError = true;
        public static bool DoLogError
        {
            get { return _doLogError; }
            set { _doLogError = value; }
        }


        static public void Log(string msg)
        {
            if (_doLogDebug)
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter("log.txt", true);

                StringBuilder sb = new StringBuilder();
                sb.Append(GetCurrentTimeString());
                sb.Append(": INFO, ");
                sb.Append(msg);
                file.WriteLine(sb.ToString());
                file.Close();
            }
        }

        static public void LogError(string msg)
        {
            if (_doLogError)
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter("log.txt", true);

                StringBuilder sb = new StringBuilder();
                sb.Append(GetCurrentTimeString());
                sb.Append(": ERROR, ");
                sb.Append(msg);
                file.WriteLine(sb.ToString());
                file.Close();
            }
        }

        static public void LogWarning(string msg)
        {
            if (_doLogWarning)
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter("log.txt", true);

                StringBuilder sb = new StringBuilder();
                sb.Append(GetCurrentTimeString());
                sb.Append(": WARN, ");
                sb.Append(msg);
                file.WriteLine(sb.ToString());
                file.Close();
            }
        }

        public static void ClearLogs()
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("log.txt", false);
            file.Close();
        }

        public static string GetCurrentTimeString()
        {
            DateTime localDate = DateTime.Now;
            CultureInfo culture = new CultureInfo("en-US");
            return localDate.ToString(culture);
        }
    }
}
