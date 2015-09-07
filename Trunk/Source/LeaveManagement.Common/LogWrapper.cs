using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeaveManagement.Common
{
    public static class LogWrapper
    {
        private static Logger _mainLogger;
        private static Logger _usageLogger;

        public static Logger MainLogger
        {
            get
            {
                if (_mainLogger == null)
                {
                    _mainLogger = LogManager.GetLogger(Resources.MainLoggerName);
                }
                return _mainLogger;
            }
        }

        public static Logger UsageLogger
        {
            get
            {
                if (_usageLogger == null)
                {
                    _usageLogger = LogManager.GetLogger(Resources.UsageLoggerName);
                }
                return _usageLogger;
            }
        }
    }
}