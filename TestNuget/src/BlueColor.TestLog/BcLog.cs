using NLog;
using NLog.Targets;
using System.Text;

namespace BlueColor.TestLog
{
    /// <summary>
    /// Bc日志
    /// </summary>
    public class BcLog
    {
        /// <summary>
        /// 返回日志记录器
        /// </summary>
        /// <param name="logFileName">日志文件名</param>
        /// <returns></returns>
        public static Logger GetLogger(string logFileName)
        {
            var target = new FileTarget();

            target.FileName = "${basedir}/logs/${shortdate}/" + logFileName + ".${level}.log";
            target.OpenFileCacheTimeout = 30;
            target.KeepFileOpen = true;
            target.Encoding = Encoding.UTF8;

            NLog.Config.SimpleConfigurator.ConfigureForTargetLogging(target, LogLevel.Trace);

            Logger logger = LogManager.GetLogger(logFileName);

            return logger;
        }
    }
}