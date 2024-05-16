using NLog;
using NLog.Targets;
using System;
using System.IO;

namespace AGOServer
{
    public class LogManagerCustom
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private static string mainLogFileFullName = "";
        private static string currentLogFileFullName = "";
        public static void InitializeLogger()
        {
            string LogDirectory = Properties.Settings.Default.LogDirectory;
            string LogFileName_Main = Properties.Settings.Default.LogFileName_Main;
            string LogFileLevel = Properties.Settings.Default.LogFileLevel;
            string LogConsoleLevel = Properties.Settings.Default.LogConsoleLevel;

            string LogFilePath = getNewLogFullPath(LogDirectory, LogFileName_Main);
            
            var config = new NLog.Config.LoggingConfiguration();

            var logfile = new NLog.Targets.FileTarget("logfile")
            {   FileName = LogFilePath ,
                Layout = "${date:format=yyyyMMdd-HHmmss}|${level}|${callsite}|${message}"
            };
            var logconsole = new NLog.Targets.ConsoleTarget("logconsole");

            config.AddRule(LogLevel.FromString(LogConsoleLevel), LogLevel.Fatal, logconsole);
            config.AddRule(LogLevel.FromString(LogFileLevel), LogLevel.Fatal, logfile);

            NLog.LogManager.Configuration = config;
            currentLogFileFullName = LogFilePath;
            mainLogFileFullName = LogFilePath;
        }
        public static void changeLogFileName(string newLogName)
        {
            string LogDirectory = Properties.Settings.Default.LogDirectory;
            changeLogFileName(LogDirectory, newLogName);
        }
        public static void changeLogFileName(string logSubDirectory, string newLogName)
        {
            string LogDirectory = Path.Combine(Properties.Settings.Default.LogDirectory, logSubDirectory);
            string logFilePath = getNewLogFullPath(LogDirectory, newLogName);
            logger.Info("changing logfile path to " + logFilePath);
            FileTarget target = LogManager.Configuration.FindTargetByName<FileTarget>("logfile");
            target.FileName = logFilePath;
            currentLogFileFullName = logFilePath;
        }
        public static void resetLogFileFullName()
        {
            FileTarget target = LogManager.Configuration.FindTargetByName<FileTarget>("logfile");
            target.FileName = mainLogFileFullName;
            currentLogFileFullName = mainLogFileFullName;
        }
        private static string getNewLogFullPath(string logDirectory,string logName)
        {
            string logExtension = ".txt";
            string logFilePath = Path.Combine(logDirectory, logName);
            return logFilePath + "_" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + logExtension;
        }
        public static string getCurrentLogFilePath()
        {
            return currentLogFileFullName;
        }
        public static void debug(string message)
        {
            logger.Debug(message);
        }
    }
}
