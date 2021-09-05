using System;
using System.IO;
using System.Reflection;

namespace EguibarIT.Housekeeping
{
    /// <summary>
    ///
    /// </summary>
    public enum LogTarget
    {
        /// <summary>
        ///
        /// </summary>
        File
    }

    /// <summary>
    ///
    /// </summary>
    public abstract class LogBase
    {
        /// <summary>
        ///
        /// </summary>
        protected readonly object lockObj = new object();

        /// <summary>
        ///
        /// </summary>
        /// <param name="message"></param>
        public abstract void Log(string message);
    }

    /// <summary>
    ///
    /// </summary>
    public class FileLogger : LogBase
    {
        /// <summary>
        ///
        /// </summary>
        public string filePath = @"D:\IDGLog.txt";

        private FileStream fs;

        /// <summary>
        ///
        /// </summary>
        /// <param name="message"></param>
        public override void Log(string message)
        {
            if (!File.Exists(filePath)) //No File? Create
            {
                fs = File.Create(filePath);
                fs.Close();
            }
            else
            {
                //fs = File.Open(filePath);

                if (File.ReadAllBytes(filePath).Length >= 100 * 1024 * 1024) // (100mB) File to big? Create new
                {
                    string filenamebase = "myLogFile"; //Insert the base form of the log file, the same as the 1st filename without .log at the end
                    if (File.Exists(filePath).ToString().Contains("-")) //Check if older log contained -x
                    {
                        int lognumber = Int32.Parse(filePath.Substring(filePath.LastIndexOf("-") + 1, filePath.Length - 4)); //Get old number, Can cause exception if the last digits aren't numbers
                        lognumber++; //Increment lognumber by 1
                        filePath = filenamebase + "-" + lognumber + ".log"; //Override filename
                    }
                    else
                    {
                        filePath = filenamebase + "-1.log"; //Override filename
                    }
                    fs = File.Create(filePath);
                    fs.Close();

                    lock (lockObj)
                    {
                        using (StreamWriter streamWriter = new StreamWriter(filePath))
                        {
                            streamWriter.WriteLine(message);
                            streamWriter.Close();
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    ///
    /// </summary>
    public static class LogHelper
    {
        private static LogBase logger = null;

        /// <summary>
        ///
        /// </summary>
        /// <param name="target"></param>
        /// <param name="message"></param>
        public static void Log(LogTarget target, string message)
        {
            switch (target)
            {
                case LogTarget.File:
                    logger = new FileLogger();
                    logger.Log(message);
                    break;

                default:
                    return;
            }
        }
    }

    /// <summary>
    ///
    /// </summary>
    public static class LogWriter
    {
        private static string m_exePath = string.Empty;

        /// <summary>
        ///
        /// </summary>
        /// <param name="logMessage"></param>
        public static void LogWrite(string logMessage)
        {
            m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!File.Exists(m_exePath + "\\" + "log.txt"))
                File.Create(m_exePath + "\\" + "log.txt");

            try
            {
                using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
                    AppendLog(logMessage, w);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void AppendLog(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.Write("\r\nLog Entry : ");
                txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                txtWriter.WriteLine("  :");
                txtWriter.WriteLine("  :{0}", logMessage);
                txtWriter.WriteLine("-------------------------------");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}