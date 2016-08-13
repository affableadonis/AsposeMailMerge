using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MailMerge.Libs
{
    public static class ExceptionManager
    {
        public static void LogError(Exception ex)
        {
            try
            {
                string content = "";
                string logPath = HttpContext.Current.Server.MapPath("~/Data/ErrorLog.txt");

                content += DateTime.Now.ToString("dd/MMM/yyyy hh:mm:ss");
                content += Environment.NewLine;
                content += ex.Message + " " + ex.InnerException.Message != null ? ex.InnerException.Message : "";
                content += Environment.NewLine;
                content += ex.StackTrace;
                content += Environment.NewLine;
                content += "********************************************";
                content += Environment.NewLine;

                System.IO.File.AppendAllText(logPath, content);
            }
            catch { }
        }
    }
}