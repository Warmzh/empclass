using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Odbc;
using System.Data;
using System.IO;

namespace sftpUplod
{
    class ProgressHelp
    {
        public DataTable QueryProgress(string sqlStr)
        {
            string ControlString1 = "DSN=KSHR;UID=csruser;PWD=csruser";
            OdbcConnection odbcCon = new OdbcConnection(ControlString1);
            try
            {
                //string SqlStr = "select * from PUB.deptcsr WHERE SwipeCardTime>='2017/8/16 00:00:00'";
                OdbcDataAdapter odbcAdapter = new OdbcDataAdapter(sqlStr, odbcCon);
                DataTable ds = new DataTable();
                odbcAdapter.Fill(ds);
                string strFileName = Directory.GetCurrentDirectory() + '\\' + "Log";
                if (Directory.Exists(strFileName))
                {
                    Directory.CreateDirectory(strFileName);
                }
               
                //WriteLog("OK:" + dt.Rows.Count.ToString());

               
                return ds;
            }
            catch (System.Exception ex)
            {
                WriteLog(DateTime.Now + "-->Query Progress Error:" + ex.Message);
                return null;
            }
            finally
            {
                odbcCon.Close();
            }
        }

        

        private void WriteLog(String strLog)
        {
            DateTime currenttime = System.DateTime.Now;
            //create log folder
            string sLogPath = System.Environment.CurrentDirectory + "\\sftplog";
            if (!Directory.Exists(sLogPath))
            {
                Directory.CreateDirectory(sLogPath);
            }
            //write log to local
            StreamWriter sw = null;
            try
            {
                string sTitle = string.Format("{0:yyyyMMdd}", currenttime);
                string LogStr = currenttime + " " + strLog;
                sw = new StreamWriter(sLogPath + "\\" + sTitle + ".txt", true);
                sw.WriteLine(LogStr);
                //sw.WriteLine('\r');
            }
            catch
            {
            }
            finally
            {
                if (sw != null)
                {
                    sw.Close();
                }
            }
        }
    }
}
