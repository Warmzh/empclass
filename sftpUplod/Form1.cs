using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using MySql.Data.MySqlClient;
using System.Data.Odbc;

namespace sftpUplod
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //if(File.Exists(strFilePath))
            //{
            //    txtReadPath.Text=inifile.IniReadValue("Path","ReadPath",strFilePath);
            //    txtBackupPath.Text = inifile.IniReadValue("Path", "BackupPath", strFilePath);
            //    //txtTime.Text = inifile.IniReadValue("QueryTime", "QTime", strFilePath);
            //}
            WriteLog("Open Auto======================" + '\r', 1);


            //InsertMySql_testemployee_SqlServer("CSBG");
            //InsertMySql_testemployee_SqlServer("ASBG");

           //InsertMySql_empClass0330("CSBG");
             //InsertMySql_empClass0330("ASBG");

             //InsertMySql_empClass1530("CSBG");
             //InsertMySql_empClass1530("ASBG");
            //InsertMysql_empClassFromProgress("CSBG");
            //InsertMySql_empClass("CSBG");
            // InsertMySql_empClass("ASBG");
           // InsertMySql_LmtDept("CSBG");
            //InsertMySql_LmtDept("ASBG");
           //MySql_testswipecardtimeTOprogress_CARDSR("CSBG");
           // MySql_testswipecardtimeTOprogress_CARDSR("ASBG");
            //MySql_testswipecardtimeTOprogress_CARDSR("BBBB");
            //Mysql_rawToProgress("CSBG");
            //Mysql_rawToProgress("ASBG");
            //InsertMysql_empClass4("CSBG");
            //InsertMysql_empClass4("ASBG");
            // this.Close();

        }

        /// <summary>
        /// 开始按钮逻辑
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, EventArgs e)
        {


            if (btnStart.Text.Equals("Start"))
            {
                //检查路径是否正确
                //if (!CheckPath()) return;

                btnStart.Text = "Pause";
                // WriteLog("01s:00 InsertMySQL，10:00 InsertProgress" + '\r', 1);
                timer1.Enabled = true;
            }
            else
            {
                btnStart.Text = "Start";
                WriteLog("Pause Read..." + '\r', 1);
                timer1.Enabled = false;
            }
        }




        /// <summary>
        /// 輸出訊息在畫面上,並寫入Log檔
        /// </summary>
        /// <param name="intColorType">0:Gray,1:Blue,2:Red</param>
        private void WriteLog(String sMessage, Int32 intColorType = 0)
        {
            DateTime currenttime = System.DateTime.Now;
            string sDateTime = string.Format("{0:yyyy-MM-dd HH:mm:ss}", currenttime);
            RichTxtLog.AppendText(sDateTime + " => " + sMessage);
            //RichTxtLog.Text=sMessage+'\r'+RichTxtLog.Text;
            RichTxtLog.Select(RichTxtLog.TextLength - sMessage.Length, RichTxtLog.TextLength);
            switch (intColorType)
            {
                //case 0:
                //    //一般訊息
                //    RichTxtLog.SelectionColor = Color.Gray;
                //    break;
                case 1:
                    //上傳成功
                    RichTxtLog.SelectionColor = Color.Green;
                    break;
                case 2:
                    //出錯
                    RichTxtLog.SelectionColor = Color.Red;
                    break;
            }
            ////保留最新的100筆訊息
            //String[] lines = new String[100];
            //if (RichTxtLog.Lines.Length > lines.Length)
            //{
            //    Array.Copy(RichTxtLog.Lines, lines, lines.Length);
            //    this.RichTxtLog.Lines = lines;
            //}

            if (RichTxtLog.Lines.Length > 100)
            {
                string[] sLines = RichTxtLog.Lines;
                string[] sNewLines = new string[sLines.Length - 100];


                Array.Copy(sLines, 100, sNewLines, 0, sNewLines.Length);
                RichTxtLog.Lines = sNewLines;
            }

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
                string LogStr = currenttime.ToString("HH:mm:ss:fffffff") + " " + sMessage;
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

        private void RichTxtLog_TextChanged(object sender, EventArgs e)
        {

            RichTxtLog.SelectionStart = RichTxtLog.Text.Length; //Set the current caret position at the end     
            RichTxtLog.ScrollToCaret();
            //RichTxtLog.ScrollToCaret(); //Now scroll it automatically
        }




        /// <summary>
        /// 定时器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {

            timer1.Enabled = true;


            #region 01:00 
            if (DateTime.Now.ToString("HH:mm") == "03:30")
            //|| DateTime.Now.ToString("HH:mm") == "09:00" || DateTime.Now.ToString("HH:mm") == "13:00" || DateTime.Now.ToString("HH:mm") == "17:00"
            {
                InsertMySql_empClass0330("CSBG");
                InsertMySql_empClass0330("ASBG");

                //InsertMySql_empClass1530("CSBG");
                //InsertMySql_empClass1530("ASBG");

                InsertMySql_LmtDept("CSBG");
                InsertMySql_LmtDept("ASBG");



            }
            if (DateTime.Now.ToString("HH:mm") == "15:30")
            //|| DateTime.Now.ToString("HH:mm") == "09:00" || DateTime.Now.ToString("HH:mm") == "13:00" || DateTime.Now.ToString("HH:mm") == "17:00"
            {
                //InsertMySql_empClass0330("CSBG");
                //InsertMySql_empClass0330("ASBG");

                InsertMySql_empClass1530("CSBG");
                InsertMySql_empClass1530("ASBG");
                InsertMySql_LmtDept("CSBG");
                InsertMySql_LmtDept("ASBG");

            }
            #endregion

            if (DateTime.Now.ToString("HH:mm") == "10:00")
            {
                MySql_testswipecardtimeTOprogress_CARDSR("CSBG");
                MySql_testswipecardtimeTOprogress_CARDSR("ASBG");



            }
            if (DateTime.Now.ToString("HH:mm") == "10:20")
            {
                Mysql_rawToProgress("CSBG");
                Mysql_rawToProgress("ASBG");
            }

            if (DateTime.Now.Minute == 30 || DateTime.Now.Minute == 00)
            {
                InsertMySql_testemployee_SqlServer("CSBG");
                InsertMySql_testemployee_SqlServer("ASBG");
                InsertMysql_empClass4("CSBG");
                InsertMysql_empClass4("ASBG");
            }



            /* if (DateTime.Now.ToString("HH:mm") =="12:00") {
                 InsertMySql_LmtDept("CSBG");
                 InsertMySql_LmtDept("ASBG");
             }
             */
            //timer1.Enabled = true;


        }
        private void InsertMysql_empClass4(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {

                String connsql = "server=192.168.78.152;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                //int UpdateSumk = 0;
                //int InsertSumk = 0;
                int Insert4 = 0;
                string strSQLUpdate = "";
                try
                {
                    //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為4
                    //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為502
                    strSQLUpdate = "select emp.id from swipecard.testemployee emp where emp.isonwork = 0 and emp.id not in (select class.id from swipecard.emp_class class where class.emp_date = curdate())";
                    DataTable DTID = MySqlHelp.QueryDBC(strSQLUpdate);
                    foreach (DataRow dr in DTID.Rows)
                    {
                        strSQLUpdate = "insert into swipecard.emp_class values ('" + dr["ID"].ToString() + "',curdate(),'502',curdate())";
                        cmd.CommandText = strSQLUpdate;
                        cmd.ExecuteNonQuery();
                        Insert4++;
                    }
                    tx.Commit();
                    WriteLog("-->Update152_EmpClass4:" + Insert4 + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert152_EmpClass4Error,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                String connsql = "server=192.168.78.153;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                //int UpdateSumk = 0;
                //int InsertSumk = 0;
                int Insert4 = 0;
                string strSQLUpdate = "";
                try
                {
                    //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為4
                    //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為502
                    strSQLUpdate = "select emp.id from swipecard.testemployee emp where emp.isonwork = 0 and emp.id not in (select class.id from swipecard.emp_class class where class.emp_date = curdate())";
                    DataTable DTID = MySqlHelp.QueryDbAsbg(strSQLUpdate);
                    foreach (DataRow dr in DTID.Rows)
                    {
                        strSQLUpdate = "insert into swipecard.emp_class values ('" + dr["ID"].ToString() + "',curdate(),'502',curdate())";
                        cmd.CommandText = strSQLUpdate;
                        cmd.ExecuteNonQuery();
                        Insert4++;
                    }
                    tx.Commit();
                    WriteLog("-->Update153_EmpClass4:" + Insert4 + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert153_EmpClass4Error,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }

        }
        private void Mysql_rawToProgress(string strBG)
        {
            string sql = @"SELECT id,
       cardid,
       swipecardtime,
       update_time
      FROM swipecard.raw_record
       WHERE     swipecardtime >= DATE_ADD(DATE_SUB(CURDATE(), INTERVAL 1 DAY), INTERVAL 10 HOUR)
       AND swipecardtime < DATE_ADD(CURDATE(), INTERVAL 10 HOUR)";
            //            string sql = @"SELECT id,
            //       cardid,
            //       swipecardtime,
            //       update_time
            //  FROM swipecard.raw_record
            // WHERE swipecardtime < DATE_ADD(CURDATE(), INTERVAL 10 HOUR)";
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                #region CSBG
                DataTable DT2 = MySqlHelp.QueryDBC(sql);
                if (DT2 == null) { WriteLog("-->Query 152RawError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 152RawSum:0" + '\r', 2); return; }
                WriteLog("-->Query 152RawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "DSN=KSHR;UID=csruser;PWD=csruser";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {
                    string sa = "";
                    string sb = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                        sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                        string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                       + sb + ",'"
                                                                                                                       + dr["swipecardtime"].ToString() + "','"
                                                                                                                       + dr["update_time"].ToString() + "')";

                        cmd.CommandText = strSQL;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    tx.Commit();

                    WriteLog("-->152RawToProgressOK" + '\r', 1);
                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->152RawToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
                #endregion
            }
            else if (strBG == "ASBG")
            {
                #region ASBG
                DataTable DT2 = MySqlHelp.QueryDbAsbg(sql);
                if (DT2 == null) { WriteLog("-->Query 153RawError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 153RawSum:0" + '\r', 2); return; }
                WriteLog("-->Query 153RawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "DSN=KSHR;UID=csruser;PWD=csruser";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {
                    string sa = "";
                    string sb = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                        sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                        string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                       + sb + ",'"
                                                                                                                       + dr["swipecardtime"].ToString() + "','"
                                                                                                                       + dr["update_time"].ToString() + "')";

                        cmd.CommandText = strSQL;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    tx.Commit();

                    WriteLog("-->153RawToProgressOK" + '\r', 1);
                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->153RawToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
                #endregion
            }


        }
        private void MySql_testswipecardtimeTOprogress_CARDSR(string strBG)
        {
            //            string sql = @"SELECT emp.ID, cardtime.*
            //  FROM swipecard.testemployee emp, swipecard.testswipecardtime cardtime
            // WHERE     emp.cardid = cardtime.CardID
            //       AND (   (DATE(swipecardtime) =
            //                   DATE_SUB(CURDATE(), INTERVAL 1 DAY))
            //            OR (    DATE(swipecardtime2) =
            //                       DATE_SUB(CURDATE(), INTERVAL 1 DAY)
            //                AND cardtime.SwipeCardTime IS NULL)
            //            OR (    DATE(swipecardtime2) = CURDATE()
            //                AND shift = 'N'))";

            string sql = @"SELECT cardtime.ID , cardtime.RecordID, cardtime.CardID, cardtime.name, cardtime.SwipeCardTime, cardtime.SwipeCardTime2, cardtime.CheckState, cardtime.PROD_LINE_CODE, cardtime.WorkshopNo, cardtime.PRIMARY_ITEM_NO, cardtime.RC_NO, cardtime.Shift, cardtime.OvertimeState, cardtime.overtimeType, cardtime.overtimeCal 
  FROM  swipecard.testswipecardtime cardtime
 WHERE   (   (DATE_FORMAT(cardtime.swipecardtime, '%Y%m%d') =
                   DATE_SUB(CURDATE(), INTERVAL 1 DAY))
            OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') =
                       DATE_SUB(CURDATE(), INTERVAL 1 DAY)
                AND cardtime.SwipeCardTime IS NULL and Shift ='D')
            OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') = CURDATE()
                AND shift = 'N' and cardtime.SwipeCardTime is NULL))";
            string sql2 = @"SELECT emp.ID,
       cardtime.RecordID,
       cardtime.CardID,
       cardtime.name,
       cardtime.SwipeCardTime,
       cardtime.SwipeCardTime2,
       cardtime.CheckState,
       cardtime.PROD_LINE_CODE,
       cardtime.WorkshopNo,
       cardtime.PRIMARY_ITEM_NO,
       cardtime.RC_NO,
       cardtime.Shift,
       cardtime.OvertimeState,
       cardtime.overtimeType,
       cardtime.overtimeCal
  FROM swipecard.testemployee emp, swipecard.testswipecardtime cardtime
 WHERE     emp.cardid = cardtime.CardID
       AND (   (DATE_FORMAT(swipecardtime, '%Y%m%d') =
                   DATE_SUB(CURDATE(), INTERVAL 1 DAY))
            OR (    DATE_FORMAT(swipecardtime2, '%Y%m%d') =
                       DATE_SUB(CURDATE(), INTERVAL 1 DAY)
                AND cardtime.SwipeCardTime IS NULL)
            OR (    DATE_FORMAT(swipecardtime2, '%Y%m%d') = CURDATE()
                AND shift = 'N'))";
            //            string sql2 = @"SELECT DISTINCT RecordID
            //  FROM pub.cardsr
            // WHERE     to_char(swipecardtime2, 'YYYYMMDD') =
            //              TO_CHAR(sysdate - 1, 'YYYYMMDD')
            //       AND SwipeCardTime IS NULL";
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
                DataTable DT2 = MySqlHelp.QueryDBC(sql);
                if (DT2 == null) { WriteLog("-->Query 152wipecardError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 152wipecardSum:0" + '\r', 2); return; }

                //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 
                WriteLog("-->Query 152wipecardSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                //DataTable DT3 = ProHelp.QueryProgress(sql2);
                //if (DT3 == null)
                //{
                //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
                //    return;
                //}
                //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键 

                //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "DSN=KSHR;UID=csruser;PWD=csruser";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {

                    string sa = "";
                    string sb = "";
                    string ss = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                        //if (Progressdr == null)
                        //{
                        ss = dr[0].Equals(DBNull.Value) ? "NULL" : dr[0].ToString();
                        sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                        sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                        string strSQL = "insert into pub.CARDSR  VALUES  ('" + ss + "','"
                                                                             + dr[1].ToString() + "','"
                                                                             + dr[2].ToString() + "','"
                                                                             + dr[3].ToString() + "',"
                                                                             + sa + ","
                                                                             + sb + ",'"
                                                                             + dr[6].ToString() + "','"
                                                                             + dr[7].ToString() + "','"
                                                                             + dr[8].ToString() + "','"
                                                                             + dr[9].ToString() + "','"
                                                                             + dr[10].ToString() + "','"
                                                                             + dr[11].ToString() + "','"
                                                                             + dr[12].ToString() + "','"
                                                                             + dr[13].ToString() + "','"
                                                                             + dr[14].ToString() + "')";

                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //}

                    }


                    tx.Commit();

                    WriteLog("-->152WipToProgressOK" + '\r', 1);



                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->153WipToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
            }
            else if (strBG == "ASBG")
            {
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
                DataTable DT2 = MySqlHelp.QueryDbAsbg(sql);
                if (DT2 == null) { WriteLog("-->Query 153wipecardError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 153wipecardSum:0" + '\r', 2); return; }

                //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 
                WriteLog("-->Query 153wipecardSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                //DataTable DT3 = ProHelp.QueryProgress(sql2);
                //if (DT3 == null)
                //{
                //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
                //    return;
                //}
                //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键
                //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "DSN=KSHR;UID=csruser;PWD=csruser";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {

                    string sa = "";
                    string sb = "";
                    string ss = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                        //if (Progressdr == null)
                        //{
                        ss = dr[0].Equals(DBNull.Value) ? "NULL" : dr[0].ToString();
                        sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                        sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                        string strSQL = "insert into pub.CARDSR  VALUES  ('" + ss + "','"
                                                                             + dr[1].ToString() + "','"
                                                                             + dr[2].ToString() + "','"
                                                                             + dr[3].ToString() + "',"
                                                                             + sa + ","
                                                                             + sb + ",'"
                                                                             + dr[6].ToString() + "','"
                                                                             + dr[7].ToString() + "','"
                                                                             + dr[8].ToString() + "','"
                                                                             + dr[9].ToString() + "','"
                                                                             + dr[10].ToString() + "','"
                                                                             + dr[11].ToString() + "','"
                                                                             + dr[12].ToString() + "','"
                                                                             + dr[13].ToString() + "','"
                                                                             + dr[14].ToString() + "')";

                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //}
                    }


                    tx.Commit();
                    WriteLog("-->153WipToProgressOK" + '\r', 1);



                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->153WipToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
            }


        }

        /* private void MySql_testswipecardtimeTOprogress_CARDSR(string strBG)
         {
             MySqlHelper MySqlHelp = new MySqlHelper();
             if (strBG == "CSBG")
             {
                 //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
                 string sql = "select cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
                 //string sql = "select cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID";
                 DataTable DT2 = MySqlHelp.QueryDBC(sql);
                 if (DT2 == null)
                 {
                     WriteLog("-->Query 152swipecardError" + '\r', 2); return;
                 }
                 else if (DT2.Rows.Count == 0)
                 {
                     WriteLog("-->Query 152swipecardSum:0" + '\r', 2); return;
                 }
                 WriteLog("-->Query 152swipecardSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                 string ControlString1 = "DSN=KSHR;UID=csruser;PWD=csruser";
                 OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                 odbcCon.Open();
                 OdbcTransaction tx = odbcCon.BeginTransaction();
                 OdbcCommand cmd = new OdbcCommand();
                 cmd.Connection = odbcCon;
                 cmd.Transaction = tx;
                 try
                 {

                     string sa = "";
                     string sb = "";
                     foreach (DataRow dr in DT2.Rows)
                     {
                         sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                         sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                         string strSQL = "insert into pub.CARDSR  VALUES  ('" + dr[1].ToString() + "','"
                                                                              + dr[0].ToString() + "','"
                                                                              + dr[2].ToString() + "','"
                                                                              + dr[3].ToString() + "',"
                                                                              + sa + ","
                                                                              + sb + ",'"
                                                                              + dr[6].ToString() + "','"
                                                                              + dr[7].ToString() + "','"
                                                                              + dr[8].ToString() + "','"
                                                                              + dr[9].ToString() + "','"
                                                                              + dr[10].ToString() + "','"
                                                                              + dr[11].ToString() + "','"
                                                                              + dr[12].ToString() + "','"
                                                                              + dr[13].ToString() + "','"
                                                                              + dr[14].ToString() + "')";

                         cmd.CommandText = strSQL;
                         cmd.ExecuteNonQuery();

                     }


                     tx.Commit();

                     WriteLog("-->152WipToProgressOK" + '\r', 1);



                 }
                 catch (System.Exception ex)
                 {
                     tx.Rollback();
                     WriteLog("-->152WipToProgressError" + ex.Message + '\r', 2);
                 }
                 finally
                 {
                     tx.Dispose();
                     odbcCon.Close();

                 }
             }
             else if (strBG == "ASBG")
             {
                // string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
                  string sql = "select cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
                 DataTable DT2 = MySqlHelp.QueryDbAsbg(sql);
                 if (DT2 == null) { WriteLog("-->Query 153wipecardError" + '\r', 2); return; }
                 else if (DT2.Rows.Count == 0) { WriteLog("-->Query 112wipecardSum:0" + '\r', 2); return; }
                 WriteLog("-->Query 153wipecardSum:" + DT2.Rows.Count.ToString() + '\r', 1);
                 string ControlString1 = "DSN=KSHR;UID=csruser;PWD=csruser";
                 OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                 odbcCon.Open();
                 OdbcTransaction tx = odbcCon.BeginTransaction();
                 OdbcCommand cmd = new OdbcCommand();
                 cmd.Connection = odbcCon;
                 cmd.Transaction = tx;
                 try
                 {

                     string sa = "";
                     string sb = "";
                     foreach (DataRow dr in DT2.Rows)
                     {
                         sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                         sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                         string strSQL = "insert into pub.CARDSR  VALUES  ('" + dr[1].ToString() + "','"
                                                                              + dr[0].ToString() + "','"
                                                                              + dr[2].ToString() + "','"
                                                                              + dr[3].ToString() + "',"
                                                                              + sa + ","
                                                                              + sb + ",'"
                                                                              + dr[6].ToString() + "','"
                                                                              + dr[7].ToString() + "','"
                                                                              + dr[8].ToString() + "','"
                                                                              + dr[9].ToString() + "','"
                                                                              + dr[10].ToString() + "','"
                                                                              + dr[11].ToString() + "','"
                                                                              + dr[12].ToString() + "','"
                                                                              + dr[13].ToString() + "','"
                                                                              + dr[14].ToString() + "')";

                         cmd.CommandText = strSQL;
                         cmd.ExecuteNonQuery();

                     }


                     tx.Commit();
                     WriteLog("-->153WipToProgressOK" + '\r', 1);



                 }
                 catch (System.Exception ex)
                 {
                     tx.Rollback();
                     WriteLog("-->153WipToProgressError" + ex.Message + '\r', 2);
                 }
                 finally
                 {
                     tx.Dispose();
                     odbcCon.Close();

                 }
             }


         }*/

        private void InsertMySql_LmtDept(string strBG)
        {
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                string SQL = "SELECT bmbh, bmmc, kbmbh, kbmmc FROM PUB.deptcsr";
                DataTable DT1 = ProHelp.QueryProgress(SQL);
                if (DT1 == null)
                {
                    WriteLog("-->Query ProgressDeptError" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressDeptSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressDeptSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.78.152;userid=root;password=foxlink;database=swipecard;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                try
                {
                    string strSQL = "delete from swipecard.lmt_dept";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    string sa = "";
                    string sb = "";
                    string sc = "";
                    strSQL = "insert into swipecard.lmt_dept VALUES ";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        sa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + dr[1].ToString() + "'";
                        sb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + dr[2].ToString() + "'";
                        sc = dr[3].Equals(DBNull.Value) ? "NULL" : "'" + dr[3].ToString() + "'";
                        strSQL += "  ('" + dr[0].ToString() + "',"
                                         + sa + ","
                                         + sb + ","
                                         + sc + "),";
                    }
                    strSQL = strSQL.Substring(0, strSQL.Length - 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();
                    tx.Commit();
                    WriteLog("-->insert152_DeptOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert152_DeptError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                string SQL = "SELECT bmbh, bmmc, kbmbh, kbmmc FROM PUB.deptcsr";
                DataTable DT1 = ProHelp.QueryProgress(SQL);
                if (DT1 == null)
                {
                    WriteLog("-->Query ProgressDeptError" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressDeptSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressdeptSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.78.153;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                try
                {
                    string strSQL = "delete from swipecard.lmt_dept";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    string sa = "";
                    string sb = "";
                    string sc = "";
                    strSQL = "insert into swipecard.lmt_dept VALUES ";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        sa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + dr[1].ToString() + "'";
                        sb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + dr[2].ToString() + "'";
                        sc = dr[3].Equals(DBNull.Value) ? "NULL" : "'" + dr[3].ToString() + "'";
                        strSQL += "  ('" + dr[0].ToString() + "',"
                                         + sa + ","
                                         + sb + ","
                                         + sc + "),";
                    }
                    strSQL = strSQL.Substring(0, strSQL.Length - 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();
                    tx.Commit();
                    WriteLog("-->insert153_DeptOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert153_DeptError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }

        }
        // private void InsertMysql_empNo(string sqlStr) {
        //     ProgressHelp progressHelp = new ProgressHelp();
        //     MySqlHelper mySqlHelper = new MySqlHelper();

        //  }
        private void InsertMySql_testemployee_SqlServer(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                //78.152
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 152TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                //String sql = "SELECT 編號 bh, 卡號 kh FROM V_RYBZ_TX"; // 查询语句
                String sql = "SELECT zgbh, icid FROM mf_employee where lrzx in ('PQ','MQ')";
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["zgbh"] };

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD,'cardid', EXP_DEPT, EMP_STATUS FROM PUB.EMPR";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                    return;
                }

                WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.78.152;userid=root;password=foxlink;database=;charset=utf8";
                //String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLUpdateIswork = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        DataRow MSdr = DT2.Rows.Find(dr["ID"].ToString());
                        if (MSdr == null)
                        {
                            strSQLUpdateIswork = "update swipecard.testemployee SET isOnWork =1,updateDate =curdate() where isOnWork=0 and ID = '" + dr["ID"].ToString() + "'";
                            cmd.CommandText = strSQLUpdateIswork;
                            int i = cmd.ExecuteNonQuery();
                            // if (i > 0) { WriteLog("-->Update152_" + dr["ID"].ToString() + " Iswork=1OK:" + UpdateSumk + strSQLUpdate + '\r', 1); }
                        }

                    }
                    foreach (DataRow dr in DT3.Rows)
                    {
                        //DataRow[] Tempdr = DT1.Select("ID='" + dr["EMP_CD"].ToString() + "'");
                        DataRow SqlServerdr = DT2.Rows.Find(dr["EMP_CD"].ToString());
                        if (SqlServerdr != null)
                        {
                            DataRow Tempdr = DT1.Rows.Find(dr["EMP_CD"].ToString());
                            if (Tempdr != null)
                            {
                                if (Tempdr["ID"].ToString() != dr["EMP_CD"].ToString() || Tempdr["Name"].ToString() != dr["EMP_NAME"].ToString() ||
                                    Tempdr["depid"].ToString() != dr["DEPT_CD"].ToString() || Tempdr["depname"].ToString() != dr["DEPT_NAME"].ToString() ||
                                    Tempdr["Direct"].ToString() != dr["D_I_CD"].ToString() || Tempdr["cardid"].ToString() != SqlServerdr["icid"].ToString() ||
                                    Tempdr["costID"].ToString() != dr["EXP_DEPT"].ToString() || Tempdr["isOnWork"].ToString() == dr["EMP_STATUS"].ToString())
                                {
                                    if (dr["EMP_STATUS"].ToString() == string.Empty)
                                    { Istatus = 0; }
                                    else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                    { Istatus = 1; }
                                    else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                    { Istatus = 0; }
                                    strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["EMP_NAME"].ToString() +
                                                                                  "',depid = '" + dr["DEPT_CD"].ToString() +
                                                                                  "',depname = '" + dr["DEPT_NAME"].ToString() +
                                                                                  "',Direct = '" + dr["D_I_CD"].ToString() +
                                                                                  "',cardid = '" + SqlServerdr["icid"].ToString() +
                                                                                  "',costID = '" + dr["EXP_DEPT"].ToString() +
                                                                                  "',isOnWork = " + Istatus + "  ,updateDate =curdate() WHERE ID = '" + dr["EMP_CD"].ToString() + "' ";
                                    cmd.CommandText = strSQLUpdate;
                                    cmd.ExecuteNonQuery();
                                    UpdateSumk++;
                                    //WriteLog("-->Update111_EmpOK:" + UpdateSumk + strSQLUpdate + '\r', 1);
                                }
                            }
                            else
                            {
                                ++InsertSumk;
                                strSQLInsert += " ('" + dr["EMP_CD"].ToString() + "','"
                                                      + dr["EMP_NAME"].ToString() + "','"
                                                      + dr["DEPT_CD"].ToString() + "','"
                                                      + dr["DEPT_NAME"].ToString() + "','"
                                                      + dr["D_I_CD"].ToString() + "','"
                                                      + SqlServerdr["icid"].ToString() + "','"
                                                      + dr["EXP_DEPT"].ToString() + "',"
                                                      + "curdate()" + "),";
                                //WriteLog("-->Insert111_EmpOK:" + InsertSumk + " ('" + dr["EMP_CD"].ToString() + "','"
                                //                      + dr["EMP_NAME"].ToString() + "','"
                                //                      + dr["DEPT_CD"].ToString() + "','"
                                //                      + dr["DEPT_NAME"].ToString() + "','"
                                //                      + dr["D_I_CD"].ToString() + "','"
                                //                      + SqlServerdr["kh"].ToString() + "','"
                                //                      + dr["EXP_DEPT"].ToString() + "',"
                                //                      + "curdate()" + '\r', 1);
                                if (InsertSumk > 0 && InsertSumk % 1000 == 0)
                                {
                                    strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                    cmd.CommandText = strSQLInsert;
                                    cmd.ExecuteNonQuery();

                                    strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                                }
                            }
                        }
                    }
                    if (InsertSumk % 1000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update152_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert152_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert152_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }

            }
            else if (strBG == "ASBG")
            {
                //78.153
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 153TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                String sql = "SELECT zgbh, icid FROM mf_employee where lrzx in ('IR','AI')"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["zgbh"] };

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD,'cardid', EXP_DEPT, EMP_STATUS FROM PUB.EMPR";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                    return;
                }

                WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                //String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                String connsql = "server=192.168.78.153;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    string strSQLUpdateIswork = "";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        DataRow MSdr = DT2.Rows.Find(dr["ID"].ToString());
                        if (MSdr == null)
                        {
                            strSQLUpdateIswork = "update swipecard.testemployee SET isOnWork =1,updateDate =curdate() where isOnWork=0 and ID = '" + dr["ID"].ToString() + "'";
                            cmd.CommandText = strSQLUpdateIswork;
                            int i = cmd.ExecuteNonQuery();
                            // if (i > 0) { WriteLog("-->Update153_" + dr["ID"].ToString() + " Iswork=1OK:" + UpdateSumk + strSQLUpdate + '\r', 1); }
                        }

                    }
                    foreach (DataRow dr in DT3.Rows)
                    {
                        //DataRow[] Tempdr = DT1.Select("ID='" + dr["EMP_CD"].ToString() + "'");
                        DataRow SqlServerdr = DT2.Rows.Find(dr["EMP_CD"].ToString());
                        if (SqlServerdr != null)
                        {
                            DataRow Tempdr = DT1.Rows.Find(dr["EMP_CD"].ToString());
                            if (Tempdr != null)
                            {
                                if (Tempdr["ID"].ToString() != dr["EMP_CD"].ToString() || Tempdr["Name"].ToString() != dr["EMP_NAME"].ToString() ||
                                    Tempdr["depid"].ToString() != dr["DEPT_CD"].ToString() || Tempdr["depname"].ToString() != dr["DEPT_NAME"].ToString() ||
                                    Tempdr["Direct"].ToString() != dr["D_I_CD"].ToString() || Tempdr["cardid"].ToString() != SqlServerdr["icid"].ToString() ||
                                    Tempdr["costID"].ToString() != dr["EXP_DEPT"].ToString() || Tempdr["isOnWork"].ToString() == dr["EMP_STATUS"].ToString())
                                {
                                    if (dr["EMP_STATUS"].ToString() == string.Empty)
                                    { Istatus = 0; }
                                    else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                    { Istatus = 1; }
                                    else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                    { Istatus = 0; }
                                    strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["EMP_NAME"].ToString() +
                                                                                  "',depid = '" + dr["DEPT_CD"].ToString() +
                                                                                  "',depname = '" + dr["DEPT_NAME"].ToString() +
                                                                                  "',Direct = '" + dr["D_I_CD"].ToString() +
                                                                                  "',cardid = '" + SqlServerdr["icid"].ToString() +
                                                                                  "',costID = '" + dr["EXP_DEPT"].ToString() +
                                                                                  "',isOnWork = " + Istatus + "  ,updateDate =curdate() WHERE ID = '" + dr["EMP_CD"].ToString() + "' ";
                                    cmd.CommandText = strSQLUpdate;
                                    cmd.ExecuteNonQuery();
                                    UpdateSumk++;
                                    //WriteLog("-->Update112_EmpOK:" + UpdateSumk + strSQLUpdate + '\r', 1);
                                }
                            }
                            else
                            {
                                ++InsertSumk;
                                strSQLInsert += " ('" + dr["EMP_CD"].ToString() + "','"
                                                      + dr["EMP_NAME"].ToString() + "','"
                                                      + dr["DEPT_CD"].ToString() + "','"
                                                      + dr["DEPT_NAME"].ToString() + "','"
                                                      + dr["D_I_CD"].ToString() + "','"
                                                      + SqlServerdr["icid"].ToString() + "','"
                                                      + dr["EXP_DEPT"].ToString() + "',"
                                                      + "curdate()" + "),";
                                //WriteLog("-->Insert112_EmpOK:" + InsertSumk + " ('" + dr["EMP_CD"].ToString() + "','"
                                //                      + dr["EMP_NAME"].ToString() + "','"
                                //                      + dr["DEPT_CD"].ToString() + "','"
                                //                      + dr["DEPT_NAME"].ToString() + "','"
                                //                      + dr["D_I_CD"].ToString() + "','"
                                //                      + SqlServerdr["kh"].ToString() + "','"
                                //                      + dr["EXP_DEPT"].ToString() + "',"
                                //                      + "curdate()" + '\r', 1);

                                if (InsertSumk > 0 && InsertSumk % 1000 == 0)
                                {
                                    strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                    cmd.CommandText = strSQLInsert;
                                    cmd.ExecuteNonQuery();

                                    strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                                }
                            }
                        }
                    }
                    if (InsertSumk % 1000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update153_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert153_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert153_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }

            }

        }
        private void InsertMySql_empClass1530(string strBG)
        {
            //如果Mysql有当天资料但是班别为空，则update，如果Mysql不存在当天资料则Insert
            //Update和Insert明天资料。
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) order by ID, emp_date";
            //string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where  emp_date=date_add(CURDATE(),interval 1 day) order by ID, emp_date";
            if (strBG == "CSBG")
            {
                #region CSBG
                //60.111

                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 152ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
                WriteLog("-->Query 152ClassSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                //String sql = "SELECT DISTINCT 編號  FROM V_RYBZ_TX"; // 查询语句
                ////String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                //if (DT2 == null)
                //{
                //    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                //    return;
                //}
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;
                //}
                //string strEMP = "";
                //foreach (DataRow dr in DT2.Rows)
                //{
                //    strEMP += " EMP_CD='" + dr[0].ToString() + "' OR";
                //}
                //strEMP = strEMP.Substring(0, strEMP.Length - 2);
                //WriteLog("-->SelectSqlServerEmpSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, START_DATE, CLASS_CD from PUB.EMPCLASSR ";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEmpClassSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEmpClassSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.78.152;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 30;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {
                        //WriteLog(dr[0].ToString() + '\r', 1);
                        TimeSpan ts = DateTime.Now.Date - Convert.ToDateTime(dr[1].ToString());
                        int d = ts.Days;
                        string strClass = dr[2].ToString().Substring(0, dr[2].ToString().Length - 1);
                        List<string> list = strClass.Split(new string[] { "," }, StringSplitOptions.None).ToList<string>();
                        for (int i = 0; i < 2; i++)
                        {
                            if ((d + i) >= list.Count) continue;
                            if (!ClassArray3.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + list[d + i]))
                            {
                                if (ClassArray2.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                                {
                                    string sa = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    if (i == 0)
                                    { strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where (class_no ='' or class_no='502') and ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'"; }
                                    else
                                    { strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'"; }
                                    cmd.CommandText = strSQLUpdate;
                                    int it = cmd.ExecuteNonQuery();
                                    if (it > 0) UpdateSumk++;
                                    //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                                }
                                else
                                {

                                    ++InsertSumk;
                                    string sb = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                         + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                         + sb + "',"
                                                         + "curdate()" + "),";

                                    if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                                    {
                                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                        cmd.CommandText = strSQLInsert;
                                        cmd.ExecuteNonQuery();

                                        strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                    }

                                }
                            }
                        }

                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();

                        //tx.Commit();
                    }
                    //else
                    //{

                    //    tx.Commit();
                    //}
                    //如果emp_class有資料，但是班別為空，則給默認班別502
                    strSQLUpdate = "update swipecard.emp_class set class_no='502' where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) and class_no=''";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();


                    tx.Commit();

                    WriteLog("-->Update152_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert152_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert152_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            else if (strBG == "ASBG")
            {
                #region ASBG
                //60.112

                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 153ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
                WriteLog("-->Query 153ClassSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                //String sql = "SELECT DISTINCT 編號  FROM V_RYBZ_TX"; // 查询语句
                ////String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                //if (DT2 == null)
                //{
                //    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                //    return;
                //}
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;
                //}
                //string strEMP = "";
                //foreach (DataRow dr in DT2.Rows)
                //{
                //    strEMP += " EMP_CD='" + dr[0].ToString() + "' OR";
                //}
                //strEMP = strEMP.Substring(0, strEMP.Length - 2);
                //WriteLog("-->SelectSqlServerEmpSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, START_DATE, CLASS_CD from PUB.EMPCLASSR ";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEmpClassSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEmpClassSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.78.153;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {

                        TimeSpan ts = DateTime.Now.Date - Convert.ToDateTime(dr[1].ToString());
                        int d = ts.Days;
                        string strClass = dr[2].ToString().Substring(0, dr[2].ToString().Length - 1);
                        List<string> list = strClass.Split(new string[] { "," }, StringSplitOptions.None).ToList<string>();
                        for (int i = 0; i < 2; i++)
                        {
                            if ((d + i) >= list.Count) continue;
                            if (!ClassArray3.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + list[d + i]))
                            {
                                if (ClassArray2.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                                {
                                    string sa = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    if (i == 0)
                                    { strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where (class_no ='' or class_no='502') and ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'"; }
                                    else
                                    { strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'"; }
                                    cmd.CommandText = strSQLUpdate;
                                    int it = cmd.ExecuteNonQuery();
                                    if (it > 0) UpdateSumk++;
                                    //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                                }
                                else
                                {
                                    ++InsertSumk;
                                    string sb = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                         + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                         + sb + "',"
                                                         + "curdate()" + "),";

                                    if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                                    {
                                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                        cmd.CommandText = strSQLInsert;
                                        cmd.ExecuteNonQuery();

                                        strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                    }

                                }
                            }
                        }

                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();

                        //tx.Commit();
                    }
                    //else
                    //{

                    //    tx.Commit();
                    //}
                    //如果emp_class有資料，但是班別為空，則給默認班別4
                    //如果emp_class有資料，但是班別為空，則給默認班別502 2017-10-18
                    strSQLUpdate = "update swipecard.emp_class set class_no='502' where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) and class_no=''";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();


                    tx.Commit();

                    WriteLog("-->Update153_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert153_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert153_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }

        }

        private void InsertMySql_empClass0330(string strBG)
        {
            //今天和明天的都会Update和Insert
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) order by ID, emp_date";
            //string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where  emp_date=date_add(CURDATE(),interval 1 day) order by ID, emp_date";
            if (strBG == "CSBG")
            {
                #region CSBG
                //60.111

                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 152ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
                WriteLog("-->Query 152ClassSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                //String sql = "SELECT DISTINCT 編號  FROM V_RYBZ_TX"; // 查询语句
                ////String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                //if (DT2 == null)
                //{
                //    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                //    return;
                //}
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;
                //}
                //string strEMP = "";
                //foreach (DataRow dr in DT2.Rows)
                //{
                //    strEMP += " EMP_CD='" + dr[0].ToString() + "' OR";
                //}
                //strEMP = strEMP.Substring(0, strEMP.Length - 2);
                //WriteLog("-->SelectSqlServerEmpSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, START_DATE, CLASS_CD from PUB.EMPCLASSR ";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEmpClassSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEmpClassSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.78.152;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 30;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {
                        //WriteLog(dr[0].ToString() + '\r', 1);
                        TimeSpan ts = DateTime.Now.Date - Convert.ToDateTime(dr[1].ToString());
                        int d = ts.Days;
                        string strClass = dr[2].ToString().Substring(0, dr[2].ToString().Length - 1);
                        List<string> list = strClass.Split(new string[] { "," }, StringSplitOptions.None).ToList<string>();
                        for (int i = 0; i < 2; i++)
                        {
                            if ((d + i) >= list.Count) continue;
                            if (!ClassArray3.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + list[d + i]))
                            {
                                if (ClassArray2.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                                {
                                    string sa = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'";
                                    cmd.CommandText = strSQLUpdate;
                                    cmd.ExecuteNonQuery();
                                    UpdateSumk++;
                                    //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                                }
                                else
                                {

                                    ++InsertSumk;
                                    string sb = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                         + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                         + sb + "',"
                                                         + "curdate()" + "),";

                                    if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                                    {
                                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                        cmd.CommandText = strSQLInsert;
                                        cmd.ExecuteNonQuery();

                                        strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                    }

                                }
                            }
                        }

                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        //tx.Commit();
                    }
                    //else
                    //{
                    //    tx.Commit();
                    //}
                    //如果emp_class有資料，但是班別為空，則給默認班別4
                    //如果emp_class有資料，但是班別為空，則給默認班別502
                    strSQLUpdate = "update swipecard.emp_class set class_no='502' where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) and class_no=''";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();
                    tx.Commit();




                    WriteLog("-->Update152_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert152_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert152_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            else if (strBG == "ASBG")
            {
                #region ASBG
                //60.112

                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 153ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
                WriteLog("-->Query 153ClassSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                //String sql = "SELECT DISTINCT 編號  FROM V_RYBZ_TX"; // 查询语句
                ////String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                //if (DT2 == null)
                //{
                //    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                //    return;
                //}
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;
                //}
                //string strEMP = "";
                //foreach (DataRow dr in DT2.Rows)
                //{
                //    strEMP += " EMP_CD='" + dr[0].ToString() + "' OR";
                //}
                //strEMP = strEMP.Substring(0, strEMP.Length - 2);
                //WriteLog("-->SelectSqlServerEmpSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, START_DATE, CLASS_CD from PUB.EMPCLASSR ";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEmpClassSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEmpClassSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.78.153;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {

                        TimeSpan ts = DateTime.Now.Date - Convert.ToDateTime(dr[1].ToString());
                        int d = ts.Days;
                        string strClass = dr[2].ToString().Substring(0, dr[2].ToString().Length - 1);
                        List<string> list = strClass.Split(new string[] { "," }, StringSplitOptions.None).ToList<string>();
                        for (int i = 0; i < 2; i++)
                        {
                            if ((d + i) >= list.Count) continue;
                            if (!ClassArray3.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + list[d + i]))
                            {
                                if (ClassArray2.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                                {
                                    string sa = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'";
                                    cmd.CommandText = strSQLUpdate;
                                    cmd.ExecuteNonQuery();
                                    UpdateSumk++;
                                    //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                                }
                                else
                                {
                                    ++InsertSumk;
                                    string sb = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                         + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                         + sb + "',"
                                                         + "curdate()" + "),";

                                    if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                                    {
                                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                        cmd.CommandText = strSQLInsert;
                                        cmd.ExecuteNonQuery();

                                        strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                    }

                                }
                            }
                        }

                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();

                        //tx.Commit();
                    }
                    //else
                    //{

                    //    tx.Commit();
                    //}
                    //如果emp_class有資料，但是班別為空，則給默認班別4
                    //如果emp_class有資料，但是班別為空，則給默認班別502
                    strSQLUpdate = "update swipecard.emp_class set class_no='502' where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) and class_no=''";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();
                    tx.Commit();



                    WriteLog("-->Update153_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert153_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert153_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            

        }
        private void InsertMySql_empClass1(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            if (strBG == "CSBG")
            {
                //60.111
                string strEMP = "";
                string strSqlMySqlEmp = "select distinct ID from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 152EmpError" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog("-->Query 152EmpSUM:0" + '\r', 2);
                    return;
                }

                WriteLog("-->Query 152EmpSum:" + DT1.Rows.Count.ToString() + '\r', 1);
                foreach (DataRow dr in DT1.Rows)
                {
                    strEMP += " zgbh='" + dr[0].ToString() + "' OR";
                }
                strEMP = strEMP.Substring(0, strEMP.Length - 2);


                String sql = "select zgbh ,rq,bcbh from mf_class where rq>=CONVERT(varchar(100), GETDATE()-3, 111) and rq<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }
                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=localhost;userid=root;password=foxlink;database=swipecard;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                cmd.CommandTimeout = 0;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;
                int i = 1;
                try
                {
                    string strSQL = "delete from swipecard.emp_class where emp_date>=curdate()-3 and emp_date<curdate()+2 and ID in (select distinct ID from swipecard.testemployee)";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    strSQL = "insert into swipecard.emp_class VALUES ";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        string sa = dr[2].ToString() == "\\" ? "\\\\" : dr[2].ToString();
                        strSQL += " ('" + dr[0].ToString() + "','"
                                        + Convert.ToDateTime(dr[1].ToString()).ToString("yyyy/MM/dd") + "','"
                                        + sa + "',"
                                        + "curdate()" + "),";
                        if (i > 0 && i % 3000 == 0)
                        {
                            strSQL = strSQL.Substring(0, strSQL.Length - 1);
                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();

                            strSQL = "insert into swipecard.emp_class  VALUES ";
                        }
                        i++;
                    }
                    if (DT2.Rows.Count % 3000 != 0)
                    {
                        strSQL = strSQL.Substring(0, strSQL.Length - 1);
                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //throw new ArgumentOutOfRangeException("shoudong");
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }
                    WriteLog("-->insert152_EmpClassOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert152_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                string strEMP = "";
                string strSqlMySqlEmp = "select distinct ID from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 153EmpError" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog("-->Query 153EmpSUM:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query 153EmpSum:" + DT1.Rows.Count.ToString() + '\r', 1);
                foreach (DataRow dr in DT1.Rows)
                {
                    strEMP += " zgbh='" + dr[0].ToString() + "' OR";
                }
                strEMP = strEMP.Substring(0, strEMP.Length - 2);


                String sql = "select zgbh ,rq,bcbh from mf_class where rq>=CONVERT(varchar(100), GETDATE()-3, 111) and rq<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where ygbh!='' and ygbh is not null "; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }
                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=localhost;userid=root;password=foxlink;database=swipecard;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                cmd.CommandTimeout = 0;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;
                int i = 1;
                try
                {
                    string strSQL = "delete from swipecard.emp_class where emp_date>=curdate()-3 and emp_date<curdate()+2 and ID in (select distinct ID from swipecard.testemployee)";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    strSQL = "insert into swipecard.emp_class VALUES ";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        string sa = dr[2].ToString() == "\\" ? "\\\\" : dr[2].ToString();
                        strSQL += " ('" + dr[0].ToString() + "','"
                                        + Convert.ToDateTime(dr[1].ToString()).ToString("yyyy/MM/dd") + "','"
                                        + sa + "',"
                                        + "curdate()" + "),";
                        if (i > 0 && i % 3000 == 0)
                        {
                            strSQL = strSQL.Substring(0, strSQL.Length - 1);
                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();

                            strSQL = "insert into swipecard.emp_class  VALUES ";
                        }
                        i++;
                    }
                    if (DT2.Rows.Count % 3000 != 0)
                    {
                        strSQL = strSQL.Substring(0, strSQL.Length - 1);
                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //throw new ArgumentOutOfRangeException("shoudong");
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }
                    WriteLog("-->insert153_EmpClassOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert152_EmpClassOK,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }


            }
        }

        private void txtReadPath_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtBackupPath_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }

}
