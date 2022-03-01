using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using NLog;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace TrueVmsService
{
    public partial class Service1 : ServiceBase
    {

        System.Timers.Timer updateOTPTimer;
        System.Timers.Timer sendEmail;
        Logger log = NLog.LogManager.GetCurrentClassLogger();
        private static bool IS_DEBUG = Convert.ToBoolean(ConfigurationManager.AppSettings["isDebug"]);


        static string DATABASE_USERNAME = ConfigurationManager.AppSettings["dbUser"];
        static string DATABASE_PASSWORD = ConfigurationManager.AppSettings["dbPassword"];
        static string DATABASE_NAME = ConfigurationManager.AppSettings["dbName"];
        static string DATABASE_IP = ConfigurationManager.AppSettings["dbIp"];
        static string WEB_SERVER = ConfigurationManager.AppSettings["webserver"];


        static int WARNING_LOOP1 = Convert.ToInt32(ConfigurationManager.AppSettings["warningLoop1"]);
        static int WARNING_LOOP2 = Convert.ToInt32(ConfigurationManager.AppSettings["warningLoop2"]);
        static int WARNING_LOOP3 = Convert.ToInt32(ConfigurationManager.AppSettings["warningLoop3"]);
        static int CLEAR_OTP = Convert.ToInt32(ConfigurationManager.AppSettings["clearOtp"]);

        


        static int SendingActivateEmail = Convert.ToInt32(ConfigurationManager.AppSettings["SendingActivateEmail"]);
        static int minAppUserId = Convert.ToInt32(ConfigurationManager.AppSettings["minAppUserId"]);
        static int maxAppUserId = Convert.ToInt32(ConfigurationManager.AppSettings["maxAppUserId"]);





        private static string CONNNECTION_STRING = "";

        public Service1()
        {
            InitializeComponent();
            initialNLog();

            CONNNECTION_STRING = "server=" + DATABASE_IP + ";uid=" + DATABASE_USERNAME + ";" +
                "pwd=" + DATABASE_PASSWORD + ";database=" + DATABASE_NAME + ";";
          
           
        }



        private void initialNLog()
        {
            var config = new NLog.Config.LoggingConfiguration();

            // Targets where to log to: File and Console
            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = @"c:\temp\TVSlog" + DateTime.Now.ToString("ddMMyyy") + ".txt" };
            var logconsole = new NLog.Targets.ConsoleTarget("logconsole");

            // Rules for mapping loggers to targets            
            config.AddRule(LogLevel.Info, LogLevel.Fatal, logfile);
            //config.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);

            // Apply config           
            NLog.LogManager.Configuration = config;
        }

        protected override void OnStart(string[] args)
        {
            try
            {

                log.Info("*************************************************");
                log.Info("True VMS Data Collector Service Start version 1.0");

                insertOTP();

                cleaeOldOtp();
                clearExpireProject();
                clearExpireWorkpermit();
                clearExpireStaff();
                sendingReviewProject();
                sendingReviewWorkpermit();

                sendingEmailLoop3();
                sendingEmailLoop2();
                sendingEmailLoop1();

                if(SendingActivateEmail == 1)
                    sendingActivateEmail();
            }
            catch { }

            updateOTPTimer = new System.Timers.Timer();
            updateOTPTimer.Elapsed += new System.Timers.ElapsedEventHandler(updateOTP);
            updateOTPTimer.Interval = 900000;
            updateOTPTimer.Enabled = true;
            updateOTPTimer.AutoReset = true;
            updateOTPTimer.Start();

            sendEmail = new System.Timers.Timer();
            sendEmail.Elapsed += new System.Timers.ElapsedEventHandler(sendWarningEmail1);
            sendEmail.Interval = getNextMidnight();
            sendEmail.Enabled = true;
            sendEmail.AutoReset = true;
            sendEmail.Start();
        }

        private double getNextMidnight()
        {
            TimeSpan ts = DateTime.Today.AddDays(1) - DateTime.Now;

            //log.Info("DateTime.Today.AddDays(1) :"+ DateTime.Today.AddDays(1));
            //log.Info("ts :" + ts.TotalHours + " "+ts.TotalMinutes);
            //log.Info("ts TotalMilliseconds :" + ts.TotalMilliseconds);
            return ts.TotalMilliseconds;
        }



        protected override void OnStop()
        {
            try
            {
                if (updateOTPTimer != null)
                {
                    updateOTPTimer.Stop();
                }

                if (sendEmail != null)
                {
                    sendEmail.Stop();
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
            }



            log.Info("Service Stop");
        }

        internal void resetTimeInterval()
        {
            sendEmail.Stop();
            sendEmail.Interval = 86400000;
            sendEmail.Start();
        }

        public TokenModel getToken(string email)
        {
            try
            {
                log.Info("getTokenResetPass :" + email + " " + WEB_SERVER + "/api/Account/ResetPassword");

                var client = new RestClient(WEB_SERVER + "/api/Account/ResetPassword");



                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Content-Type", "application/json");
                var body = @"{
" + "\n" +
                @"    ""Email"": """ + email + @"""
" + "\n" +
                @"}";
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                log.Info(response.Content);



                TokenModel m = JsonConvert.DeserializeObject<TokenModel>(response.Content);

                return m;
            }
            catch(Exception ex)
            {

                log.Info("4");

                log.Info(ex.ToString());
                throw ex;
            }
        }

        private void sendingEmailLoop1()
        {

            MySql.Data.MySqlClient.MySqlConnection conn = null;
            MySql.Data.MySqlClient.MySqlConnection conn2 = null;
            string myConnectionString;

            myConnectionString = CONNNECTION_STRING;

            try
            {

                //log.Info("Connect to database...");

                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();


                conn2 = new MySql.Data.MySqlClient.MySqlConnection();
                conn2.ConnectionString = myConnectionString;
                conn2.Open();


                log.Info("Geting user expired..");

                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conn;


                MySqlCommand cmd2 = new MySqlCommand();
                cmd2.Connection = conn2;



                string cmdSelect = "SELECT USER_ID,CUST_UPDATER_EMAIL,CUST_UPDATER_NAME,WARNING_COUNT,WARNING_EMAIL_DATE " +
                    "FROM cust_updater " +
                    "where TIMESTAMPDIFF(DAY, NEXT_PASSWORD_CHANGE, NOW()) = "+ WARNING_LOOP1 + " and STATUS=13 and (WARNING_COUNT is null or WARNING_COUNT = 0)";
                log.Info(cmdSelect);
                cmd.CommandText = cmdSelect;
                MySqlDataReader reader =   cmd.ExecuteReader();



                while (reader.Read())
                {
                    string email = reader.GetString("CUST_UPDATER_EMAIL");
                    int updaterId = reader.GetInt32("USER_ID");

                    log.Info("Warning password expire : " + email);

                    TokenModel t = getToken(email);

                    String body = MailManager.getBody(Convert.ToString(updaterId),t.token);
                    MailManager.QuickSend(email, "[Warning #1] Account temporarily susspended", body);

                    log.Info("Update warning_count updater id :"+updaterId);

                    string cmdUpdate = "update cust_updater set WARNING_COUNT = 1 , WARNING_EMAIL_DATE = now() where user_id = "+ updaterId;
                    //log.Info(cmdUpdate);
                    cmd2.CommandText = cmdUpdate;
                    int rowEffect = cmd2.ExecuteNonQuery();

                    log.Info("Update warning_count completed :" + rowEffect);
                }                
                reader.Close();

                //SELECT USER_ID,CUST_UPDATER_EMAIL,CUST_UPDATER_NAME,WARNING_COUNT,WARNING_EMAIL_DATE FROM htvms.cust_updater where NEXT_PASSWORD_CHANGE <= now() and datediff(now(),WARNING_EMAIL_DATE) = 3 and STATUS=13 and  WARNING_COUNT = 1;

            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }

                if (conn2 != null)
                {
                    conn2.Close();
                    conn2.Dispose();
                }
            }
        }


        private void sendingEmailLoop2()
        {

            MySql.Data.MySqlClient.MySqlConnection conn = null;
            MySql.Data.MySqlClient.MySqlConnection conn2 = null;
            string myConnectionString;

            myConnectionString = CONNNECTION_STRING;

            try
            {

                //log.Info("Connect to database...");

                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();


                conn2 = new MySql.Data.MySqlClient.MySqlConnection();
                conn2.ConnectionString = myConnectionString;
                conn2.Open();


                log.Info("Geting user expired..");

                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conn;


                MySqlCommand cmd2 = new MySqlCommand();
                cmd2.Connection = conn2;



                string cmdSelect = "SELECT USER_ID,CUST_UPDATER_EMAIL,CUST_UPDATER_NAME,WARNING_COUNT,WARNING_EMAIL_DATE " +
                    "FROM cust_updater " +
                    "where TIMESTAMPDIFF(DAY, NEXT_PASSWORD_CHANGE, NOW()) = "+WARNING_LOOP2+" and STATUS=13 and WARNING_COUNT = 1";
                log.Info("Loop 2 :" + cmdSelect);
                cmd.CommandText = cmdSelect;
                MySqlDataReader reader = cmd.ExecuteReader();



                while (reader.Read())
                {
                    string email = reader.GetString("CUST_UPDATER_EMAIL");
                    int updaterId = reader.GetInt32("USER_ID");

                    log.Info("Warning password expire : " + email);

                    TokenModel t = getToken(email);

                    String body = MailManager.getBody(Convert.ToString(updaterId),t.token);
                    MailManager.QuickSend(email, "[Warning #2] Account temporarily susspended", body);

                    log.Info("Update warning_count updater id :" + updaterId);

                    string cmdUpdate = "update cust_updater set WARNING_COUNT = 2 , WARNING_EMAIL_DATE = now() where user_id = " + updaterId;
                    //log.Info(cmdUpdate);
                    cmd2.CommandText = cmdUpdate;
                    int rowEffect = cmd2.ExecuteNonQuery();

                    log.Info("Update warning_count completed :" + rowEffect);
                }
                reader.Close();

                //SELECT USER_ID,CUST_UPDATER_EMAIL,CUST_UPDATER_NAME,WARNING_COUNT,WARNING_EMAIL_DATE FROM htvms.cust_updater where NEXT_PASSWORD_CHANGE <= now() and datediff(now(),WARNING_EMAIL_DATE) = 3 and STATUS=13 and  WARNING_COUNT = 1;

            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }

                if (conn2 != null)
                {
                    conn2.Close();
                    conn2.Dispose();
                }
            }


        }


        private void sendingEmailLoop3()
        {

            MySql.Data.MySqlClient.MySqlConnection conn = null;
            MySql.Data.MySqlClient.MySqlConnection conn2 = null;
            string myConnectionString;

            myConnectionString = CONNNECTION_STRING;

            try
            {

                //log.Info("Connect to database...");

                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();


                conn2 = new MySql.Data.MySqlClient.MySqlConnection();
                conn2.ConnectionString = myConnectionString;
                conn2.Open();


                log.Info("Geting user expired..");

                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conn;


                MySqlCommand cmd2 = new MySqlCommand();
                cmd2.Connection = conn2;



                string cmdSelect = "SELECT USER_ID,CUST_UPDATER_EMAIL,CUST_UPDATER_NAME,WARNING_COUNT,WARNING_EMAIL_DATE " +
                    "FROM cust_updater " +
                    "where TIMESTAMPDIFF(DAY, NEXT_PASSWORD_CHANGE, NOW()) = "+ WARNING_LOOP3 + " and STATUS=13 and WARNING_COUNT = 2";
                log.Info("Loop 3 : " + cmdSelect);
                cmd.CommandText = cmdSelect;
                MySqlDataReader reader = cmd.ExecuteReader();



                while (reader.Read())
                {
                    string email = reader.GetString("CUST_UPDATER_EMAIL");
                    int updaterId = reader.GetInt32("USER_ID");

                    log.Info("Warning password expire : " + email);
                    TokenModel t = getToken(email);
                    log.Info("get token : " + t.ToString());

                    String body = MailManager.getBody(Convert.ToString(updaterId),t.token);
                    MailManager.QuickSend(email, "[Warning #3] Account temporarily susspended", body);

                    log.Info("Update warning_count updater id :" + updaterId);

                    string cmdUpdate = "update cust_updater set WARNING_COUNT = 3 , WARNING_EMAIL_DATE = now(), STATUS=15 where user_id = " + updaterId;
                    //log.Info(cmdUpdate);
                    cmd2.CommandText = cmdUpdate;
                    int rowEffect = cmd2.ExecuteNonQuery();

                    log.Info("Update warning_count completed :" + rowEffect);
                }
                reader.Close();

                //SELECT USER_ID,CUST_UPDATER_EMAIL,CUST_UPDATER_NAME,WARNING_COUNT,WARNING_EMAIL_DATE FROM htvms.cust_updater where NEXT_PASSWORD_CHANGE <= now() and datediff(now(),WARNING_EMAIL_DATE) = 3 and STATUS=13 and  WARNING_COUNT = 1;

            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }

                if (conn2 != null)
                {
                    conn2.Close();
                    conn2.Dispose();
                }
            }


        }



        public void sendWarningEmail1(object sender, System.Timers.ElapsedEventArgs args)
        {
            sendingReviewProject();
            sendingReviewWorkpermit();

            sendingEmailLoop3();
            sendingEmailLoop2();
            sendingEmailLoop1();
            resetTimeInterval();
        }

        private void insertOTP()
        {
            MySql.Data.MySqlClient.MySqlConnection conn = null;
            string myConnectionString = CONNNECTION_STRING;

            //myConnectionString = "server=49.0.82.124;uid=vmsadmin;" +
            //    "pwd=p@ssw0rd1234;database=htvms";


            try
            {

                //log.Info("Connect to database..."+ myConnectionString);

                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();


                //log.Info("Connection complete");

                MySqlCommand cmd = new MySqlCommand();
               

                cmd.Connection = conn;
                
                string randomOTP  = new Random().Next(0, 9999).ToString("D4");
                
                if(IS_DEBUG)
                    randomOTP = "2327";

                string cmdInsert = "INSERT INTO `vms_override_otp` ( `OVERRIDE_OTP`, `OTP_GENERATE_DATETIME`,`OTP_LIFETIME_SECOND`) " +
                "VALUES ('"+ randomOTP + "', '"+ DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss")+ "', 900000)";
                //log.Info(cmdInsert);
                cmd.CommandText = cmdInsert;
                cmd.ExecuteNonQuery();

                log.Info("Insert Otp Complete :"+ randomOTP);


            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }



        private void cleaeOldOtp()
        {
            MySql.Data.MySqlClient.MySqlConnection conn = null;
            string myConnectionString = CONNNECTION_STRING;
            try
            {


                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();

                MySqlCommand cmd = new MySqlCommand();


                cmd.Connection = conn;

                string cmdUpdate = "update vms_qr_code set status = 37 where status != 37 and TIMESTAMPDIFF(MINUTE, qr_code_generate_datetime, NOW())> "+ CLEAR_OTP;
                //log.Info(cmdUpdate);
                cmd.CommandText = cmdUpdate;
                int row = cmd.ExecuteNonQuery();

                log.Info("Clear old OTP Complete :"+row);


            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }


        private void clearExpireProject()
        {
            MySql.Data.MySqlClient.MySqlConnection conn = null;
            string myConnectionString = CONNNECTION_STRING;
            try
            {


                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();

                MySqlCommand cmd = new MySqlCommand();


                cmd.Connection = conn;

                string cmdUpdate = "update cust_project set status = 12 where status != 12 and  TIMESTAMPDIFF(MINUTE, cust_project_end, NOW()) >= 0";
                //log.Info(cmdUpdate);
                cmd.CommandText = cmdUpdate;
                int row = cmd.ExecuteNonQuery();

                log.Info("Clear expire project Complete :" + row);
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }


        private void clearExpireWorkpermit()
        {
            MySql.Data.MySqlClient.MySqlConnection conn = null;
            string myConnectionString = CONNNECTION_STRING;
            try
            {


                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();

                MySqlCommand cmd = new MySqlCommand();


                cmd.Connection = conn;

                string cmdUpdate = "update workpermit set status = 2 where status = 1 and TIMESTAMPDIFF(minute, work_end_datetime, NOW()) >= 0";
                //log.Info(cmdUpdate);
                cmd.CommandText = cmdUpdate;
                int row = cmd.ExecuteNonQuery();

                log.Info("Clear expire workpermit complete :" + row);
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }


        private void clearExpireStaff()
        {
            MySql.Data.MySqlClient.MySqlConnection conn = null;
            string myConnectionString = CONNNECTION_STRING;
            try
            {


                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();

                MySqlCommand cmd = new MySqlCommand();


                cmd.Connection = conn;

                string cmdUpdate = "update cust_staff set status = 35 where status = 34 and TIMESTAMPDIFF(minute, next_review_date, NOW()) >= 0";
                //log.Info(cmdUpdate);
                cmd.CommandText = cmdUpdate;
                int row = cmd.ExecuteNonQuery();

                log.Info("Clear expire staff complete :" + row);
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }


        string warninEmail = ConfigurationManager.AppSettings["warninEmail"];

        private void sendingReviewProject()
        {

            MySql.Data.MySqlClient.MySqlConnection conn = null;
            //MySql.Data.MySqlClient.MySqlConnection conn2 = null;
            string myConnectionString;

            myConnectionString = CONNNECTION_STRING;

            try
            {

                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();


                //conn2 = new MySql.Data.MySqlClient.MySqlConnection();
                //conn2.ConnectionString = myConnectionString;
                //conn2.Open();


                log.Info("Geting project review..");

                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conn;


                //MySqlCommand cmd2 = new MySqlCommand();
                //cmd2.Connection = conn2;



                string cmdSelect = "SELECT CUSTOMER_CODE,CUST_PROJECT_NAME,NEXT_REVIEW_DATE FROM cust_project where status = 9 and  TIMESTAMPDIFF(day, next_review_date, NOW()) >= -15";
                log.Info(cmdSelect);
                cmd.CommandText = cmdSelect;
                MySqlDataReader reader = cmd.ExecuteReader();



                while (reader.Read())
                {
                    string CUSTOMER_CODE = reader.GetString("CUSTOMER_CODE");
                    string CUST_PROJECT_NAME = reader.GetString("CUST_PROJECT_NAME");
                    string NEXT_REVIEW_DATE = reader.GetString("NEXT_REVIEW_DATE");

                    log.Info("Warning project expire to : " + warninEmail);

                    String body = MailManager.getProjectExpireBody(CUSTOMER_CODE,CUST_PROJECT_NAME,NEXT_REVIEW_DATE);
                    MailManager.QuickSend(warninEmail, "[Warning] Project expire", body);
                }
                reader.Close();

                
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }

                //if (conn2 != null)
                //{
                //    conn2.Close();
                //    conn2.Dispose();
                //}
            }
        }


        private void sendingReviewWorkpermit()
        {

            MySql.Data.MySqlClient.MySqlConnection conn = null;
            //MySql.Data.MySqlClient.MySqlConnection conn2 = null;
            string myConnectionString;

            myConnectionString = CONNNECTION_STRING;

            try
            {

                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();


                //conn2 = new MySql.Data.MySqlClient.MySqlConnection();
                //conn2.ConnectionString = myConnectionString;
                //conn2.Open();


                log.Info("Geting workpermit expire..");

                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conn;


                //MySqlCommand cmd2 = new MySqlCommand();
                //cmd2.Connection = conn2;



                //string cmdSelect = "SELECT WORKPERMIT_CODE,WORK_END_DATETIME,cust_project.CUST_PROJECT_NAME FROM workpermit inner join cust_project on workpermit.CUST_PROJECT_ID = cust_project.CUST_PROJECT_ID  where workpermit.status = 1 and TIMESTAMPDIFF(day, work_end_datetime, NOW()) >= -1";
                string cmdSelect = "SELECT WORKPERMIT_CODE,WORK_END_DATETIME,CUST_PROJECT_NAME,UPDATER_EMAIL,CREATED_BY_EMAIL,CUST_COMPANY_NAME FROM rpt_workpermit_vw where rpt_workpermit_vw.status = 1 and TIMESTAMPDIFF(day, work_end_datetime, NOW()) >= -1";
                log.Info(cmdSelect);
                cmd.CommandText = cmdSelect;
                MySqlDataReader reader = cmd.ExecuteReader();



                while (reader.Read())
                {
                    string WORKPERMIT_CODE = reader.GetString("WORKPERMIT_CODE");
                    string WORK_END_DATETIME = reader.GetString("WORK_END_DATETIME");
                    string CUST_PROJECT_NAME = reader.GetString("CUST_PROJECT_NAME");
                    string CUST_NAME = reader.GetString("CUST_COMPANY_NAME");


                    //string UPDATER_EMAIL = reader.GetString("UPDATER_EMAIL");
                    string CREATED_BY_EMAIL = reader.GetString("CREATED_BY_EMAIL");                    

                    log.Info("Warning workpermit expire to : " + CREATED_BY_EMAIL);

                    String body = MailManager.getWorkpermitExpireBody(WORKPERMIT_CODE, WORK_END_DATETIME, CUST_PROJECT_NAME, CUST_NAME);
                    MailManager.QuickSend(CREATED_BY_EMAIL, "True IDC Site entry: Work permit expiry reminder "+ WORKPERMIT_CODE, body);
                }
                reader.Close();


            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }

                //if (conn2 != null)
                //{
                //    conn2.Close();
                //    conn2.Dispose();
                //}
            }
        }


        private void sendActivateEmail(string email)
        {
            try
            {
                log.Info("sendActivateEmail :" + email + " " + WEB_SERVER + "/api/Account/ResendActivateEmail");

                var client = new RestClient(WEB_SERVER + "/api/Account/ResendActivateEmail");

                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Cookie", "__cfruid=33f201b2a608fa5b25603ca2c846e653de2ab9dd-1640097552");
                var body = @"{
    " + "\n" +
                    @"    ""email"": """ + email + @"""
    " + "\n" +
                    @"}";
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                log.Info(response.Content);
            }
            catch (Exception ex)
            {
                log.Info(ex.ToString());
                throw ex;
            }
        }


        private void sendingActivateEmail()
        {

            MySql.Data.MySqlClient.MySqlConnection conn = null;
            string myConnectionString;

            myConnectionString = CONNNECTION_STRING;

            try
            {

                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();

                log.Info("Geting account for activate email..");

                MySqlCommand cmd = new MySqlCommand();
                cmd.Connection = conn;

                 string cmdSelect = "SELECT Email,UserName,Id,AppUserId FROM AspNetUsers where EmailConfirmed = 0 and AppUserId > "+ minAppUserId + " and AppUserId < "+ maxAppUserId + " order by AppUserId asc ";
                log.Info(cmdSelect);
                cmd.CommandText = cmdSelect;
                MySqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string Email = reader.GetString("Email");
                    string UserName = reader.GetString("UserName");
                    string Id = reader.GetString("Id");


                    string AppUserId = reader.GetString("AppUserId");

                    log.Info("send activate email to : " + Email + " AppUserId :" + AppUserId);

                    sendActivateEmail(Email);
                }
                reader.Close();

            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                log.Error(ex.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }

                //if (conn2 != null)
                //{
                //    conn2.Close();
                //    conn2.Dispose();
                //}
            }
        }



        public void updateOTP(object sender, System.Timers.ElapsedEventArgs args)
        {         
            try
            {
                insertOTP();

                cleaeOldOtp();

                clearExpireProject();

                clearExpireWorkpermit();

                clearExpireStaff();
            }
            catch (Exception ex){
                log.Error(ex.ToString());
            }
        }
    }
}
