using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace TrueVmsService
{
    public class MailManager
    {


        private static NLog.Logger log = NLog.LogManager.GetCurrentClassLogger();

        

        public static string getBody(string updaterID, string linkToken)
        {

            string webserver = ConfigurationManager.AppSettings["webserver"];

            string body = string.Format(@"
                    <table width=""100%"" cellspacing=""0"" cellpadding=""0"">
                      <tr>
                          <td style='text-align:center; vertical-align:middle'>
                            <img src=""{0}/dist/img/trueidclogo.png"" style=""display: block; margin: 0px auto"" width=""136"" height=""46"" />
                          </td>
                      </tr>
                    </table>
                    ", webserver);
            body += "<br />";
            body += "<hr />";
            body += "<br />";
            body += "<br />";


            body += "Dear Customer, <br /> " +
                          "<p>Your account is temporarily susspended, Please proceed to change your </p>  " +
                          "<p>password by click link below. in order to protect and secure your account all </p> " +
                          "<p>account password must be change every 90 days. </p>  <br /> " +

                          "<p><left><a href=\""+ linkToken + "\"><img style=\"width: 180px\" src=\"" + webserver+"/dist/img/EmailBtnChangePassword.png\"></a></left><p> <br /> " +

                          "Sincerely,  <br /> " +
                          "Your True IDC Account Administrator ";



            return body;
        }


        public static string getProjectExpireBody(string projectCode, string projectname, string nextReview)
        {

            string webserver = ConfigurationManager.AppSettings["webserver"];

            string body = string.Format(@"
                    <table width=""100%"" cellspacing=""0"" cellpadding=""0"">
                      <tr>
                          <td style='text-align:center; vertical-align:middle'>
                            <img src=""{0}/dist/img/trueidclogo.png"" style=""display: block; margin: 0px auto"" width=""136"" height=""46"" />
                          </td>
                      </tr>
                    </table>
                    ", webserver);
            body += "<br />";
            body += "<hr />";
            body += "<br />";
            body += "<br />";


            body += "Dear Administrator, <br /> " +
                          "<p>Your project is temporarily susspended, Please proceed review project. </p>  " +
                          "<p>Project must review every 90 days. </p>  <br /> " +

                          "<p><left> Project : "+ projectCode + "-" +projectname+ " next review : "+nextReview+"</left><p> <br /> " +

                          "Sincerely,  <br /> " +
                          "Your True IDC Account Administrator ";



            return body;
        }


        public static string getWorkpermitExpireBody(string workpermit,  string endDate, string project)
        {

            string webserver = ConfigurationManager.AppSettings["webserver"];

            string body = string.Format(@"
                    <table width=""100%"" cellspacing=""0"" cellpadding=""0"">
                      <tr>
                          <td style='text-align:center; vertical-align:middle'>
                            <img src=""{0}/dist/img/trueidclogo.png"" style=""display: block; margin: 0px auto"" width=""136"" height=""46"" />
                          </td>
                      </tr>
                    </table>
                    ", webserver);
            body += "<br />";
            body += "<hr />";
            body += "<br />";
            body += "<br />";


            body += "Dear Administrator, <br /> " +
                          "<p>Your workpermit time is going out. </p>  " +

                          "<p><left> workpermit : " + workpermit + "(project "+project+ ") time is going out at : " + endDate + "</left><p> <br /> " +

                          "Sincerely,  <br /> " +
                          "Your True IDC Account Administrator ";



            return body;
        }


        public static void QuickSend(
            string ToEmail,
            string Subject,
            string Body)
        {
            try
            {
                string fromName = ConfigurationManager.AppSettings["mailFrom"];
                MailMessage message = new MailMessage(fromName, ToEmail,
                   Subject,
                   Body);

                log.Info("Send mail to " + ToEmail);

                //message.To.Add(ToEmail);
                message.Bcc.Add("akkachat.t@hitop.co.th");

                string smtpServerName = ConfigurationManager.AppSettings["mailSMTPServerName"];

                SmtpClient client = new SmtpClient(smtpServerName);

                // Add credentials if the SMTP server requires them.
                string username = ConfigurationManager.AppSettings["mailSMTPUserName"];
                string password = ConfigurationManager.AppSettings["mailSMTPUserPassword"];
                string port = ConfigurationManager.AppSettings["mailSMTPServerPort"];
                string ssl = ConfigurationManager.AppSettings["mailSMTPSSL"];
                if (username != null)
                {
                    client.Credentials = new NetworkCredential(username, password);

                }
                if ("true".Equals(ssl))
                {
                    client.EnableSsl = true;
                }
                if (port != "0")
                {
                    client.Port = Convert.ToInt32(port);
                }

                try
                {
                    //log.Info("client.Credentials :" + client.Credentials.ToString());
                    //log.Info("client.EnableSsl :" + client.EnableSsl);
                    //log.Info("client.Port :" + client.Port);

                    message.IsBodyHtml = true;
                    client.Send(message);
                }
                catch (Exception ex)
                {
                    log.Error(ex.ToString());
                    throw ex;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }
    }
}
