using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
//using Outlook = Microsoft.Office.Interop.Outlook;

namespace Email_Configuration
{
    public static class Email
    {
        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }
        public static void CreateLogFiles(string log)
        {
            try
            {
                string sPathName = System.IO.Path.Combine(AssemblyDirectory, "logs");
                string sLogFormat = DateTime.Now.ToShortDateString().ToString() + " " + DateTime.Now.ToLongTimeString().ToString() + " ==> ";

                //this variable used to create log filename format "
                //for example filename : ErrorLogYYYYMMDD
                string sYear = DateTime.Now.Year.ToString();
                string sMonth = DateTime.Now.Month.ToString();
                string sDay = DateTime.Now.Day.ToString();
                string sErrorTime = sYear + sMonth + sDay + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + ".txt";
                StreamWriter sw = new StreamWriter(System.IO.Path.Combine(sPathName, sErrorTime), true);
                sw.WriteLine(sLogFormat + log);
                sw.Flush();
                sw.Close();
            }
            catch { }
        }

        public static void SendEmail(string smtpHost, string smtpToEmail, string smtpFromEmail, string SubjectMsg, string MailBody, string OutputPdf, string DisplayName = "QuikTap Design Configuration - Engineer-To-Order")
        {
            if (smtpHost != null && smtpToEmail != null && smtpFromEmail != null && SubjectMsg != null && MailBody != null)
            {
                try
                {
                    CreateLogFiles("Host - " + smtpHost + "; To -" + smtpToEmail + "; From -" + smtpFromEmail + "; Subject -" + SubjectMsg + "; Body -" + MailBody + "; Drawing path - ");
                    SmtpClient SmtpServer = new SmtpClient();

                    //SmtpServer.Credentials = new System.Net.NetworkCredential(smtpUsername, smtpPwd);
                    //SmtpServer.Port = smtpPort;
                    SmtpServer.Host = smtpHost;
                    //SmtpServer.EnableSsl = enableSSL;
                    MailMessage mail = new MailMessage();

                    //string toAddress = smtpToEmail;

                    mail.To.Add(smtpToEmail);
                    mail.From = new MailAddress(smtpFromEmail);
                    mail.Subject = SubjectMsg;
                    mail.IsBodyHtml = true;
                    mail.Body = MailBody;
                    if (!string.IsNullOrEmpty(OutputPdf))
                    {
                        if (OutputPdf.Length > 0)
                        {
                            if (System.IO.File.Exists(OutputPdf))
                                mail.Attachments.Add(new Attachment(OutputPdf));
                        }
                    }
                    
                    mail.Priority = MailPriority.Normal;
                    //mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                    SmtpServer.Send(mail);
                    CreateLogFiles("Mail Sent Successfully");
                }
                catch (Exception ex)
                {
                    string exMsg = ex.Message;
                    if (ex.InnerException != null)
                        exMsg += Environment.NewLine + ex.InnerException.Message;
                    if (ex.StackTrace != null)
                        exMsg += Environment.NewLine + ex.StackTrace;
                    System.Diagnostics.Debug.WriteLine(exMsg);
                    CreateLogFiles(ex.ToString());
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Email congifuration validation failed!");
            }
        }

    }
}
