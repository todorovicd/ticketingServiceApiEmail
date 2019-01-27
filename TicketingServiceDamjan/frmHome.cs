using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;
using S22.Imap;
using System.IO;

namespace TicketingServiceDamjan
{
    public partial class frmHome : Form
    {
        private static string StartDate;
        private static string EndDate;

        private static string emailAddress = System.Environment.GetEnvironmentVariable("emailAddress", EnvironmentVariableTarget.User);
        private static string password = System.Environment.GetEnvironmentVariable("emailPassword", EnvironmentVariableTarget.User);

        static frmHome f;


        public frmHome()
        {
            InitializeComponent();
            f = this;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            StartReceiving();
        }

        private void StartReceiving()
        {
            Task.Run(() =>
            {
                using (ImapClient client = new ImapClient("imap.gmail.com", 993, emailAddress, password, AuthMethod.Login, true))

                {
                    if (client.Supports("IDLE") == false)
                    {                       
                        MessageBox.Show("Server does not support IMAP IDLE");
                        return;
                    }
                    client.NewMessage += new EventHandler<IdleMessageEventArgs>(OnNewMessage);
                    while (true) ;

                }
            });
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void OnNewMessage(object sender, IdleMessageEventArgs e)
        {
            try
            {
                MailMessage m = e.Client.GetMessage(e.MessageUID, FetchOptions.Normal);
                f.Invoke((MethodInvoker)delegate
                {
                    if (m.Subject.Contains("Help Desk Submission"))
                    {
                        string emailBody = m.Body.ToString();

                        if (emailBody.Contains("High Priority"))
                        {
                          
                            StartDate = DateTime.Now.ToString("yyyyMMdd");

                            EndDate = DateTime.Now.AddDays(2).ToString("yyyyMMdd");

                        }
                        else if (emailBody.Contains("Medium Priority"))
                        {

                            StartDate = DateTime.Now.ToString("yyyyMMdd");

                            EndDate = DateTime.Now.AddDays(5).ToString("yyyyMMdd");
                        }

                        else if (emailBody.Contains("Low Priority"))
                        {
                            StartDate = DateTime.Now.ToString("yyyyMMdd");

                            EndDate = DateTime.Now.AddDays(10).ToString("yyyyMMdd");
                        }

                        // here extract title for the ticket

                        int a = emailBody.IndexOf("Short Description");
                        int b = emailBody.IndexOf("Priority");


                        int startIndex1 = a + 48;
                        int endIndex1 = b - 3;


                        int length1 = endIndex1 - startIndex1;

                        string ticketTitle1 = emailBody.Substring(startIndex1, length1);

                        String ticketTitle = ticketTitle1.Replace("\r\n", " ");


                        //////////////////////////////////


                        //// get email address to email user block


                        int i = emailBody.IndexOf("Email Address");
                        int z = emailBody.IndexOf("email address after @ change");


                        int startIndex = i + 17;
                        int endIndex = z + 9;


                        int length = endIndex - startIndex;

                        string userEmailAddress = emailBody.Substring(startIndex, length);

                        //////////////////////////////////////////////


                        string content = System.IO.File.ReadAllText(@"File Location Change");
                        int x = Int32.Parse(content);
                        int newTicketNum = x + 1;


                        string updTicketNum = newTicketNum.ToString();

                        string[] update = { updTicketNum };

                        StreamWriter File = new StreamWriter(@"File Location Change");
                        File.Write(updTicketNum);
                        File.Close();


                        string timeStampTicket = DateTime.Now.ToString("dd/MM/yyyy");


                        f.txtEmailText.AppendText(newTicketNum + " " + "-" + " " + ticketTitle + Environment.NewLine);

                        string[] lines = { StartDate + ";" + EndDate + ";" + "1000" + ";" + StartDate + ";" + EndDate + ";" + "0930" + ";" + userEmailAddress + ";" + newTicketNum + Environment.NewLine + "JR;AD;DT;DM;JH" + Environment.NewLine + ticketTitle + Environment.NewLine + m.Body };


                        using (StreamWriter outputFile = new StreamWriter(Path.Combine(@"File Location Change", "IT" + "-" + EndDate + "-" + "0930" + "-" + "JRADDTDMJH" + "-" + "DT" + "-" + newTicketNum + "." + "dat")))


                        {
                            foreach (string line in lines)
                                outputFile.WriteLine(line);
                        }

                        var message = new MailMessage(emailAddress, userEmailAddress);
                        message.Subject = ("Support Request" + "(" + newTicketNum + ")");
                        message.Body = ("IT DEPARTMENT" + Environment.NewLine + Environment.NewLine + "A support ticket has been registered for you." + Environment.NewLine + "We will attend to the matter as soon as possible." + Environment.NewLine + Environment.NewLine + "Ticket Number: " + newTicketNum + Environment.NewLine + "Issue: " + ticketTitle + Environment.NewLine + "If you have any question regarding this matter, please call extension: 4444");

                        using (SmtpClient mailer = new SmtpClient("smtp.gmail.com", 587))
                        {
                            mailer.Credentials = new NetworkCredential(emailAddress, password);
                            mailer.EnableSsl = true;
                            mailer.Send(message);
                        }
                    }
                    else if (m.Subject.Contains("Contact Admin Message from Password Reset PRO"))
                    {
                        string emailBody = m.Body.ToString();

                        //// get email address to email user block

                        int i = emailBody.IndexOf("Email:");
                        int z = emailBody.IndexOf("User Message:");

                        int startIndex = i + 7;
                        int endIndex = z;


                        int length = endIndex - startIndex;

                        string userEmailAddress = emailBody.Substring(startIndex, length);

                        if (userEmailAddress.Contains("@"))
                        {
                            var message = new MailMessage(emailAddress, userEmailAddress);
                            message.Subject = ("Password Reset Request");
                            message.Body = ("IT DEPARTMENT" + Environment.NewLine + Environment.NewLine + "Hello, you've contacted IT Department regarding your login information via Password Reset Pro.  Passwords cannot be reset by emailed request   You must phone or visit the technology center, or use the password recovery portal located on the MyMoval page at http://www.moval.edu/home-page/mymoval/." + Environment.NewLine + Environment.NewLine + "MoVal Email Login:" + Environment.NewLine + "Most Missouri Valley College Network users have an alternate email address on file such as a yahoo, gmail, or hotmail account where their login information was sent when it was generated, or the last time they contacted the support desk to reset it.  Please check this mailbox for the last message from the Missouri Valley College IT Dept.  If you are unable to locate the message or unable to access the account, please phone the tech center at 999-999-9999." + Environment.NewLine + Environment.NewLine + Environment.NewLine");

                            using (SmtpClient mailer = new SmtpClient("smtp.gmail.com", 587))
                            {
                                mailer.Credentials = new NetworkCredential(emailAddress, password);
                                mailer.EnableSsl = true;
                                mailer.Send(message);
                            }

                            f.txtEmailText.AppendText(userEmailAddress + " " + "-" + " " + "Just got an email!" + Environment.NewLine);
                        }
                        else
                        {
                            System.Media.SystemSounds.Exclamation.Play();
                            MessageBox.Show("No email address");

                            int a = emailBody.IndexOf("Name:");
                            int b = emailBody.IndexOf("Phone:");

                            int startIndex2 = a + 6;
                            int endIndex2 = b;

                            int length2 = endIndex2 - startIndex2;

                            string userName = emailBody.Substring(startIndex2, length2);

                            string log = userName + DateTime.Now.ToString();

                            File.AppendAllText(@"C:\failedEmails.txt", log + Environment.NewLine);
                        }
                    }
                    else if (m.Subject.Contains("Suspicious Login"))
                    {
                        string emailBody = m.Body.ToString();

                        //get IP address

                        int a = emailBody.IndexOf("IP from which the login attempt was detected:");
                        int b = emailBody.IndexOf("Go to Alert Center to see more details for this alert.");

                        int startIndex1 = a + 45;
                        int endIndex1 = b;

                        int length1 = endIndex1 - startIndex1;

                        string ticketTitle1 = emailBody.Substring(startIndex1, length1);

                        string ipAddress = ticketTitle1.Replace("\r\n", " ");


                        //////////////////////////////////////

                        //get user email address

                        int c = emailBody.IndexOf("User:");
                        int d = emailBody.IndexOf("IP from which the login attempt was detected:");

                        int startIndex2 = c + 5;
                        int endIndex3 = d;

                        int length2 = endIndex3 - startIndex2;

                        string emailAddressUserSusp1 = emailBody.Substring(startIndex2, length2);

                        string emailAddressUserSusp2 = emailAddressUserSusp1.Replace("\r\n", " ");

                        /////////////////////////////////////



                        string URL1 = "http://ip-api.com/line/";

                        string URL = URL1 + ipAddress.TrimStart();

                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);
                        request.Method = "POST";
                        request.ContentType = "application/json";
                        request.ContentLength = ipAddress.Length;
                        using (Stream webStream = request.GetRequestStream())
                        using (StreamWriter requestWriter = new StreamWriter(webStream, System.Text.Encoding.ASCII))
                        {
                            requestWriter.Write(ipAddress);
                        }

                        try
                        {
                            WebResponse webResponse = request.GetResponse();
                            using (Stream webStream = webResponse.GetResponseStream() ?? Stream.Null)
                            using (StreamReader responseReader = new StreamReader(webStream))
                            {
                                string response = responseReader.ReadToEnd();

                                var message = new MailMessage(emailAddress, "email address here");
                                message.Subject = ("Suspicious Login Information");
                                message.Body = ("User: "+ emailAddressUserSusp2 + Environment.NewLine + Environment.NewLine + "IP Address: " + ipAddress + Environment.NewLine + Environment.NewLine + "Response From IP-API: " + Environment.NewLine + Environment.NewLine + response);

                                using (SmtpClient mailer = new SmtpClient("smtp.gmail.com", 587))
                                {
                                    mailer.Credentials = new NetworkCredential(emailAddress, password);
                                    mailer.EnableSsl = true;
                                    mailer.Send(message);
                                }
                                f.txtEmailText.AppendText("Name just got an email!" + Environment.NewLine);
                            }
                        }
                        catch (Exception exe)
                        {

                        }
                    }
               });
            }
            catch (Exception ex)
            {

            }
         }
     }
 }
    