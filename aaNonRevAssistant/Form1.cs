using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using zoyobar.shared.panzer.web.ib;
using System.Web;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;

namespace aaNonRevAssistant
{
    public partial class Form1 : Form
    {
        private const int INTERNET_OPTION_END_BROWSER_SESSION = 42;
        [DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int lpdwBufferLength);
        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetSetCookie(string lpszUrl, string lpszCookieName, string lpszCookieData);

        public string url = "";
        public string aaid = "407372";
        private string password = "";
        public string confirmationNumber = "";
        
        public IEBrowser ie = null;
        int windowtimes = -1;

        public Form1(string cNum, string pswd)
        {
            this.password = pswd;
            this.confirmationNumber = cNum;
            InitializeComponent();
            openAA();
        }

        private void openAA()
        {
            url = "https://nrtp.jetnet.aa.com/jciWeb/displayterm.do?pnrNo=" + confirmationNumber + "&personNo=00407372&source=nrtp";
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_END_BROWSER_SESSION, IntPtr.Zero, 0);
            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser1.Navigate(url);
            windowtimes = 0;
        }

        private void webBrowser_Complete(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if ((e.Url != webBrowser1.Url) || (webBrowser1.ReadyState != WebBrowserReadyState.Complete))
            {
                ((WebBrowser)sender).Document.Window.Error += new HtmlElementErrorEventHandler(Window_Error);
                return;
            }

            if (windowtimes == 0)
            {
                //maybe login page or check in page
                ie = new IEBrowser(this.webBrowser1);
                string doc = this.webBrowser1.DocumentText;
                if (doc.Contains("There was an error")) //cache issue, need refresh browser
                {
                }
                else if (doc.Contains("not eligible")) //not within 24 hours
                {
                    
                }
                else if (doc.Contains("Passengers"))
                {
                    //already in check in page, no need to log in
                    windowtimes = 1;
                }
                else
                {
                    string fillScript = "setTimeout(function(){document.getElementById('userLoginId').value = '" + aaid;
                    fillScript += "';document.getElementById('userPassword').value='" + password;
                    fillScript += "';document.getElementById('loginButton').click();},100);";
                    ie.ExecuteScript(fillScript);
                }
            }
            
            if (windowtimes == 1)
            {
                HtmlElementCollection checkbox = webBrowser1.Document.GetElementsByTagName("input");
                
                foreach (HtmlElement cb in checkbox)
                {
                    if (cb.GetAttribute("type").Equals("button") && cb.GetAttribute("name").Equals("submit1"))
                    {
                        cb.SetAttribute("id", "AAButton");
                        break;
                    }
                }

                ie.ExecuteScript("setTimeout(function(){document.getElementById('AAButton').click();},1000);");
            }
            else if (windowtimes == 2)
            {
                sendMail();
                Application.Exit();
            }
            windowtimes++;
        }

        private void Window_Error(object sender, HtmlElementErrorEventArgs e)
        {
            // Ignore the error and suppress the error dialog box.   
            e.Handled = false;
        }

        private void sendMail()
        {
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress("mmnnmit@gmail.com");
            msg.To.Add(new MailAddress("alexjhang@gmail.com"));

            msg.Subject = "AA Non-rev checked in " + confirmationNumber;
            msg.Body = " ";
            SmtpClient smtp = new SmtpClient("smtp.gmail.com");
            smtp.UseDefaultCredentials = false;
            smtp.EnableSsl = true;
            smtp.Port = 587;
            smtp.Credentials = new NetworkCredential("mmnnmit@gmail.com", "laohutu123");
            smtp.Send(msg);

        }
    }
}
