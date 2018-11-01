using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace tang.cdt_ec_order
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();

            EnsureLogin();
        }

        private void EnsureLogin()
        {
            webBrowserLogin.Navigate("http://tang.cdt-ec.com/workbench/index-zh_CN.html#/ifr/%252Fcpu-portal-fe%252Fportalcas.html%2523%252Fpages%252Fsaleorder%252Fsaleorderlist");
        }

        private void confirmLoginBtn_Click(object sender, EventArgs e)
        {
            MainForm mainForm = (MainForm)Owner;

            mainForm.Cookie = GetCookies("http://tang.cdt-ec.com/workbench/index-zh_CN.html");

            Close();
        }

        private static string GetCookies(string url)
        {
            uint datasize = 256;

            StringBuilder cookieData = new StringBuilder((int)datasize);

            if (InternetGetCookieEx(url, null, cookieData, ref datasize, 0x2000, IntPtr.Zero))
                return cookieData.ToString();

            cookieData = new StringBuilder((int)datasize);

            if (!InternetGetCookieEx(url, null, cookieData, ref datasize, 0x00002000, IntPtr.Zero))
                return null;

            return cookieData.ToString();
        }

        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetGetCookieEx(string pchURL, string pchCookieName, StringBuilder pchCookieData, ref System.UInt32 pcchCookieData, int dwFlags, IntPtr lpReserved);
    }
}
