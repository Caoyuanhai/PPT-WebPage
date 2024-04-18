using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace vstoStudy
{
    public partial class PPTWebPage : Form
    {
        public PPTWebPage(string url,int x,int y,int height,int width)
        {
            InitializeComponent();
            InitializeWebView2Async(url);
            Location = new Point(x, y);
            Size = new Size(width, height);
        }

        public async void InitializeWebView2Async(string url)
        {
            var env = await CoreWebView2Environment.CreateAsync(null, "C:\\temp");
            await webView21.EnsureCoreWebView2Async(env);
            webView21.CoreWebView2.Navigate(url);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
