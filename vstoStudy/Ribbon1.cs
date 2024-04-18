using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using System.Net.NetworkInformation;
using System.Net;
using System.Threading.Tasks;
using System.IO;

namespace vstoStudy
{
    public partial class Ribbon1
    {
        
        PowerPoint.Application app; //实例化PPT

        Dictionary<int, List<PPTWebPage>> webPagesDictionary ;//web窗体集合


        private List<PPTWebPage> openPPTWebPage = new List<PPTWebPage>();

        Timer timer; //定时器判断是否是最后一页

        int totalSlide; //幻灯片总页数

        static int port = 15688; // 端口号

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
           
            app = Globals.ThisAddIn.Application;

            webPagesDictionary = new Dictionary<int,List<PPTWebPage>>();
            app.SlideShowBegin += App_SlideShowBegin;
            app.SlideShowNextSlide += App_SlideShowNextSlide;
            app.SlideShowEnd += App_SlideShowEnd;
        
        }

        /// <summary>
        /// 定时器（因为演示环节最后一页的黑幻灯片无法触发PPT页面切换事件，所以添加定时器实现在最后一页清除web窗体的功能）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timer_Tick(object sender, EventArgs e)
        {
            int currentIndex = app.ActivePresentation.SlideShowWindow.View.CurrentShowPosition;
            if (currentIndex > totalSlide)
            {
                CloseOpenForms();
            }
        }


        /// <summary>
        /// 幻灯片开始放映
        /// </summary>
        /// <param name="Wn"></param>
        private void App_SlideShowBegin(SlideShowWindow Wn)
        {
            HttpServer.runningServer = true;
            Presentation presentation= app.ActivePresentation;
            totalSlide = presentation.Slides.Count;
            int x2 = 0; //x轴偏移量
            float ScalingFactor = 1;

            Screen screen = getSlideShowScreen();
            if (screen != null)
            {
                x2 = screen.Bounds.Left;
                ScalingFactor = GetScalingFactor(screen);
            }
            else {
                ScalingFactor = GetScalingFactor(Screen.PrimaryScreen);
            }
            
            int slideIndex = 1;
            float widthRatio = getWidthRatio();
            float heightRatio = getHeightRatio();
          
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                List<PPTWebPage> pages = new List<PPTWebPage>();


                for (int i = 1; i < slide.Shapes.Count+1; i++)
                {
                    if (slide.Shapes[i].Name== "_webPage")
                    {
                        var shape = slide.Shapes[i];
                        int x = x2 + Convert.ToInt32(shape.Left * widthRatio / ScalingFactor);
                        int y = Convert.ToInt32(shape.Top * heightRatio / ScalingFactor);
                        int width = Convert.ToInt32(shape.Width * widthRatio / ScalingFactor);
                        int height = Convert.ToInt32(shape.Height * heightRatio / ScalingFactor);
                        if (shape.Tags["LOCALPATH"] != "")
                        {
                            string path = shape.Tags["LOCALPATH"];
                            string selectedFileFolder = Path.GetDirectoryName(path);
                            string fileName = Path.GetFileName(path);
                            startServer(selectedFileFolder);
                            string url = "http://localhost:" + port + "/" + fileName;
                            PPTWebPage pPTWebPage = new PPTWebPage(url, x, y, height, width);
                            pages.Add(pPTWebPage);

                        }
                        else {
                            string url = shape.Tags["URL"];
                            PPTWebPage pPTWebPage = new PPTWebPage(url,x,y,height,width);
                            pages.Add(pPTWebPage);
                        }


                    }
                }
                if (pages.Count>0)
                {
                    webPagesDictionary.Add(slideIndex, pages);
                }

                slideIndex++;
            }
            timer = new Timer(); // 每秒检测一次
            timer.Interval = 100;
            timer.Tick += Timer_Tick;
            timer.Start();
        }


        public static void startServer(string path)
        {
            if (!PortInUse(port))
            {
                Task.Run(() =>
                {
                    try
                    {
                        HttpServer.StartServer(path, port);
                    }
                    catch (HttpListenerException ex)
                    {
                        Console.WriteLine(ex);
                    }
                });
            }
            else
            {
                port++;
                startServer(path);
            }
        }

        public static bool PortInUse(int port)
        {
            bool inUse = false;

            IPGlobalProperties ipProperties = IPGlobalProperties.GetIPGlobalProperties();
            IPEndPoint[] ipEndPoints = ipProperties.GetActiveTcpListeners();

            foreach (IPEndPoint endPoint in ipEndPoints)
            {
                if (endPoint.Port == port)
                {
                    inUse = true;
                    break;
                }
            }
            return inUse;
        }
        /// <summary>
        /// 幻灯片结束放映
        /// </summary>
        /// <param name="Pres"></param>
        private void App_SlideShowEnd(Presentation Pres)
        {
            foreach (KeyValuePair<int, List<PPTWebPage>> kvp in webPagesDictionary)
            {
                if (kvp.Value.Count>0)
                {
                    foreach (PPTWebPage page in kvp.Value)
                    {
                        page.Close();
                    }
                }
            }

            webPagesDictionary.Clear();

            timer.Stop();
            timer.Dispose();

            HttpServer.stopAllServer();
        }


        /// <summary>
        /// 幻灯片切换事件
        /// </summary>
        /// <param name="Wn"></param>
        private void App_SlideShowNextSlide(SlideShowWindow Wn)
        {
            CloseOpenForms();
            //当前页码 
            int currentPage = Wn.View.CurrentShowPosition;
            List<PPTWebPage> pages;
            if (webPagesDictionary.TryGetValue(currentPage, out pages))
            {
                foreach (var page in pages)
                {
                    page.Show();
                    openPPTWebPage.Add(page);
                }
            }
        }

        // 关闭已打开的窗口
        private void CloseOpenForms()
        {
            foreach (var form in openPPTWebPage)
            {
                form.Hide();
            }
            openPPTWebPage.Clear();
        }

        private void InsertWebPage_Click(object sender, RibbonControlEventArgs e)
        {
            InsertWebPagePane.taskPane.Visible = true;
        }

        #region 获取屏幕分辨率及缩放倍数相关方法

        private float getWidthRatio() {

            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
      

            return screenWidth / 1920f * 2;
        }

        private float getHeightRatio()
        {
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;
            return screenHeight / 1080f * 2;
        }

        private Screen getSlideShowScreen() {
            Screen[] screens = Screen.AllScreens;
            float left =Convert.ToInt32( app.ActivePresentation.SlideShowWindow.Left);
            foreach (var screen in screens)
            {
                if (screen.Bounds.Left<0&& left<0|| screen.Bounds.Left > 0 && left > 0)
                {
                    return screen;
                }
                
            }
            return null;
        }

        public static float GetScalingFactor(Screen screen)
        {
            IntPtr hdc = NativeMethods.CreateDC("DISPLAY", screen.DeviceName, null, IntPtr.Zero);
            if (hdc != IntPtr.Zero)
            {
                int logicalScreenHeight = NativeMethods.GetDeviceCaps(hdc, (int)DeviceCap.VERTRES);
                int physicalScreenHeight = NativeMethods.GetDeviceCaps(hdc, (int)DeviceCap.DESKTOPVERTRES);
                NativeMethods.DeleteDC(hdc);
                float scalingFactor = (float)physicalScreenHeight / (float)logicalScreenHeight;
                return scalingFactor;
            }
            else
            {
                throw new Exception("Failed to create device context for the screen.");
            }
        }


        // 声明获取设备信息的相关方法
        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117,
        }

        internal static class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("gdi32.dll")]
            public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

            [System.Runtime.InteropServices.DllImport("gdi32.dll")]
            public static extern IntPtr CreateDC(string lpszDriver, string lpszDevice, string lpszOutput, IntPtr lpInitData);

            [System.Runtime.InteropServices.DllImport("gdi32.dll")]
            public static extern bool DeleteDC(IntPtr hdc);
        }

        #endregion

    }
}
