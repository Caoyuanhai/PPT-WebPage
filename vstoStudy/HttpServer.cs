using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace vstoStudy
{
    class HttpServer
    {

        public static bool runningServer = true;

        public static List<HttpListener> httpListeners = new List<HttpListener>();

        public static void StartServer(string path, int port)
        {
            string basePath = path;


            HttpListener listener = new HttpListener();
            listener.Prefixes.Add("http://localhost:" + port + "/"); // 设置监听的地址和端口

            listener.Start();
            httpListeners.Add(listener);


            while (runningServer)
            {
                HttpListenerContext context = listener.GetContext();
                HttpListenerRequest request = context.Request;
                HttpListenerResponse response = context.Response;

                string filename = Path.Combine(basePath, request.Url.AbsolutePath.Substring(1));
                if (File.Exists(filename))
                {
                    byte[] buffer = File.ReadAllBytes(filename);

                    // 设置Content-Type
                    string contentType = GetContentType(filename);
                    response.ContentType = contentType;

                    response.ContentLength64 = buffer.Length;
                    response.OutputStream.Write(buffer, 0, buffer.Length);
                }
                response.Close();
            }

        }

        public static void stopAllServer() {
            runningServer = false;
            if (httpListeners.Count>0)
            {
                foreach (var item in httpListeners)
                {
                    item.Stop();
                }
            }
        }


        // 根据文件扩展名获取对应的Content-Type
        private static string GetContentType(string filename)
        {
            string extension = Path.GetExtension(filename);
            switch (extension)
            {
                case ".html":
                    return "text/html";
                case ".css":
                    return "text/css";
                case ".js":
                    return "application/javascript";
                case ".wasm":
                    return "application/wasm";
                case ".mp4": // 添加对MP4文件的支持
                    return "video/mp4";
                case ".png": // 添加对PNG文件的支持
                    return "image/png";
                default:
                    return "application/octet-stream"; // 默认值
            }
        }
    }
}
