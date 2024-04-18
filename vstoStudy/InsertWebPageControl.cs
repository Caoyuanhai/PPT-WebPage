using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;

namespace vstoStudy
{
    public partial class InsertWebPageControl : UserControl
    {

        string imgUrl;
      
        static string lastSelectedFile = "lastSelectedFile.txt"; // 存储上次选择的文件路径的文件名
        public InsertWebPageControl()
        {
            InitializeComponent();
        }

        private void InsertWebButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(InsertWebUrl.Text))
            {
                if (IsUrl(InsertWebUrl.Text))
                {
                    AddWebShapeToSlide(InsertWebUrl.Text);
                }
                else
                {
                    ShowErrorMessage("请输入有效的网址");
                }
            }
            else
            {
                ShowErrorMessage("网址不能为空");
            }

        }

        /// <summary>
        /// 添加网址
        /// </summary>
        /// <param name="url"></param>
        void AddWebShapeToSlide(string url)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            Slide slide = app.ActiveWindow.View.Slide;

            // 检查是否有图片加载到 pictureBox 控件中
            if (this.pictureBox.Image != null)
            {
           
                // 如果 pictureBox 中有图片，则将其用于创建形状
                Shape imageShape = slide.Shapes.AddPicture2(
                    // 获取 pictureBox 中的图片路径
                    imgUrl,
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue,
                    50, 50, 200, 200);

                imageShape.Name = "_webPage";
                imageShape.Tags.Add("URL", url);
            }
            else
            {
                // 如果 pictureBox 中没有图片，则创建普通的矩形形状
                Shape webShape = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    50, 50, 200, 200);

                webShape.Name = "_webPage";
                webShape.TextFrame.TextRange.Text = url + "\n 网页占位图形";
                webShape.Tags.Add("URL", url);
                // 设置形状边框不可见
                webShape.Line.Visible = Office.MsoTriState.msoFalse;
            }
        }

        void ShowErrorMessage(string message)
        {
            MessageBox.Show(message);
        }


        static bool IsUrl(string input)
        {
            // 使用正则表达式匹配网址格式
            Uri uriResult;
            return Uri.TryCreate(input, UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp|| uriResult.Scheme == Uri.UriSchemeHttps);
        }

        private void chooseImg_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Multiselect = false;  
            openFileDialog2.Title = "请选择文件";
            openFileDialog2.Filter = "图片(*.jpg,*.png)|*.jpg;*.png";

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrWhiteSpace(openFileDialog2.FileName))
                {
                    try
                    {
                        // Load the selected image file
                        Image selectedImage = Image.FromFile(openFileDialog2.FileName);

                        // Display the image in PictureBox
                        this.pictureBox.Image = selectedImage;

                        imgUrl = openFileDialog2.FileName;
                    }
                    catch (Exception ex)
                    {
                        // Handle any exceptions that might occur during file loading
                        MessageBox.Show("加载图片出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void clearImg_Click(object sender, EventArgs e)
        {
            this.pictureBox.Image = null;
        }

        private void startLocalServer_Click(object sender, EventArgs e)
        {
            // 尝试从文件中读取上次选择的文件路径
            string lastSelectedFilePath = ReadLastSelectedFile();

            // 创建一个打开文件对话框实例
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 设置对话框的标题和过滤条件
            openFileDialog.Title = "选择HTML文件";
            openFileDialog.Filter = "HTML文件 (*.html;*.htm)|*.html;*.htm";

            // 如果上次有记忆的文件路径，则设置默认路径
            if (!string.IsNullOrEmpty(lastSelectedFilePath) && File.Exists(lastSelectedFilePath))
            {
                openFileDialog.InitialDirectory = Path.GetDirectoryName(lastSelectedFilePath);
                openFileDialog.FileName = Path.GetFileName(lastSelectedFilePath);
            }

            // 显示打开文件对话框，并检查用户是否点击了确定按钮
            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                Slide slide = app.ActiveWindow.View.Slide;
                // 用户点击了确定按钮，获取所选文件的路径
                string selectedFilePath = openFileDialog.FileName;
                string selectedFileFolder = Path.GetDirectoryName(selectedFilePath);
                string fileName = Path.GetFileName(selectedFilePath);

                // 如果 pictureBox 中没有图片，则创建普通的矩形形状
                Shape webShape = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    50, 50, 200, 200);

                webShape.Name = "_webPage";
                webShape.TextFrame.TextRange.Text = selectedFilePath + "\n 本地站点占位图形";
                webShape.Tags.Add("LOCALPATH", selectedFilePath);

                // 设置形状边框不可见
                webShape.Line.Visible = Office.MsoTriState.msoFalse;
          
                // 将选择的文件路径保存到文件中
                SaveLastSelectedFile(selectedFilePath);

            }
            else
            {
                // 用户取消了选择，显示提示信息
                Console.WriteLine("用户取消了选择文件。");
            }
        }

        static string ReadLastSelectedFile()
        {
            // 如果存储文件存在，则读取其中的内容
            if (File.Exists(lastSelectedFile))
            {
                return File.ReadAllText(lastSelectedFile);
            }
            return null;
        }

        static void SaveLastSelectedFile(string filePath)
        {
            // 将选择的文件路径写入存储文件中
            File.WriteAllText(lastSelectedFile, filePath);
        }


    }
}
