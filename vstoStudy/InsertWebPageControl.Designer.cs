
namespace vstoStudy
{
    partial class InsertWebPageControl
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.InsertWebButton = new System.Windows.Forms.Button();
            this.InsertWebUrl = new System.Windows.Forms.TextBox();
            this.pictureBox = new System.Windows.Forms.PictureBox();
            this.chooseImg = new System.Windows.Forms.Button();
            this.clearImg = new System.Windows.Forms.Button();
            this.startLocalServer = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // InsertWebButton
            // 
            this.InsertWebButton.Location = new System.Drawing.Point(3, 226);
            this.InsertWebButton.Name = "InsertWebButton";
            this.InsertWebButton.Size = new System.Drawing.Size(75, 23);
            this.InsertWebButton.TabIndex = 0;
            this.InsertWebButton.Text = "插入网页";
            this.InsertWebButton.UseVisualStyleBackColor = true;
            this.InsertWebButton.Click += new System.EventHandler(this.InsertWebButton_Click);
            // 
            // InsertWebUrl
            // 
            this.InsertWebUrl.Location = new System.Drawing.Point(3, 255);
            this.InsertWebUrl.Name = "InsertWebUrl";
            this.InsertWebUrl.Size = new System.Drawing.Size(357, 21);
            this.InsertWebUrl.TabIndex = 1;
            // 
            // pictureBox
            // 
            this.pictureBox.Location = new System.Drawing.Point(3, 32);
            this.pictureBox.Name = "pictureBox";
            this.pictureBox.Size = new System.Drawing.Size(356, 188);
            this.pictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox.TabIndex = 2;
            this.pictureBox.TabStop = false;
            // 
            // chooseImg
            // 
            this.chooseImg.Location = new System.Drawing.Point(3, 3);
            this.chooseImg.Name = "chooseImg";
            this.chooseImg.Size = new System.Drawing.Size(104, 23);
            this.chooseImg.TabIndex = 3;
            this.chooseImg.Text = "选择占位图片";
            this.chooseImg.UseVisualStyleBackColor = true;
            this.chooseImg.Click += new System.EventHandler(this.chooseImg_Click);
            // 
            // clearImg
            // 
            this.clearImg.Location = new System.Drawing.Point(113, 3);
            this.clearImg.Name = "clearImg";
            this.clearImg.Size = new System.Drawing.Size(75, 23);
            this.clearImg.TabIndex = 4;
            this.clearImg.Text = "清除占位图";
            this.clearImg.UseVisualStyleBackColor = true;
            this.clearImg.Click += new System.EventHandler(this.clearImg_Click);
            // 
            // startLocalServer
            // 
            this.startLocalServer.Location = new System.Drawing.Point(4, 283);
            this.startLocalServer.Name = "startLocalServer";
            this.startLocalServer.Size = new System.Drawing.Size(94, 23);
            this.startLocalServer.TabIndex = 5;
            this.startLocalServer.Text = "选择本地网站";
            this.startLocalServer.UseVisualStyleBackColor = true;
            this.startLocalServer.Click += new System.EventHandler(this.startLocalServer_Click);
            // 
            // InsertWebPageControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.startLocalServer);
            this.Controls.Add(this.clearImg);
            this.Controls.Add(this.chooseImg);
            this.Controls.Add(this.pictureBox);
            this.Controls.Add(this.InsertWebUrl);
            this.Controls.Add(this.InsertWebButton);
            this.Name = "InsertWebPageControl";
            this.Size = new System.Drawing.Size(500, 497);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button InsertWebButton;
        private System.Windows.Forms.TextBox InsertWebUrl;
        private System.Windows.Forms.PictureBox pictureBox;
        private System.Windows.Forms.Button chooseImg;
        private System.Windows.Forms.Button clearImg;
        private System.Windows.Forms.Button startLocalServer;
    }
}
