namespace tang.cdt_ec_order
{
    partial class LoginForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.webBrowserLogin = new System.Windows.Forms.WebBrowser();
            this.confirmLoginBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // webBrowserLogin
            // 
            this.webBrowserLogin.Location = new System.Drawing.Point(1, 12);
            this.webBrowserLogin.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowserLogin.Name = "webBrowserLogin";
            this.webBrowserLogin.Size = new System.Drawing.Size(800, 403);
            this.webBrowserLogin.TabIndex = 0;
            // 
            // confirmLoginBtn
            // 
            this.confirmLoginBtn.Location = new System.Drawing.Point(697, 422);
            this.confirmLoginBtn.Name = "confirmLoginBtn";
            this.confirmLoginBtn.Size = new System.Drawing.Size(75, 23);
            this.confirmLoginBtn.TabIndex = 1;
            this.confirmLoginBtn.Text = "确认已登陆";
            this.confirmLoginBtn.UseVisualStyleBackColor = true;
            this.confirmLoginBtn.Click += new System.EventHandler(this.confirmLoginBtn_Click);
            // 
            // LoginForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.confirmLoginBtn);
            this.Controls.Add(this.webBrowserLogin);
            this.Name = "LoginForm";
            this.Text = "登录大唐电子";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowserLogin;
        private System.Windows.Forms.Button confirmLoginBtn;
    }
}