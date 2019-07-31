namespace tang.cdt_ec_order
{
    partial class MainForm
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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.btnLoadSOData = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ResultTextBox = new System.Windows.Forms.RichTextBox();
            this.btnLoadDOData = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnLoadSOData
            // 
            this.btnLoadSOData.Location = new System.Drawing.Point(18, 21);
            this.btnLoadSOData.Name = "btnLoadSOData";
            this.btnLoadSOData.Size = new System.Drawing.Size(111, 25);
            this.btnLoadSOData.TabIndex = 4;
            this.btnLoadSOData.Text = "获取销售订单数据";
            this.btnLoadSOData.UseVisualStyleBackColor = true;
            this.btnLoadSOData.Click += new System.EventHandler(this.btnLoadSOData_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ResultTextBox);
            this.groupBox1.Location = new System.Drawing.Point(12, 71);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(763, 416);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "信息：";
            // 
            // ResultTextBox
            // 
            this.ResultTextBox.Location = new System.Drawing.Point(6, 20);
            this.ResultTextBox.Name = "ResultTextBox";
            this.ResultTextBox.Size = new System.Drawing.Size(751, 390);
            this.ResultTextBox.TabIndex = 6;
            this.ResultTextBox.Text = "";
            // 
            // btnLoadDOData
            // 
            this.btnLoadDOData.Location = new System.Drawing.Point(185, 21);
            this.btnLoadDOData.Name = "btnLoadDOData";
            this.btnLoadDOData.Size = new System.Drawing.Size(111, 25);
            this.btnLoadDOData.TabIndex = 7;
            this.btnLoadDOData.Text = "获取发货管理数据";
            this.btnLoadDOData.UseVisualStyleBackColor = true;
            this.btnLoadDOData.Click += new System.EventHandler(this.BtnLoadDOData_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(787, 499);
            this.Controls.Add(this.btnLoadDOData);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnLoadSOData);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "大唐电子订单工具（科力普内部专用）";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnLoadSOData;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RichTextBox ResultTextBox;
        private System.Windows.Forms.Button btnLoadDOData;
    }
}

