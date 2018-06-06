namespace ExcelConvertor
{
    partial class Form1
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.lblImportPath = new System.Windows.Forms.Label();
            this.btnImportBrowse = new System.Windows.Forms.Button();
            this.groupbox = new System.Windows.Forms.GroupBox();
            this.radioButton6 = new System.Windows.Forms.RadioButton();
            this.radioButton4 = new System.Windows.Forms.RadioButton();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton5 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.btnConvert = new System.Windows.Forms.Button();
            this.txtCellWidth = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblExportPath = new System.Windows.Forms.Label();
            this.btnExportBrowse = new System.Windows.Forms.Button();
            this.radioButton7 = new System.Windows.Forms.RadioButton();
            this.groupbox.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "目标：";
            // 
            // lblImportPath
            // 
            this.lblImportPath.Location = new System.Drawing.Point(63, 16);
            this.lblImportPath.Name = "lblImportPath";
            this.lblImportPath.Size = new System.Drawing.Size(252, 44);
            this.lblImportPath.TabIndex = 1;
            // 
            // btnImportBrowse
            // 
            this.btnImportBrowse.Location = new System.Drawing.Point(321, 15);
            this.btnImportBrowse.Name = "btnImportBrowse";
            this.btnImportBrowse.Size = new System.Drawing.Size(50, 25);
            this.btnImportBrowse.TabIndex = 2;
            this.btnImportBrowse.Text = "浏览";
            this.btnImportBrowse.UseVisualStyleBackColor = true;
            this.btnImportBrowse.Click += new System.EventHandler(this.btnImportBrowse_Click);
            // 
            // groupbox
            // 
            this.groupbox.Controls.Add(this.radioButton7);
            this.groupbox.Controls.Add(this.radioButton6);
            this.groupbox.Controls.Add(this.radioButton4);
            this.groupbox.Controls.Add(this.radioButton3);
            this.groupbox.Controls.Add(this.radioButton5);
            this.groupbox.Controls.Add(this.radioButton2);
            this.groupbox.Controls.Add(this.radioButton1);
            this.groupbox.Location = new System.Drawing.Point(18, 108);
            this.groupbox.Name = "groupbox";
            this.groupbox.Size = new System.Drawing.Size(353, 120);
            this.groupbox.TabIndex = 3;
            this.groupbox.TabStop = false;
            this.groupbox.Text = "转换方式";
            // 
            // radioButton6
            // 
            this.radioButton6.AutoSize = true;
            this.radioButton6.Location = new System.Drawing.Point(259, 58);
            this.radioButton6.Name = "radioButton6";
            this.radioButton6.Size = new System.Drawing.Size(89, 16);
            this.radioButton6.TabIndex = 7;
            this.radioButton6.TabStop = true;
            this.radioButton6.Text = "转换繁体Xml";
            this.radioButton6.UseVisualStyleBackColor = true;
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Location = new System.Drawing.Point(131, 58);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(113, 16);
            this.radioButton4.TabIndex = 6;
            this.radioButton4.TabStop = true;
            this.radioButton4.Text = "客户端繁体Excel";
            this.radioButton4.UseVisualStyleBackColor = true;
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(6, 58);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(113, 16);
            this.radioButton3.TabIndex = 5;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "服务器繁体Excel";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // radioButton5
            // 
            this.radioButton5.AutoSize = true;
            this.radioButton5.Location = new System.Drawing.Point(259, 30);
            this.radioButton5.Name = "radioButton5";
            this.radioButton5.Size = new System.Drawing.Size(65, 16);
            this.radioButton5.TabIndex = 4;
            this.radioButton5.TabStop = true;
            this.radioButton5.Text = "转换Xml";
            this.radioButton5.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(131, 30);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(89, 16);
            this.radioButton2.TabIndex = 1;
            this.radioButton2.Text = "客户端Excel";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Checked = true;
            this.radioButton1.Location = new System.Drawing.Point(6, 30);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(89, 16);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "服务器Excel";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(163, 275);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(75, 35);
            this.btnConvert.TabIndex = 4;
            this.btnConvert.Text = "转换";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // txtCellWidth
            // 
            this.txtCellWidth.Location = new System.Drawing.Point(99, 240);
            this.txtCellWidth.Name = "txtCellWidth";
            this.txtCellWidth.Size = new System.Drawing.Size(35, 21);
            this.txtCellWidth.TabIndex = 5;
            this.txtCellWidth.Text = "15";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 244);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "单元格宽度：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 69);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 7;
            this.label4.Text = "导出：";
            // 
            // lblExportPath
            // 
            this.lblExportPath.Location = new System.Drawing.Point(65, 68);
            this.lblExportPath.Name = "lblExportPath";
            this.lblExportPath.Size = new System.Drawing.Size(250, 36);
            this.lblExportPath.TabIndex = 8;
            // 
            // btnExportBrowse
            // 
            this.btnExportBrowse.Location = new System.Drawing.Point(321, 66);
            this.btnExportBrowse.Name = "btnExportBrowse";
            this.btnExportBrowse.Size = new System.Drawing.Size(50, 23);
            this.btnExportBrowse.TabIndex = 9;
            this.btnExportBrowse.Text = "浏览";
            this.btnExportBrowse.UseVisualStyleBackColor = true;
            this.btnExportBrowse.Click += new System.EventHandler(this.btnExportBrowse_Click);
            // 
            // radioButton7
            // 
            this.radioButton7.AutoSize = true;
            this.radioButton7.Location = new System.Drawing.Point(6, 86);
            this.radioButton7.Name = "radioButton7";
            this.radioButton7.Size = new System.Drawing.Size(71, 16);
            this.radioButton7.TabIndex = 8;
            this.radioButton7.TabStop = true;
            this.radioButton7.Text = "转换Json";
            this.radioButton7.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(383, 330);
            this.Controls.Add(this.btnExportBrowse);
            this.Controls.Add(this.lblExportPath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtCellWidth);
            this.Controls.Add(this.btnConvert);
            this.Controls.Add(this.groupbox);
            this.Controls.Add(this.btnImportBrowse);
            this.Controls.Add(this.lblImportPath);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "ExcelConvertor";
            this.groupbox.ResumeLayout(false);
            this.groupbox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblImportPath;
        private System.Windows.Forms.Button btnImportBrowse;
        private System.Windows.Forms.GroupBox groupbox;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.TextBox txtCellWidth;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblExportPath;
        private System.Windows.Forms.Button btnExportBrowse;
        private System.Windows.Forms.RadioButton radioButton5;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.RadioButton radioButton4;
        private System.Windows.Forms.RadioButton radioButton6;
        private System.Windows.Forms.RadioButton radioButton7;
    }
}

