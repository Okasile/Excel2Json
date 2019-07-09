namespace ExcelRead
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.pathExcel1 = new System.Windows.Forms.TextBox();
            this.excelPathFindBtn1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.pathFileSave1 = new System.Windows.Forms.TextBox();
            this.fileSaveFindBtn1 = new System.Windows.Forms.Button();
            this.saveBtn = new System.Windows.Forms.Button();
            this.sheet1 = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.RemoveHeadLine1 = new System.Windows.Forms.CheckBox();
            this.RemoveHeadLine2 = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.sheet2 = new System.Windows.Forms.NumericUpDown();
            this.fileSaveFindBtn2 = new System.Windows.Forms.Button();
            this.pathFileSave2 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.excelPathFindBtn2 = new System.Windows.Forms.Button();
            this.pathExcel2 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.sheet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sheet2)).BeginInit();
            this.SuspendLayout();
            // 
            // pathExcel1
            // 
            this.pathExcel1.Location = new System.Drawing.Point(12, 33);
            this.pathExcel1.Name = "pathExcel1";
            this.pathExcel1.Size = new System.Drawing.Size(583, 21);
            this.pathExcel1.TabIndex = 0;
            // 
            // excelPathFindBtn1
            // 
            this.excelPathFindBtn1.Location = new System.Drawing.Point(613, 25);
            this.excelPathFindBtn1.Name = "excelPathFindBtn1";
            this.excelPathFindBtn1.Size = new System.Drawing.Size(54, 35);
            this.excelPathFindBtn1.TabIndex = 1;
            this.excelPathFindBtn1.Text = "浏览";
            this.excelPathFindBtn1.UseVisualStyleBackColor = true;
            this.excelPathFindBtn1.Click += new System.EventHandler(this.excelPathFindBtn1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "Excel目录:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "Json保存目录:";
            // 
            // pathFileSave1
            // 
            this.pathFileSave1.Location = new System.Drawing.Point(14, 104);
            this.pathFileSave1.Name = "pathFileSave1";
            this.pathFileSave1.Size = new System.Drawing.Size(583, 21);
            this.pathFileSave1.TabIndex = 4;
            // 
            // fileSaveFindBtn1
            // 
            this.fileSaveFindBtn1.Location = new System.Drawing.Point(615, 98);
            this.fileSaveFindBtn1.Name = "fileSaveFindBtn1";
            this.fileSaveFindBtn1.Size = new System.Drawing.Size(54, 35);
            this.fileSaveFindBtn1.TabIndex = 5;
            this.fileSaveFindBtn1.Text = "浏览";
            this.fileSaveFindBtn1.UseVisualStyleBackColor = true;
            this.fileSaveFindBtn1.Click += new System.EventHandler(this.fileSaveFindBtn1_Click);
            // 
            // saveBtn
            // 
            this.saveBtn.Location = new System.Drawing.Point(257, 439);
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Size = new System.Drawing.Size(143, 47);
            this.saveBtn.TabIndex = 6;
            this.saveBtn.Text = "点击转换";
            this.saveBtn.UseVisualStyleBackColor = true;
            this.saveBtn.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // sheet1
            // 
            this.sheet1.Location = new System.Drawing.Point(461, 60);
            this.sheet1.Name = "sheet1";
            this.sheet1.Size = new System.Drawing.Size(120, 21);
            this.sheet1.TabIndex = 7;
            this.sheet1.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(372, 62);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 12);
            this.label3.TabIndex = 8;
            this.label3.Text = "Sheet Index :";
            // 
            // Remove1stLine1
            // 
            this.RemoveHeadLine1.AutoSize = true;
            this.RemoveHeadLine1.Location = new System.Drawing.Point(272, 60);
            this.RemoveHeadLine1.Name = "Remove1stLine1";
            this.RemoveHeadLine1.Size = new System.Drawing.Size(84, 16);
            this.RemoveHeadLine1.TabIndex = 9;
            this.RemoveHeadLine1.Text = "去掉第一行";
            this.RemoveHeadLine1.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.RemoveHeadLine2.AutoSize = true;
            this.RemoveHeadLine2.Location = new System.Drawing.Point(274, 262);
            this.RemoveHeadLine2.Name = "checkBox1";
            this.RemoveHeadLine2.Size = new System.Drawing.Size(84, 16);
            this.RemoveHeadLine2.TabIndex = 18;
            this.RemoveHeadLine2.Text = "去掉第一行";
            this.RemoveHeadLine2.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(374, 264);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 12);
            this.label4.TabIndex = 17;
            this.label4.Text = "Sheet Index :";
            // 
            // sheet2
            // 
            this.sheet2.Location = new System.Drawing.Point(463, 262);
            this.sheet2.Name = "sheet2";
            this.sheet2.Size = new System.Drawing.Size(120, 21);
            this.sheet2.TabIndex = 16;
            this.sheet2.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // fileSaveFindBtn2
            // 
            this.fileSaveFindBtn2.Location = new System.Drawing.Point(617, 300);
            this.fileSaveFindBtn2.Name = "fileSaveFindBtn2";
            this.fileSaveFindBtn2.Size = new System.Drawing.Size(54, 35);
            this.fileSaveFindBtn2.TabIndex = 15;
            this.fileSaveFindBtn2.Text = "浏览";
            this.fileSaveFindBtn2.UseVisualStyleBackColor = true;
            this.fileSaveFindBtn2.Click += new System.EventHandler(this.fileSaveFindBtn2_Click);
            // 
            // pathFileSave2
            // 
            this.pathFileSave2.Location = new System.Drawing.Point(16, 306);
            this.pathFileSave2.Name = "pathFileSave2";
            this.pathFileSave2.Size = new System.Drawing.Size(583, 21);
            this.pathFileSave2.TabIndex = 14;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 280);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 12);
            this.label5.TabIndex = 13;
            this.label5.Text = "Json保存目录:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(14, 211);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 12);
            this.label6.TabIndex = 12;
            this.label6.Text = "Excel目录:";
            // 
            // excelPathFindBtn2
            // 
            this.excelPathFindBtn2.Location = new System.Drawing.Point(615, 227);
            this.excelPathFindBtn2.Name = "excelPathFindBtn2";
            this.excelPathFindBtn2.Size = new System.Drawing.Size(54, 35);
            this.excelPathFindBtn2.TabIndex = 11;
            this.excelPathFindBtn2.Text = "浏览";
            this.excelPathFindBtn2.UseVisualStyleBackColor = true;
            this.excelPathFindBtn2.Click += new System.EventHandler(this.excelPathFindBtn2_Click);
            // 
            // pathExcel2
            // 
            this.pathExcel2.Location = new System.Drawing.Point(14, 235);
            this.pathExcel2.Name = "pathExcel2";
            this.pathExcel2.Size = new System.Drawing.Size(583, 21);
            this.pathExcel2.TabIndex = 10;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(748, 563);
            this.Controls.Add(this.RemoveHeadLine2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.sheet2);
            this.Controls.Add(this.fileSaveFindBtn2);
            this.Controls.Add(this.pathFileSave2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.excelPathFindBtn2);
            this.Controls.Add(this.pathExcel2);
            this.Controls.Add(this.RemoveHeadLine1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.sheet1);
            this.Controls.Add(this.saveBtn);
            this.Controls.Add(this.fileSaveFindBtn1);
            this.Controls.Add(this.pathFileSave1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.excelPathFindBtn1);
            this.Controls.Add(this.pathExcel1);
            this.Name = "Form1";
            this.Text = "Excel2Json";
            ((System.ComponentModel.ISupportInitialize)(this.sheet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sheet2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox pathExcel1;
        private System.Windows.Forms.Button excelPathFindBtn1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox pathFileSave1;
        private System.Windows.Forms.Button fileSaveFindBtn1;
        private System.Windows.Forms.Button saveBtn;
        private System.Windows.Forms.NumericUpDown sheet1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox RemoveHeadLine1;
        private System.Windows.Forms.CheckBox RemoveHeadLine2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown sheet2;
        private System.Windows.Forms.Button fileSaveFindBtn2;
        private System.Windows.Forms.TextBox pathFileSave2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button excelPathFindBtn2;
        private System.Windows.Forms.TextBox pathExcel2;
    }
}

