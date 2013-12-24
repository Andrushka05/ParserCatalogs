namespace ParserCatalog
{
    partial class Form1
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.label2 = new System.Windows.Forms.Label();
            this.Start = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.path = new System.Windows.Forms.TextBox();
            this.Open = new System.Windows.Forms.Button();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.photoCheck = new System.Windows.Forms.CheckBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.timeStripStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.countStripStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.groupBox1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 161);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 13);
            this.label2.TabIndex = 9;
            // 
            // Start
            // 
            this.Start.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Start.Location = new System.Drawing.Point(25, 174);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(142, 37);
            this.Start.TabIndex = 8;
            this.Start.Text = "Начать парсинг";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.path);
            this.groupBox1.Controls.Add(this.Open);
            this.groupBox1.Location = new System.Drawing.Point(22, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(265, 82);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Выбор папки для сохранения";
            // 
            // path
            // 
            this.path.Location = new System.Drawing.Point(6, 26);
            this.path.Name = "path";
            this.path.Size = new System.Drawing.Size(253, 20);
            this.path.TabIndex = 1;
            // 
            // Open
            // 
            this.Open.Location = new System.Drawing.Point(89, 52);
            this.Open.Name = "Open";
            this.Open.Size = new System.Drawing.Size(75, 23);
            this.Open.TabIndex = 0;
            this.Open.Text = "Открыть";
            this.Open.UseVisualStyleBackColor = true;
            this.Open.Click += new System.EventHandler(this.Open_Click);
            // 
            // treeView1
            // 
            this.treeView1.CheckBoxes = true;
            this.treeView1.Location = new System.Drawing.Point(291, 12);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(298, 274);
            this.treeView1.TabIndex = 11;
            this.treeView1.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.node_AfterCheck);
            // 
            // photoCheck
            // 
            this.photoCheck.AutoSize = true;
            this.photoCheck.Location = new System.Drawing.Point(28, 114);
            this.photoCheck.Name = "photoCheck";
            this.photoCheck.Size = new System.Drawing.Size(144, 17);
            this.photoCheck.TabIndex = 12;
            this.photoCheck.Text = "Сохранять фото в excel";
            this.photoCheck.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.timeStripStatus,
            this.countStripStatus});
            this.statusStrip1.Location = new System.Drawing.Point(0, 302);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(601, 22);
            this.statusStrip1.TabIndex = 13;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // timeStripStatus
            // 
            this.timeStripStatus.Name = "timeStripStatus";
            this.timeStripStatus.Size = new System.Drawing.Size(0, 17);
            // 
            // countStripStatus
            // 
            this.countStripStatus.Name = "countStripStatus";
            this.countStripStatus.Size = new System.Drawing.Size(0, 17);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(601, 324);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.photoCheck);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Start);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "ParserCatalog";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox path;
        private System.Windows.Forms.Button Open;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.CheckBox photoCheck;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel timeStripStatus;
        private System.Windows.Forms.ToolStripStatusLabel countStripStatus;
    }
}

