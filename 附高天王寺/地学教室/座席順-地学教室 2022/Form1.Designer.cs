namespace 座席順_地学教室_2022
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.files = new System.Windows.Forms.ToolStripDropDownButton();
            this.環境設定ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.help = new System.Windows.Forms.ToolStripButton();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.files,
            this.toolStripButton2,
            this.toolStripButton1,
            this.help});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(350, 27);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // files
            // 
            this.files.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.files.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.環境設定ToolStripMenuItem});
            this.files.Image = ((System.Drawing.Image)(resources.GetObject("files.Image")));
            this.files.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.files.Name = "files";
            this.files.Size = new System.Drawing.Size(65, 24);
            this.files.Text = "ファイル";
            this.files.Click += new System.EventHandler(this.files_Click);
            // 
            // 環境設定ToolStripMenuItem
            // 
            this.環境設定ToolStripMenuItem.Name = "環境設定ToolStripMenuItem";
            this.環境設定ToolStripMenuItem.Size = new System.Drawing.Size(152, 26);
            this.環境設定ToolStripMenuItem.Text = "環境設定";
            this.環境設定ToolStripMenuItem.Click += new System.EventHandler(this.環境設定ToolStripMenuItem_Click);
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton2.Image")));
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(46, 24);
            this.toolStripButton2.Text = "ツール";
            this.toolStripButton2.Click += new System.EventHandler(this.toolStripButton2_Click);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(73, 24);
            this.toolStripButton1.Text = "拡張機能";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click_1);
            // 
            // help
            // 
            this.help.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.help.Image = ((System.Drawing.Image)(resources.GetObject("help.Image")));
            this.help.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.help.Name = "help";
            this.help.Size = new System.Drawing.Size(48, 24);
            this.help.Text = "ヘルプ";
            this.help.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Yu Gothic UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.textBox1.Location = new System.Drawing.Point(12, 118);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(326, 27);
            this.textBox1.TabIndex = 1;
            this.textBox1.Text = "大天才上尾先輩大感謝";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(173, 169);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(74, 29);
            this.button1.TabIndex = 2;
            this.button1.Text = "実行";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(253, 169);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(85, 29);
            this.button2.TabIndex = 3;
            this.button2.Text = "閉じる";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(326, 40);
            this.label1.TabIndex = 4;
            this.label1.Text = "生徒の人数を半角数字で入力して下さい。\r\n数値は3以上40以下の自然数である必要があります。";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(350, 221);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.toolStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ペア作成ツール 2022 - 地学教室 Edition";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ToolStrip toolStrip1;
        private ToolStripDropDownButton files;
        private ToolStripMenuItem 環境設定ToolStripMenuItem;
        private ToolStripButton help;
        private TextBox textBox1;
        private Button button1;
        private Button button2;
        private ToolStripButton toolStripButton1;
        private ToolStripButton toolStripButton2;
        private Label label1;
    }
}