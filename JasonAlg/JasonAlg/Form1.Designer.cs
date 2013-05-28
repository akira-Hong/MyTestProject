namespace JasonAlg
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.newJobToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtCurDir = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.cmbJoblist = new System.Windows.Forms.ComboBox();
            this.btnReload = new System.Windows.Forms.Button();
            this.txtTpn = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtTpsn = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPosn = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnCalc = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.StatusProgress = new System.Windows.Forms.ToolStripProgressBar();
            this.statusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.txtStart = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.cmbPercent = new System.Windows.Forms.ComboBox();
            this.menuStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newJobToolStripMenuItem,
            this.exitToolStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(590, 24);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // newJobToolStripMenuItem
            // 
            this.newJobToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copyToolStripMenuItem,
            this.exitToolStripMenuItem2});
            this.newJobToolStripMenuItem.Name = "newJobToolStripMenuItem";
            this.newJobToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.newJobToolStripMenuItem.Text = "Job";
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(102, 22);
            this.copyToolStripMenuItem.Text = "Copy";
            this.copyToolStripMenuItem.Click += new System.EventHandler(this.copyToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem2
            // 
            this.exitToolStripMenuItem2.Name = "exitToolStripMenuItem2";
            this.exitToolStripMenuItem2.Size = new System.Drawing.Size(102, 22);
            this.exitToolStripMenuItem2.Text = "Exit";
            this.exitToolStripMenuItem2.Click += new System.EventHandler(this.exitToolStripMenuItem2_Click);
            // 
            // exitToolStripMenuItem1
            // 
            this.exitToolStripMenuItem1.Name = "exitToolStripMenuItem1";
            this.exitToolStripMenuItem1.Size = new System.Drawing.Size(52, 20);
            this.exitToolStripMenuItem1.Text = "About";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "Work Directory";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "Current Job";
            // 
            // txtCurDir
            // 
            this.txtCurDir.Location = new System.Drawing.Point(119, 34);
            this.txtCurDir.Name = "txtCurDir";
            this.txtCurDir.ReadOnly = true;
            this.txtCurDir.Size = new System.Drawing.Size(395, 21);
            this.txtCurDir.TabIndex = 8;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(511, 32);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(78, 23);
            this.button1.TabIndex = 12;
            this.button1.Text = "Change";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // cmbJoblist
            // 
            this.cmbJoblist.FormattingEnabled = true;
            this.cmbJoblist.Location = new System.Drawing.Point(119, 65);
            this.cmbJoblist.Name = "cmbJoblist";
            this.cmbJoblist.Size = new System.Drawing.Size(395, 20);
            this.cmbJoblist.TabIndex = 13;
            this.cmbJoblist.SelectedIndexChanged += new System.EventHandler(this.cmbJoblist_SelectedIndexChanged);
            // 
            // btnReload
            // 
            this.btnReload.Location = new System.Drawing.Point(511, 65);
            this.btnReload.Name = "btnReload";
            this.btnReload.Size = new System.Drawing.Size(78, 23);
            this.btnReload.TabIndex = 14;
            this.btnReload.Text = "Reload";
            this.btnReload.UseVisualStyleBackColor = true;
            this.btnReload.Click += new System.EventHandler(this.btnReload_Click);
            // 
            // txtTpn
            // 
            this.txtTpn.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtTpn.Location = new System.Drawing.Point(165, 94);
            this.txtTpn.Name = "txtTpn";
            this.txtTpn.ReadOnly = true;
            this.txtTpn.Size = new System.Drawing.Size(102, 21);
            this.txtTpn.TabIndex = 16;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(117, 97);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 12);
            this.label3.TabIndex = 15;
            this.label3.Text = "TPN";
            // 
            // txtTpsn
            // 
            this.txtTpsn.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtTpsn.Location = new System.Drawing.Point(352, 98);
            this.txtTpsn.Name = "txtTpsn";
            this.txtTpsn.ReadOnly = true;
            this.txtTpsn.Size = new System.Drawing.Size(102, 21);
            this.txtTpsn.TabIndex = 18;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(287, 100);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(38, 12);
            this.label4.TabIndex = 17;
            this.label4.Text = "TPSN";
            // 
            // txtPosn
            // 
            this.txtPosn.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtPosn.Location = new System.Drawing.Point(165, 121);
            this.txtPosn.Name = "txtPosn";
            this.txtPosn.ReadOnly = true;
            this.txtPosn.Size = new System.Drawing.Size(102, 21);
            this.txtPosn.TabIndex = 20;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(117, 124);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(39, 12);
            this.label5.TabIndex = 19;
            this.label5.Text = "POSN";
            // 
            // btnCalc
            // 
            this.btnCalc.Location = new System.Drawing.Point(119, 165);
            this.btnCalc.Name = "btnCalc";
            this.btnCalc.Size = new System.Drawing.Size(206, 52);
            this.btnCalc.TabIndex = 21;
            this.btnCalc.Text = "GO";
            this.btnCalc.UseVisualStyleBackColor = true;
            this.btnCalc.Click += new System.EventHandler(this.btnCalc_Click);
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(510, 165);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(79, 52);
            this.btnExit.TabIndex = 22;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StatusProgress,
            this.statusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 223);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(590, 22);
            this.statusStrip1.TabIndex = 23;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // StatusProgress
            // 
            this.StatusProgress.Name = "StatusProgress";
            this.StatusProgress.Size = new System.Drawing.Size(100, 16);
            // 
            // statusLabel1
            // 
            this.statusLabel1.Name = "statusLabel1";
            this.statusLabel1.Size = new System.Drawing.Size(16, 17);
            this.statusLabel1.Text = "...";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // txtStart
            // 
            this.txtStart.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtStart.Location = new System.Drawing.Point(352, 123);
            this.txtStart.Name = "txtStart";
            this.txtStart.ReadOnly = true;
            this.txtStart.Size = new System.Drawing.Size(102, 21);
            this.txtStart.TabIndex = 25;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(287, 126);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 12);
            this.label6.TabIndex = 24;
            this.label6.Text = "Data Start";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(369, 165);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(135, 52);
            this.button2.TabIndex = 26;
            this.button2.Text = "Get Permission";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // cmbPercent
            // 
            this.cmbPercent.FormattingEnabled = true;
            this.cmbPercent.Items.AddRange(new object[] {
            "2%",
            "1%",
            "0.5%"});
            this.cmbPercent.Location = new System.Drawing.Point(490, 99);
            this.cmbPercent.Name = "cmbPercent";
            this.cmbPercent.Size = new System.Drawing.Size(96, 20);
            this.cmbPercent.TabIndex = 27;
            this.cmbPercent.SelectedIndexChanged += new System.EventHandler(this.cmbPercent_SelectedIndexChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(590, 245);
            this.Controls.Add(this.cmbPercent);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.txtStart);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnCalc);
            this.Controls.Add(this.txtPosn);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtTpsn);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtTpn);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnReload);
            this.Controls.Add(this.cmbJoblist);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtCurDir);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Jason Algorithm";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem newJobToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtCurDir;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox cmbJoblist;
        private System.Windows.Forms.Button btnReload;
        private System.Windows.Forms.TextBox txtTpn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtTpsn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtPosn;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnCalc;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripProgressBar StatusProgress;
        private System.Windows.Forms.ToolStripStatusLabel statusLabel1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.TextBox txtStart;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ComboBox cmbPercent;


    }
}

