namespace Maestro_PC
{
    partial class frmMain
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
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.cbxReverse = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // cbxReverse
            // 
            this.cbxReverse.AutoSize = true;
            this.cbxReverse.Location = new System.Drawing.Point(83, 369);
            this.cbxReverse.Name = "cbxReverse";
            this.cbxReverse.Size = new System.Drawing.Size(101, 22);
            this.cbxReverse.TabIndex = 0;
            this.cbxReverse.Text = "Reverse";
            this.cbxReverse.UseVisualStyleBackColor = true;
            this.cbxReverse.CheckedChanged += new System.EventHandler(this.cbxReverse_CheckedChanged);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(300, 300);
            this.Controls.Add(this.cbxReverse);
            this.Name = "frmMain";
            this.Text = "Maestro";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmMain_FormClosed);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox cbxReverse;
    }
}

