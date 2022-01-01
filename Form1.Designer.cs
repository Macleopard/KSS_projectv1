
namespace KSS_project
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
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
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.ex2w = new System.Windows.Forms.Button();
            this.w2pdf = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ex2w
            // 
            this.ex2w.Location = new System.Drawing.Point(47, 118);
            this.ex2w.Name = "ex2w";
            this.ex2w.Size = new System.Drawing.Size(100, 23);
            this.ex2w.TabIndex = 0;
            this.ex2w.Text = "Excel2Word";
            this.ex2w.UseVisualStyleBackColor = true;
            this.ex2w.Click += new System.EventHandler(this.ex2w_Click);
            // 
            // w2pdf
            // 
            this.w2pdf.Location = new System.Drawing.Point(233, 118);
            this.w2pdf.Name = "w2pdf";
            this.w2pdf.Size = new System.Drawing.Size(121, 23);
            this.w2pdf.TabIndex = 1;
            this.w2pdf.Text = "Word2SplitPDF";
            this.w2pdf.UseVisualStyleBackColor = true;
            this.w2pdf.Click += new System.EventHandler(this.w2pdf_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.w2pdf);
            this.Controls.Add(this.ex2w);
            this.Name = "Form1";
            this.Text = "KSS Project";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button ex2w;
        private System.Windows.Forms.Button w2pdf;
    }
}

