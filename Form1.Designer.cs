namespace OpenXMLPractice
{
    partial class Form1
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
            this.cmdCreateNew = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cmdCreateNew
            // 
            this.cmdCreateNew.AutoSize = true;
            this.cmdCreateNew.Location = new System.Drawing.Point(73, 97);
            this.cmdCreateNew.Name = "cmdCreateNew";
            this.cmdCreateNew.Size = new System.Drawing.Size(122, 23);
            this.cmdCreateNew.TabIndex = 0;
            this.cmdCreateNew.Text = "Create New Word File";
            this.cmdCreateNew.UseVisualStyleBackColor = true;
            this.cmdCreateNew.Click += new System.EventHandler(this.cmdCreateNew_Click);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.cmdCreateNew);
            this.Name = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button cmdCreateNew;

    }
}

