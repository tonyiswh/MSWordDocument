namespace MSWordDocument
{
    partial class PopupMenuDialog
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
            this.Dialog = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Dialog
            // 
            this.Dialog.Location = new System.Drawing.Point(56, 58);
            this.Dialog.Name = "Dialog";
            this.Dialog.Size = new System.Drawing.Size(75, 23);
            this.Dialog.TabIndex = 0;
            this.Dialog.Text = "Dialog";
            this.Dialog.UseVisualStyleBackColor = true;
            this.Dialog.Click += new System.EventHandler(this.Dialog_Click);
            // 
            // PopupMenuDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.Dialog);
            this.Name = "PopupMenuDialog";
            this.Text = "PopupMenuDialog";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Dialog;
    }
}