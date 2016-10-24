namespace MSWordDocument
{
    partial class Table
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
            this.AddTable = new System.Windows.Forms.Button();
            this.FillRows = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // AddTable
            // 
            this.AddTable.Location = new System.Drawing.Point(53, 52);
            this.AddTable.Name = "AddTable";
            this.AddTable.Size = new System.Drawing.Size(103, 29);
            this.AddTable.TabIndex = 0;
            this.AddTable.Text = "Add Table";
            this.AddTable.UseVisualStyleBackColor = true;
            this.AddTable.Click += new System.EventHandler(this.AddTable_Click);
            // 
            // FillRows
            // 
            this.FillRows.Location = new System.Drawing.Point(194, 55);
            this.FillRows.Name = "FillRows";
            this.FillRows.Size = new System.Drawing.Size(88, 23);
            this.FillRows.TabIndex = 1;
            this.FillRows.Text = "Fill Rows";
            this.FillRows.UseVisualStyleBackColor = true;
            this.FillRows.Click += new System.EventHandler(this.FillRows_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(390, 323);
            this.Controls.Add(this.FillRows);
            this.Controls.Add(this.AddTable);
            this.Name = "Form1";
            this.Text = "Table";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button AddTable;
        private System.Windows.Forms.Button FillRows;
    }
}

