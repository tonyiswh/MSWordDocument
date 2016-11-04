﻿namespace MSWordDocument
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
            this.TitleBookmarks = new System.Windows.Forms.Button();
            this.AddRow = new System.Windows.Forms.Button();
            this.CopyTable = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // AddTable
            // 
            this.AddTable.Location = new System.Drawing.Point(22, 49);
            this.AddTable.Name = "AddTable";
            this.AddTable.Size = new System.Drawing.Size(103, 29);
            this.AddTable.TabIndex = 0;
            this.AddTable.Text = "Add Table";
            this.AddTable.UseVisualStyleBackColor = true;
            this.AddTable.Click += new System.EventHandler(this.AddTable_Click);
            // 
            // FillRows
            // 
            this.FillRows.Location = new System.Drawing.Point(22, 175);
            this.FillRows.Name = "FillRows";
            this.FillRows.Size = new System.Drawing.Size(88, 23);
            this.FillRows.TabIndex = 1;
            this.FillRows.Text = "Fill Rows";
            this.FillRows.UseVisualStyleBackColor = true;
            this.FillRows.Click += new System.EventHandler(this.FillRows_Click);
            // 
            // TitleBookmarks
            // 
            this.TitleBookmarks.Location = new System.Drawing.Point(22, 124);
            this.TitleBookmarks.Name = "TitleBookmarks";
            this.TitleBookmarks.Size = new System.Drawing.Size(124, 23);
            this.TitleBookmarks.TabIndex = 2;
            this.TitleBookmarks.Text = "Add Title & Bookmarks";
            this.TitleBookmarks.UseVisualStyleBackColor = true;
            this.TitleBookmarks.Click += new System.EventHandler(this.TitleBookmarks_Click);
            // 
            // AddRow
            // 
            this.AddRow.Location = new System.Drawing.Point(22, 84);
            this.AddRow.Name = "AddRow";
            this.AddRow.Size = new System.Drawing.Size(75, 23);
            this.AddRow.TabIndex = 3;
            this.AddRow.Text = "Add Row";
            this.AddRow.UseVisualStyleBackColor = true;
            this.AddRow.Click += new System.EventHandler(this.AddRow_Click);
            // 
            // CopyTable
            // 
            this.CopyTable.Location = new System.Drawing.Point(22, 216);
            this.CopyTable.Name = "CopyTable";
            this.CopyTable.Size = new System.Drawing.Size(88, 23);
            this.CopyTable.TabIndex = 4;
            this.CopyTable.Text = "Copy Table";
            this.CopyTable.UseVisualStyleBackColor = true;
            this.CopyTable.Click += new System.EventHandler(this.CopyTable_Click);
            // 
            // Table
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(390, 323);
            this.Controls.Add(this.CopyTable);
            this.Controls.Add(this.AddRow);
            this.Controls.Add(this.TitleBookmarks);
            this.Controls.Add(this.FillRows);
            this.Controls.Add(this.AddTable);
            this.Name = "Table";
            this.Text = "Table";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button AddTable;
        private System.Windows.Forms.Button FillRows;
        private System.Windows.Forms.Button TitleBookmarks;
        private System.Windows.Forms.Button AddRow;
        private System.Windows.Forms.Button CopyTable;
    }
}

