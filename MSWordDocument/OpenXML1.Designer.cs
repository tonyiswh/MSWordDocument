namespace MSWordDocument
{
    partial class OpenXML1
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
            this.Update = new System.Windows.Forms.Button();
            this.TableOne = new System.Windows.Forms.Button();
            this.ColumnMerge = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Update
            // 
            this.Update.Location = new System.Drawing.Point(12, 22);
            this.Update.Name = "Update";
            this.Update.Size = new System.Drawing.Size(114, 23);
            this.Update.TabIndex = 0;
            this.Update.Text = "Update";
            this.Update.UseVisualStyleBackColor = true;
            this.Update.Click += new System.EventHandler(this.Update_Click);
            // 
            // TableOne
            // 
            this.TableOne.Location = new System.Drawing.Point(12, 62);
            this.TableOne.Name = "TableOne";
            this.TableOne.Size = new System.Drawing.Size(114, 23);
            this.TableOne.TabIndex = 1;
            this.TableOne.Text = "Table One";
            this.TableOne.UseVisualStyleBackColor = true;
            this.TableOne.Click += new System.EventHandler(this.TableOne_Click);
            // 
            // ColumnMerge
            // 
            this.ColumnMerge.Location = new System.Drawing.Point(12, 100);
            this.ColumnMerge.Name = "ColumnMerge";
            this.ColumnMerge.Size = new System.Drawing.Size(114, 23);
            this.ColumnMerge.TabIndex = 2;
            this.ColumnMerge.Text = "Table Column Merge";
            this.ColumnMerge.UseVisualStyleBackColor = true;
            this.ColumnMerge.Click += new System.EventHandler(this.ColumnMerge_Click);
            // 
            // OpenXML1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 273);
            this.Controls.Add(this.ColumnMerge);
            this.Controls.Add(this.TableOne);
            this.Controls.Add(this.Update);
            this.Name = "OpenXML1";
            this.Text = "OpenXML";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.OpenXML_FormClosed);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Update;
        private System.Windows.Forms.Button TableOne;
        private System.Windows.Forms.Button ColumnMerge;
    }
}