namespace RGolemAddin.View
{
    partial class Form5
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
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dateHourTo = new System.Windows.Forms.DateTimePicker();
            this.dateHourFrom = new System.Windows.Forms.DateTimePicker();
            this.dateTimeTo = new System.Windows.Forms.DateTimePicker();
            this.dateTimeFrom = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(253, 132);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 24);
            this.button1.TabIndex = 13;
            this.button1.Text = "Generuj";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dateHourTo);
            this.groupBox1.Controls.Add(this.dateHourFrom);
            this.groupBox1.Controls.Add(this.dateTimeTo);
            this.groupBox1.Controls.Add(this.dateTimeFrom);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(281, 114);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Zakres dat";
            // 
            // dateHourTo
            // 
            this.dateHourTo.CustomFormat = "HH";
            this.dateHourTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateHourTo.Location = new System.Drawing.Point(151, 76);
            this.dateHourTo.Name = "dateHourTo";
            this.dateHourTo.ShowUpDown = true;
            this.dateHourTo.Size = new System.Drawing.Size(66, 20);
            this.dateHourTo.TabIndex = 4;
            // 
            // dateHourFrom
            // 
            this.dateHourFrom.CustomFormat = "HH";
            this.dateHourFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateHourFrom.Location = new System.Drawing.Point(151, 37);
            this.dateHourFrom.Name = "dateHourFrom";
            this.dateHourFrom.ShowUpDown = true;
            this.dateHourFrom.Size = new System.Drawing.Size(66, 20);
            this.dateHourFrom.TabIndex = 2;
            // 
            // dateTimeTo
            // 
            this.dateTimeTo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimeTo.Location = new System.Drawing.Point(6, 76);
            this.dateTimeTo.MaxDate = new System.DateTime(2100, 1, 1, 0, 0, 0, 0);
            this.dateTimeTo.MinDate = new System.DateTime(2015, 1, 1, 0, 0, 0, 0);
            this.dateTimeTo.Name = "dateTimeTo";
            this.dateTimeTo.Size = new System.Drawing.Size(125, 20);
            this.dateTimeTo.TabIndex = 3;
            // 
            // dateTimeFrom
            // 
            this.dateTimeFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimeFrom.Location = new System.Drawing.Point(6, 37);
            this.dateTimeFrom.MaxDate = new System.DateTime(2100, 1, 1, 0, 0, 0, 0);
            this.dateTimeFrom.MinDate = new System.DateTime(2015, 1, 1, 0, 0, 0, 0);
            this.dateTimeFrom.Name = "dateTimeFrom";
            this.dateTimeFrom.Size = new System.Drawing.Size(125, 20);
            this.dateTimeFrom.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(21, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Do";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(21, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Od";
            // 
            // Form5
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(340, 162);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form5";
            this.Text = "Operatorzy";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker dateHourTo;
        private System.Windows.Forms.DateTimePicker dateHourFrom;
        private System.Windows.Forms.DateTimePicker dateTimeTo;
        private System.Windows.Forms.DateTimePicker dateTimeFrom;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}