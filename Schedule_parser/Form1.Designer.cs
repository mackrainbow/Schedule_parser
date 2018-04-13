namespace Schedule_parser
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
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.OpenFile_button = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.professorsComboBox = new System.Windows.Forms.ComboBox();
            this.DELETE_ME = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(17, 569);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 39);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(16, 15);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(248, 22);
            this.textBox1.TabIndex = 1;
            // 
            // OpenFile_button
            // 
            this.OpenFile_button.Location = new System.Drawing.Point(260, 15);
            this.OpenFile_button.Margin = new System.Windows.Forms.Padding(4);
            this.OpenFile_button.Name = "OpenFile_button";
            this.OpenFile_button.Size = new System.Drawing.Size(32, 25);
            this.OpenFile_button.TabIndex = 2;
            this.OpenFile_button.Text = "...";
            this.OpenFile_button.UseVisualStyleBackColor = true;
            this.OpenFile_button.Click += new System.EventHandler(this.OpenFile_button_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(17, 48);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(4);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(1246, 499);
            this.richTextBox1.TabIndex = 3;
            this.richTextBox1.Text = "";
            // 
            // professorsComboBox
            // 
            this.professorsComboBox.FormattingEnabled = true;
            this.professorsComboBox.Location = new System.Drawing.Point(347, 16);
            this.professorsComboBox.Name = "professorsComboBox";
            this.professorsComboBox.Size = new System.Drawing.Size(225, 24);
            this.professorsComboBox.TabIndex = 4;
            this.professorsComboBox.TextChanged += new System.EventHandler(this.professorsComboBox_TextChanged);
            // 
            // DELETE_ME
            // 
            this.DELETE_ME.Location = new System.Drawing.Point(863, 569);
            this.DELETE_ME.Name = "DELETE_ME";
            this.DELETE_ME.Size = new System.Drawing.Size(75, 23);
            this.DELETE_ME.TabIndex = 5;
            this.DELETE_ME.Text = "button2";
            this.DELETE_ME.UseVisualStyleBackColor = true;
            this.DELETE_ME.Click += new System.EventHandler(this.DELETE_ME_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1276, 623);
            this.Controls.Add(this.DELETE_ME);
            this.Controls.Add(this.professorsComboBox);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.OpenFile_button);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button OpenFile_button;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ComboBox professorsComboBox;
        private System.Windows.Forms.Button DELETE_ME;
    }
}

