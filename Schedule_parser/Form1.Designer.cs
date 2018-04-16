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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.OpenFile_button = new System.Windows.Forms.Button();
            this.professorsComboBox = new System.Windows.Forms.ComboBox();
            this.SaveFileButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // OpenFile_button
            // 
            this.OpenFile_button.Location = new System.Drawing.Point(13, 13);
            this.OpenFile_button.Margin = new System.Windows.Forms.Padding(4);
            this.OpenFile_button.Name = "OpenFile_button";
            this.OpenFile_button.Size = new System.Drawing.Size(225, 46);
            this.OpenFile_button.TabIndex = 2;
            this.OpenFile_button.Text = "Выберите файлы с расписаниями";
            this.OpenFile_button.UseVisualStyleBackColor = true;
            this.OpenFile_button.Click += new System.EventHandler(this.OpenFile_button_Click);
            // 
            // professorsComboBox
            // 
            this.professorsComboBox.FormattingEnabled = true;
            this.professorsComboBox.Location = new System.Drawing.Point(13, 66);
            this.professorsComboBox.Name = "professorsComboBox";
            this.professorsComboBox.Size = new System.Drawing.Size(225, 24);
            this.professorsComboBox.TabIndex = 4;
            this.professorsComboBox.TextChanged += new System.EventHandler(this.professorsComboBox_TextChanged);
            // 
            // SaveFileButton
            // 
            this.SaveFileButton.Enabled = false;
            this.SaveFileButton.Location = new System.Drawing.Point(12, 289);
            this.SaveFileButton.Name = "SaveFileButton";
            this.SaveFileButton.Size = new System.Drawing.Size(225, 46);
            this.SaveFileButton.TabIndex = 5;
            this.SaveFileButton.Text = "Сохранить результат";
            this.SaveFileButton.UseVisualStyleBackColor = true;
            this.SaveFileButton.Click += new System.EventHandler(this.SaveFileButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(251, 347);
            this.Controls.Add(this.SaveFileButton);
            this.Controls.Add(this.professorsComboBox);
            this.Controls.Add(this.OpenFile_button);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Расписание преподавателя";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button OpenFile_button;
        private System.Windows.Forms.ComboBox professorsComboBox;
        private System.Windows.Forms.Button SaveFileButton;
    }
}

