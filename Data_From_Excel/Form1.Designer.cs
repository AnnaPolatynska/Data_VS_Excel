namespace Data_From_Excel
{
    partial class Form1
    {
        /// <summary>
        /// Wymagana zmienna projektanta.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Wyczyść wszystkie używane zasoby.
        /// </summary>
        /// <param name="disposing">prawda, jeżeli zarządzane zasoby powinny zostać zlikwidowane; Fałsz w przeciwnym wypadku.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Kod generowany przez Projektanta formularzy systemu Windows

        /// <summary>
        /// Metoda wymagana do obsługi projektanta — nie należy modyfikować
        /// jej zawartości w edytorze kodu.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label13 = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.buttonLoad = new System.Windows.Forms.Button();
            this.groupBoxWprowadzanieDanych = new System.Windows.Forms.GroupBox();
            this.textBoxTelephone = new System.Windows.Forms.TextBox();
            this.textBoxEmail = new System.Windows.Forms.TextBox();
            this.textBoxSurname = new System.Windows.Forms.TextBox();
            this.textBoxName = new System.Windows.Forms.TextBox();
            this.buttonAdd = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dataGridEmptyList = new System.Windows.Forms.DataGridView();
            this.applicationBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.buttonNewFile = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBoxWprowadzanieDanych.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridEmptyList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.applicationBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.tabControl1.Location = new System.Drawing.Point(13, 13);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(775, 425);
            this.tabControl1.TabIndex = 1;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabPage1_Click);
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.DarkKhaki;
            this.tabPage1.Controls.Add(this.buttonNewFile);
            this.tabPage1.Controls.Add(this.label13);
            this.tabPage1.Controls.Add(this.txtFileName);
            this.tabPage1.Controls.Add(this.buttonLoad);
            this.tabPage1.Controls.Add(this.groupBoxWprowadzanieDanych);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(767, 396);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Dodawanie rekordu";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.label13.Location = new System.Drawing.Point(6, 32);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(283, 26);
            this.label13.TabIndex = 12;
            this.label13.Text = "Ścieżka dostępu do excela: ";
            // 
            // txtFileName
            // 
            this.txtFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.txtFileName.Location = new System.Drawing.Point(295, 33);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(348, 26);
            this.txtFileName.TabIndex = 11;
            // 
            // buttonLoad
            // 
            this.buttonLoad.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.buttonLoad.Location = new System.Drawing.Point(649, 28);
            this.buttonLoad.Name = "buttonLoad";
            this.buttonLoad.Size = new System.Drawing.Size(104, 37);
            this.buttonLoad.TabIndex = 10;
            this.buttonLoad.Text = "Otwórz";
            this.buttonLoad.UseVisualStyleBackColor = true;
            this.buttonLoad.Click += new System.EventHandler(this.buttonLoad_Click);
            // 
            // groupBoxWprowadzanieDanych
            // 
            this.groupBoxWprowadzanieDanych.BackColor = System.Drawing.Color.OliveDrab;
            this.groupBoxWprowadzanieDanych.Controls.Add(this.textBoxTelephone);
            this.groupBoxWprowadzanieDanych.Controls.Add(this.textBoxEmail);
            this.groupBoxWprowadzanieDanych.Controls.Add(this.textBoxSurname);
            this.groupBoxWprowadzanieDanych.Controls.Add(this.textBoxName);
            this.groupBoxWprowadzanieDanych.Controls.Add(this.buttonAdd);
            this.groupBoxWprowadzanieDanych.Controls.Add(this.label4);
            this.groupBoxWprowadzanieDanych.Controls.Add(this.label3);
            this.groupBoxWprowadzanieDanych.Controls.Add(this.label1);
            this.groupBoxWprowadzanieDanych.Controls.Add(this.label2);
            this.groupBoxWprowadzanieDanych.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.groupBoxWprowadzanieDanych.Location = new System.Drawing.Point(9, 158);
            this.groupBoxWprowadzanieDanych.Name = "groupBoxWprowadzanieDanych";
            this.groupBoxWprowadzanieDanych.Size = new System.Drawing.Size(755, 232);
            this.groupBoxWprowadzanieDanych.TabIndex = 13;
            this.groupBoxWprowadzanieDanych.TabStop = false;
            this.groupBoxWprowadzanieDanych.Text = "Wprowadzanie danych";
            // 
            // textBoxTelephone
            // 
            this.textBoxTelephone.Location = new System.Drawing.Point(221, 183);
            this.textBoxTelephone.Name = "textBoxTelephone";
            this.textBoxTelephone.Size = new System.Drawing.Size(307, 32);
            this.textBoxTelephone.TabIndex = 12;
            // 
            // textBoxEmail
            // 
            this.textBoxEmail.Location = new System.Drawing.Point(221, 131);
            this.textBoxEmail.Name = "textBoxEmail";
            this.textBoxEmail.Size = new System.Drawing.Size(307, 32);
            this.textBoxEmail.TabIndex = 11;
            // 
            // textBoxSurname
            // 
            this.textBoxSurname.Location = new System.Drawing.Point(221, 83);
            this.textBoxSurname.Name = "textBoxSurname";
            this.textBoxSurname.Size = new System.Drawing.Size(307, 32);
            this.textBoxSurname.TabIndex = 10;
            // 
            // textBoxName
            // 
            this.textBoxName.Location = new System.Drawing.Point(221, 31);
            this.textBoxName.Name = "textBoxName";
            this.textBoxName.Size = new System.Drawing.Size(307, 32);
            this.textBoxName.TabIndex = 9;
            // 
            // buttonAdd
            // 
            this.buttonAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.buttonAdd.Location = new System.Drawing.Point(608, 176);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(136, 39);
            this.buttonAdd.TabIndex = 8;
            this.buttonAdd.Text = "Dodaj rekord";
            this.buttonAdd.UseVisualStyleBackColor = true;
            this.buttonAdd.Click += new System.EventHandler(this.buttonAdd_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.label4.Location = new System.Drawing.Point(14, 186);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 26);
            this.label4.TabIndex = 7;
            this.label4.Text = "Telefon";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.label3.Location = new System.Drawing.Point(14, 134);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(68, 26);
            this.label3.TabIndex = 6;
            this.label3.Text = "Email";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.label1.Location = new System.Drawing.Point(14, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 26);
            this.label1.TabIndex = 4;
            this.label1.Text = "Imie";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.label2.Location = new System.Drawing.Point(14, 86);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 26);
            this.label2.TabIndex = 5;
            this.label2.Text = "Nazwisko";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.Olive;
            this.tabPage2.Controls.Add(this.dataGridEmptyList);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(767, 396);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Szczegóły rekordu";
            // 
            // dataGridEmptyList
            // 
            this.dataGridEmptyList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridEmptyList.Location = new System.Drawing.Point(7, 7);
            this.dataGridEmptyList.Name = "dataGridEmptyList";
            this.dataGridEmptyList.Size = new System.Drawing.Size(754, 383);
            this.dataGridEmptyList.TabIndex = 0;
            // 
            // applicationBindingSource
            // 
            this.applicationBindingSource.DataSource = typeof(Microsoft.Office.Interop.Excel._Application);
            // 
            // buttonNewFile
            // 
            this.buttonNewFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.buttonNewFile.Location = new System.Drawing.Point(617, 98);
            this.buttonNewFile.Name = "buttonNewFile";
            this.buttonNewFile.Size = new System.Drawing.Size(136, 38);
            this.buttonNewFile.TabIndex = 14;
            this.buttonNewFile.Text = "Utwórz nowy";
            this.buttonNewFile.UseVisualStyleBackColor = true;
            this.buttonNewFile.Click += new System.EventHandler(this.buttonNewFile_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBoxWprowadzanieDanych.ResumeLayout(false);
            this.groupBoxWprowadzanieDanych.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridEmptyList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.applicationBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Button buttonLoad;
        private System.Windows.Forms.GroupBox groupBoxWprowadzanieDanych;
        private System.Windows.Forms.TextBox textBoxTelephone;
        private System.Windows.Forms.TextBox textBoxEmail;
        private System.Windows.Forms.TextBox textBoxSurname;
        private System.Windows.Forms.TextBox textBoxName;
        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView dataGridEmptyList;
        private System.Windows.Forms.BindingSource applicationBindingSource;
        private System.Windows.Forms.Button buttonNewFile;
    }
}

