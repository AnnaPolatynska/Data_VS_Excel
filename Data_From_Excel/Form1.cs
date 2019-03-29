using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Data_From_Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tabControl1.Selecting += new TabControlCancelEventHandler(TabControl1_Selecting);
        }//Form1

        /// <summary>
        /// Obsługa przejścia pomiędzy zakładkami.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabControl1_Selecting(object sender, TabControlCancelEventArgs eventArgs)
        {
            TabPage current = (sender as TabControl).SelectedTab;

            if (string.IsNullOrEmpty(MyExcel.DB_PATH))
            {
                MessageBox.Show("Proszę wskazać ścieżkę dostępu do pliku Excel.", "Błąd!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                eventArgs.Cancel = true;
            }
        }// TabControl1_Selecting

        /// <summary>
        /// Przejście pomiędzy zakładkami.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabPage1_Click(object sender, EventArgs e)
        {
            TabControl tc = sender as TabControl;
            if (tc.SelectedIndex == 1)
            {
                dataGridEmptyList.DataSource = (BindingList<Kontakt>)MyExcel.EmptyList;
                dataGridEmptyList.AutoResizeColumns();
            }
        }//tabPage1_Click

        /// <summary>
        /// Obsługa kliknięcia dodania rekordu.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            Kontakt kontakt = new Kontakt
            {
                Imie = textBoxName.Text.ToString(),
                Nazwisko = textBoxSurname.Text.ToString(),
                Email = textBoxEmail.Text.ToString(),
                Telefon = textBoxTelephone.Text.ToString(),
                Dane = richTextData.Text.ToString()
            };

            MyExcel.WriteToExcel(kontakt);
            clearAllFields();
            MessageBox.Show("Dane rekordu zostały pomyślnie dodane do programu Excel.", "OK!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            textBoxName.Focus();
            MyExcel.CloseExcel();
            // MyBook.Close();
        }//buttonAdd_Click

        private void clearAllFields()
        {
            textBoxName.Text = "";
            textBoxSurname.Text = "";
            textBoxEmail.Text = "";
            textBoxTelephone.Text = "";
            richTextData.Text = "";
        }//clearAllFields

        private void buttonLoad_Click(object sender, EventArgs e)
        {
            existingExcel();
        }//buttonLoad_Click
       
        /// <summary>
        /// Obsługa istniejącego pliku excel.
        /// </summary>
        private void existingExcel()
        {
            OpenFileDialog ExcelDialog = new OpenFileDialog();
            ExcelDialog.Filter = "Excel Files (*.xlsx) | *.xlsx";
            ExcelDialog.InitialDirectory = @"D:\Projects\Data_From_Excel\Excele";
            ExcelDialog.Title = "Wybierz swój plik excel z Kontaktami";

            if (ExcelDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                MyExcel.DB_PATH = ExcelDialog.FileName;
                txtFileName.Text = ExcelDialog.FileName;
                txtFileName.ReadOnly = true;
                txtFileName.Click -= buttonLoad_Click;
                tabControl1.Selecting -= TabControl1_Selecting;
                //buttonLoad.Enabled = true;
                MyExcel.InitializeExcel();
                dataGridEmptyList.DataSource = MyExcel.ReadMyExcel();
            }
        }//existingExcel()

        /// <summary>
        /// Obsługa tworzenia pustego pliku excel.
        /// </summary>
        private void nowyPlik()
        {
            Stream myStream;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "Excel Files (*.xlsx) | *.xlsx";
            saveFileDialog1.InitialDirectory = @"D:\Projects\Data_From_Excel\Excele";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((myStream = saveFileDialog1.OpenFile()) != null)
                {
                    //using (var myStream = saveFileDialog1.OpenFile())
                    using (var writer = new StreamWriter(myStream))
                    {
                        MyExcel.DB_PATH = saveFileDialog1.FileName;
                        //txtFileName.Text = saveFileDialog1.FileName;
                        //txtFileName.ReadOnly = true;
                        //txtFileName.Click -= buttonLoad_Click;
                        tabControl1.Selecting -= TabControl1_Selecting;
                        MyExcel.InitializeExcel();
                        dataGridEmptyList.DataSource = MyExcel.ReadMyExcel();
                    }
                    myStream.Close();
                }
            }
        }// nowyPlik()

        /// <summary>
        /// Obsługa kliknięcia utworzenia nowego pliku excel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonNewFile_Click(object sender, EventArgs e)
        {
            nowyPlik();
        }//buttonNewFile_Click

        private void buttonExcel_Click(object sender, EventArgs e)
        {
            DirectoryInfo outputDir = new DirectoryInfo(@"D:\Projects\Data_From_Excel\Excele");

            if (!outputDir.Exists)
                outputDir.Create();

            CreateNewExcel<Kontakt>.ExportToExcel(MyExcel.ReadMyExcel(), outputDir, "Kontakty");

            Console.WriteLine("Utworzono plik Excel ");
            // Console.ReadKey();
        }
    }//class Form1 : Form

}//namespace Data_From_Excel
