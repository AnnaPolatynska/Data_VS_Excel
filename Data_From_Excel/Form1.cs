using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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


    }//class Form1 : Form

}//namespace Data_From_Excel
