using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace JurisUtilityBase
{
    public partial class ReportDisplay : Form
    {
        public ReportDisplay(System.Data.DataSet ds)
        {
            InitializeComponent();
            dataGridView1.DataSource = ds.Tables[0];
        }

        private void buttonBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            string file = "";
            saveFileDialog1.Title = "Save Invoice file to location...";
            saveFileDialog1.DefaultExt = "csv";
            saveFileDialog1.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK && dataGridView1.Rows.Count > 0)
            {
                file = saveFileDialog1.FileName;
                var sb = new StringBuilder();

                var headers = dataGridView1.Columns.Cast<DataGridViewColumn>();
                sb.AppendLine(string.Join(",", headers.Select(column => "\"" + column.HeaderText + "\"").ToArray()));

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    var cells = row.Cells.Cast<DataGridViewCell>();
                    sb.AppendLine(string.Join(",", cells.Select(cell => "\"" + cell.Value + "\"").ToArray()));
                }
                File.WriteAllText(file, sb.ToString());
                MessageBox.Show("Finished Writing File!", "Success", MessageBoxButtons.OK, MessageBoxIcon.None);
                this.Close();
            }
            else
                MessageBox.Show("Empty");

        }

        


    }
}
