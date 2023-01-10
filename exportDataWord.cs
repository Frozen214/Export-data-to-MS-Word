using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace exportWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of Word and make it visible
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;

            // Add a new document
            var wordDoc = wordApp.Documents.Add();

            // Insert a table with data from the DataGridView
            InsertTableFromDataGridView(wordDoc, товарDataGridView);
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "анализПоведенияDataSet.Товар". При необходимости она может быть перемещена или удалена.
            this.товарTableAdapter.Fill(this.анализПоведенияDataSet.Товар);
        }
        private void InsertTableFromDataGridView(Document doc, DataGridView dgv)
        {
            // Create a new table with the same number of columns as the DataGridView
            var table = doc.Tables.Add(doc.Range(), dgv.Rows.Count + 1, dgv.Columns.Count);

            // Copy the column headers from the DataGridView
            for (int j = 0; j < dgv.Columns.Count; j++)
            {
                table.Rows[1].Cells[j + 1].Range.Text = dgv.Columns[j].HeaderText;
            }

            // Copy the data from the DataGridView
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    if (dgv[j, i].Value != null)
                    {
                        table.Rows[i + 2].Cells[j + 1].Range.Text = dgv[j, i].Value.ToString();
                    }
                }
            }

            // Add formatting to the table
            table.Rows[1].Range.Font.Bold = 1;
            table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            table.Range.ParagraphFormat.SpaceAfter = 6;
        }

    }
}
