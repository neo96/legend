legend
======

its overwriting the data instead of writing it on new line
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace legend1._0
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //int _lastRow = ExcelUtility.WorkSheet.Range["A" +                 ExcelUtility.WorkSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row + 1;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.ColumnCount = ExcelUtility.ColumnCount;

            /* for(int x=0;x<dataGridView1.ColumnCount;x++)
             {
                     dataGridView1.Columns[x].Name = "Column "+x.ToString();
             }*/

            for (int i = 0; i < ExcelUtility.RowCount; i++)
            {
                dataGridView1.Rows.Add(ExcelUtility.WorkSheet.Cells[i + 1, 1].Value, ExcelUtility.WorkSheet.Cells[i + 1, 2].Value,
                ExcelUtility.WorkSheet.Cells[i + 1, 3].Value, ExcelUtility.WorkSheet.Cells[i + 1, 4].Value,
                ExcelUtility.WorkSheet.Cells[i + 1, 5].Value, ExcelUtility.WorkSheet.Cells[i + 1, 7].Value,
                ExcelUtility.WorkSheet.Cells[i + 1, 9].Value, ExcelUtility.WorkSheet.Cells[i + 1, 11].Value,
                ExcelUtility.WorkSheet.Cells[i + 1, 12].Value, ExcelUtility.WorkSheet.Cells[i + 1, 13].Value,
                ExcelUtility.WorkSheet.Cells[i + 1, 14].Value,ExcelUtility.WorkSheet.Cells[i + 1, 15].Value);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int _lastRow = ExcelUtility.WorkSheet.Range["A" + ExcelUtility.WorkSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row + 1;
            // _lastRow++;
            ExcelUtility.WorkSheet.Cells[_lastRow, 2] = textBox2.Text;
            ExcelUtility.WorkSheet.Cells[_lastRow, 3] = textBox3.Text;
            ExcelUtility.WorkSheet.Cells[_lastRow, 4] = textBox4.Text;
            if (radioButton1.Checked == true)
            {
                ExcelUtility.WorkSheet.Cells[_lastRow, 5] = p.Text;
            }
            else if(radioButton2.Checked == true)
            {
                ExcelUtility.WorkSheet.Cells[_lastRow, 7] = p.Text;
            }
            else if (radioButton3.Checked == true)
            {
                ExcelUtility.WorkSheet.Cells[_lastRow, 9] = p.Text;
            }
            else if (radioButton4.Checked == true)
            {
                ExcelUtility.WorkSheet.Cells[_lastRow, 11] = p.Text;
            }
            ExcelUtility.WorkSheet.Cells[_lastRow, 12] = textBox5.Text;
            ExcelUtility.WorkSheet.Cells[_lastRow, 13] = textBox6.Text;
            ExcelUtility.WorkSheet.Cells[_lastRow, 14] = textBox7.Text;
            ExcelUtility.WorkSheet.Cells[_lastRow, 15] = textBox8.Text;
            
        }

        private void filter(string text)
        {
           
            //int _lastRow = ExcelUtility.WorkSheet.Range["A" + ExcelUtility.WorkSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row + 1;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.ColumnCount = ExcelUtility.ColumnCount;
           
            for (int i = 0; i < ExcelUtility.RowCount; i++)
            {
                if (text == ExcelUtility.WorkSheet.Cells[i + 1, 15].Value)
                {
                    dataGridView1.Rows.Add(ExcelUtility.WorkSheet.Cells[i + 1, 1].Value, ExcelUtility.WorkSheet.Cells[i + 1, 2].Value,
                    ExcelUtility.WorkSheet.Cells[i + 1, 3].Value, ExcelUtility.WorkSheet.Cells[i + 1, 4].Value,
                    ExcelUtility.WorkSheet.Cells[i + 1, 5].Value, ExcelUtility.WorkSheet.Cells[i + 1, 7].Value,
                    ExcelUtility.WorkSheet.Cells[i + 1, 9].Value, ExcelUtility.WorkSheet.Cells[i + 1, 11].Value,
                    ExcelUtility.WorkSheet.Cells[i + 1, 12].Value, ExcelUtility.WorkSheet.Cells[i + 1, 13].Value,
                    ExcelUtility.WorkSheet.Cells[i + 1, 14].Value, ExcelUtility.WorkSheet.Cells[i + 1, 15].Value);
                }
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string text = comboBox1.Text;
            filter(text);
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
