using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
namespace CreditCalculator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView1.ColumnCount = 5;
            dataGridView1.Columns[0].Name = "Місяць";
            dataGridView1.Columns[1].Name = "Платіж";
            dataGridView1.Columns[2].Name = "Тіло кредиту";
            dataGridView1.Columns[3].Name = "Процент";
            dataGridView1.Columns[4].Name = "Остача";
            dataGridView1.ColumnHeadersBorderStyle =
            DataGridViewHeaderBorderStyle.Single;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            textBox4.Text = "20000";
            textBox5.Text = "22";
            textBox6.Text = "4";
            dataGridView1.RowHeadersVisible = false;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox4.Text != String.Empty
                && textBox5.Text != String.Empty
                && textBox6.Text != String.Empty)
            {
                dataGridView1.Rows.Clear();
                Credit credit = new Credit(Convert.ToInt32(textBox4.Text), Convert.ToDouble(textBox5.Text), Convert.ToInt32(textBox6.Text));
                credit.CreateTable(ref dataGridView1);
                textBox3.Text = Convert.ToString(credit.AmountOfPayment);
                textBox2.Text = Convert.ToString(credit.overpayment);
                textBox1.Text = Convert.ToString(credit.finalPrice);
            }
            else
            {
                if(textBox4.Text != String.Empty
                && textBox5.Text != String.Empty
                && textBox6.Text != String.Empty)
                {
                    if (textBox4.Text == String.Empty)
                    {
                        MessageBox.Show("Введіть сумму кредтиту");
                    }
                    else if (textBox5.Text == String.Empty)
                    {
                        MessageBox.Show("Введіть процентну ставку");
                    }
                    else if (textBox6.Text == String.Empty)
                    {
                        MessageBox.Show("Введіть термін кредиту");
                    }
                }
                else
                {
                    MessageBox.Show("Заповніть поля");
                }
                
            }
        }
        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
