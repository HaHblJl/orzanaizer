using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace трпо
{
    
    public partial class Form1 : Form
    {
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private readonly string excelSavePath = @"E:\трпо\TestSheet";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double a = 0;
            double b = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                a += Convert.ToInt32(dataGridView1[4,i].Value); 
            }
                textBox1.Text = a.ToString();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                b += Convert.ToInt32(dataGridView1[5, i].Value);
            }
            textBox2.Text = b.ToString();
            double sr = b - a;
            textBox3.Text = sr.ToString();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Deletebutton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                dataGridView1.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
                if (dataGridView1[0, i].FormattedValue.ToString() == SearchBox.Text)
                {
                    dataGridView1.CurrentCell = dataGridView1[5, i];
                    textBox4.Text = Convert.ToString( dataGridView1.CurrentCell.Value);
                    return;
                }
        }

        private void SearchBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            double average = 0; 
            double sum = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                
                sum += Convert.ToDouble( dataGridView1[5,i].Value);
                
            }
            average = sum / dataGridView1.Rows.Count;
            for (int i = 0; i < dataGridView1.RowCount-1; i++)
            {
                if (Convert.ToDouble(dataGridView1[5,i].Value) < average)
                {
                    dataGridView1.Rows.RemoveAt(i--);
                    
                }
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 newForm2 = new Form2(this);
            newForm2.Show();

            newForm2.dataGridView1.RowCount = dataGridView1.RowCount;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount ; j++)
                {
                    newForm2.dataGridView1[j,i].Value = dataGridView1[j,i].Value;
                }
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
          
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string sDate;
            string str;
            int rCnt;
            int cCnt;
            int str1;

            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Excel (*.XLSX)|*.XLSX";
            opf.ShowDialog();
            System.Data.DataTable tb = new System.Data.DataTable();
            string filename = opf.FileName;

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelRange;

            ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false,
                false, 0, true, 1, 0);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            ExcelRange = ExcelWorkSheet.UsedRange;
            for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)
            {
                dataGridView1.Rows.Add(1);
                for (cCnt = 1; cCnt <= 1; cCnt++)
                {
                    str = Convert.ToString((ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                    dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                }
                for (cCnt = 2; cCnt <= 2; cCnt++)
                {
                    str1 = Convert.ToInt32((ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                    dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str1;
                }
                
                for (cCnt = 3; cCnt <= 3; cCnt++)
                {
                    sDate = (ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                    //date = Convert.ToDouble(sDate);
                    DateTime dateTime = Convert.ToDateTime(DateTime.FromOADate(double.Parse(sDate)));
                    DateTime dat = dateTime.Date;
                    dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = dat;
                }
               
                for (cCnt = 4; cCnt <= 7; cCnt++)
                {
                    str1 = Convert.ToInt32((ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                    dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str1;
                }

            }
            ExcelWorkBook.Close(true, null, null);
            ExcelApp.Quit();

            releaseObject(ExcelWorkSheet);
            releaseObject(ExcelWorkBook);
            releaseObject(ExcelApp);
            
        }
        
    }

}
