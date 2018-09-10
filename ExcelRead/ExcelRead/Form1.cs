using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using System.IO;

namespace ExcelRead
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "选择要打开的Excel";
            dialog.Filter = "Excel文件|*.xlsx;*.xls;*.xml|所有文件|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = dialog.FileName;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.CheckFileExists = false;
            dialog.Title = "选择要保存";
            dialog.Filter = "保存的文件|*.txt|所有文件|*.*";
            dialog.FileName = "Excel2Json.txt";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = dialog.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = null;
            try
            {
                //open excel
                app = new Microsoft.Office.Interop.Excel.Application();
                if (app == null)
                {
                    MessageBox.Show("xiongdi ni mei zhuang excel ba?");                    
                    return;
                }
                app.Visible = false;
                try
                {
                    app.Workbooks.Open(textBox1.Text);
                }
                catch
                {
                    MessageBox.Show("大哥,Excel路径是不是错了");
                    app.Quit();
                    return;
                }
                //read excel
                _Worksheet ws = app.Sheets[(int)numericUpDown1.Value];
                int x = ws.UsedRange.Rows.Count;
                int y = ws.UsedRange.Columns.Count;                
                MessageBox.Show(string.Format("rows: {0}  columns:{1}", x, y));
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                string[][] content = new string[x][];
                for (int i = 0; i < x; i++)
                {
                    content[i] = new string[y];
                }
                for (int i = 0; i < x; i++)
                {
                    for (int j = 0; j < y; j++)
                    {
                        var v = ws.Cells[i + 1, j + 1];
                        if (v != null)
                            content[i][j] = ((Range)v).Text;
                    }
                }
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                app.Quit();
                string jsonContent = LitJson.JsonMapper.ToJson(content);
                StreamWriter sw = null;
                try
                {
                    sw = new StreamWriter(textBox2.Text, false);
                }
                catch
                {
                    MessageBox.Show("Save path error!");
                    if (sw != null)
                        sw.Close();
                    return;
                }
                sw.Write(jsonContent);
                sw.Flush();
                sw.Close();
            }
            catch
            {
                if (app != null)
                    app.Quit();
            }
        }
    }
}
