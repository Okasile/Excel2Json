using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;

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
            bool isXml;
            try
            {
                 isXml= textBox1.Text.Substring(textBox1.Text.Length - 3, 3) == "xml";
            }
            catch
            {
                MessageBox.Show("excel 路径?");
                return;
            }
            string btnRawText = button3.Text;
            button3.Hide();
            //xml
            if (isXml)
            {
                try
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(textBox1.Text);
                    XmlNode root = doc.DocumentElement;
                    XmlNode tarSheet = null;
                    List<XmlNode> sheets = new List<XmlNode>();
                    sheets.Add(null); //从1开始 null 填充0
                    foreach (XmlNode node in root.ChildNodes)
                    {
                        if (node.Name == "Worksheet")
                        {
                            sheets.Add(node);
                        }
                    }
                    tarSheet = sheets[(int)numericUpDown1.Value];//Worksheet:table::raw::cell
                    XmlNode table = null;
                    foreach (XmlNode node in tarSheet)
                    {
                        if (node.Name == "Table")
                        {
                            table = node;
                            break;
                        }
                    }
                    int x = table.ChildNodes.Count;
                    int y = table.FirstChild.ChildNodes.Count;
                    string[][] content = new string[x][];
                    for (int i = 0; i < content.Length; i++)
                    {
                        content[i] = new string[y];
                    }
                    for (int i = 0; i < x; i++)
                    {
                        for (int j = 0; j < y; j++)
                        {
                            var v = table.ChildNodes[i].ChildNodes[j];
                            if (v != null)
                                content[i][j] = v.FirstChild.InnerText;
                        }
                    }
                    //json
                    string s = LitJson.JsonMapper.ToJson(content);
                    SaveToFile(s);
                }
                catch
                {
                    button3.Show();
                }
                return;
            }

            //xls

            Microsoft.Office.Interop.Excel.Application app = null;

            //open excel
            app = new Microsoft.Office.Interop.Excel.Application();
            if (app == null)
            {
                MessageBox.Show("转存成xml试试");
                return;
            }
            try
            {
                app.Visible = false;
                try
                {
                    app.Workbooks.Open(textBox1.Text);
                }
                catch
                {
                    MessageBox.Show("大哥,Excel路径是不是错了");
                    app.Quit();
                    button3.Show();
                    return;
                }
                //read excel
                _Worksheet ws = app.Sheets[(int)numericUpDown1.Value];
                int x = ws.UsedRange.Rows.Count;
                int y = ws.UsedRange.Columns.Count;
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
                SaveToFile(jsonContent);
            }
            catch
            {
                if (app != null)
                    app.Quit();
                button3.Show();
            }
            button3.Text = btnRawText;
        }

        void SaveToFile(string jsonContent)
        {
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
                button3.Show();
                return;
            }
            sw.Write(jsonContent);
            sw.Flush();
            sw.Close();
            var m = MessageBox.Show("已保存,是否打开查看?", "保存成功", MessageBoxButtons.YesNo);            
            if (m == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(textBox2.Text);
            }
            button3.Show();
        }

    }
}
