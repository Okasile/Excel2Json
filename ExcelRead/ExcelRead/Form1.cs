using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using ExcelDataReader;

namespace ExcelRead
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            LoadPaths();
        }

        void RememberPaths()
        {
            PathsConfiguration ps = new PathsConfiguration();
            ps.excel1 = pathExcel1.Text;
            ps.excel2 = pathExcel2.Text;
            ps.saveFile1 = pathFileSave1.Text;
            ps.saveFile2 = pathFileSave2.Text;
            ps.removeHead1 = RemoveHeadLine1.Checked;
            ps.removeHead2 = RemoveHeadLine2.Checked;
            ps.sheet1 = (int)sheet1.Value;
            ps.sheet2 = (int)sheet2.Value;

            string confPath = System.Environment.CurrentDirectory + "Config.txt";

            using (StreamWriter sw = new StreamWriter(confPath, false))
            {
                sw.Write(LitJson.JsonMapper.ToJson(ps));
            }
        }

        void LoadPaths()
        {

            string confPath = System.Environment.CurrentDirectory + "Config.txt";
            if (!File.Exists(confPath))
            {
                return;
            }

            string jsonContent = string.Empty;
            using (StreamReader sr = new StreamReader(confPath))
            {
                jsonContent = sr.ReadToEnd();
            }
            if (jsonContent != string.Empty)
            {
                PathsConfiguration ps = LitJson.JsonMapper.ToObject<PathsConfiguration>(jsonContent);
                pathExcel1.Text = ps.excel1;
                pathExcel2.Text = ps.excel2;
                pathFileSave1.Text = ps.saveFile1;
                pathFileSave2.Text = ps.saveFile2;
                RemoveHeadLine1.Checked = ps.removeHead1;
                RemoveHeadLine2.Checked = ps.removeHead2;
                sheet1.Value = ps.sheet1;
                sheet2.Value = ps.sheet2;
            }
        }


        #region 按钮
        private void excelPathFindBtn1_Click(object sender, EventArgs e)
        {
            ClickBtn2SetPath(pathExcel1, true);
        }

        private void fileSaveFindBtn1_Click(object sender, EventArgs e)
        {
            ClickBtn2SetPath(pathFileSave1, false);
        }


        private void excelPathFindBtn2_Click(object sender, EventArgs e)
        {
            ClickBtn2SetPath(pathExcel2, true);
        }

        private void fileSaveFindBtn2_Click(object sender, EventArgs e)
        {
            ClickBtn2SetPath(pathFileSave2, false);
        }


        private void sheet2_ValueChanged(object sender, EventArgs e)
        {
            RememberPaths();
        }

        private void sheet1_ValueChanged(object sender, EventArgs e)
        {
            RememberPaths();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            ReadAndSave(pathExcel1.Text, (int)sheet1.Value, pathFileSave1.Text, RemoveHeadLine1.Checked);
            ReadAndSave(pathExcel2.Text, (int)sheet2.Value, pathFileSave2.Text, RemoveHeadLine2.Checked);
            MessageBox.Show("完成");
        }


        #endregion

        void ClickBtn2SetPath(System.Windows.Forms.TextBox setTar, bool isExcel)
        {
            if (isExcel)
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = false;
                dialog.Title = "选择要打开的Excel";
                dialog.Filter = "Excel文件|*.xlsx;*.xls;*.xml|所有文件|*.*";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    setTar.Text = dialog.FileName;
                }
            }
            else
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = false;
                dialog.CheckFileExists = false;
                dialog.Title = "选择要保存";
                dialog.Filter = "保存的文件|*.txt|所有文件|*.*";
                dialog.FileName = "Excel2Json.txt";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    setTar.Text = dialog.FileName;
                }
            }
            RememberPaths();
        }

        void ReadAndSave(string pathExcel, int sheetPage, string pathSaveFile, bool reamoveHead)
        {
            string btnRawText = saveBtn.Text;
            #region xml 不写了
            //             bool isXml;
            //             button3.Hide();
            //             //xml
            //             if (isXml)
            //             {
            //                 try
            //                 {
            //                     XmlDocument doc = new XmlDocument();
            //                     doc.Load(textBox1.Text);
            //                     XmlNode root = doc.DocumentElement;
            //                     XmlNode tarSheet = null;
            //                     List<XmlNode> sheets = new List<XmlNode>();
            //                     sheets.Add(null); //从1开始 null 填充0
            //                     foreach (XmlNode node in root.ChildNodes)
            //                     {
            //                         if (node.Name == "Worksheet")
            //                         {
            //                             sheets.Add(node);
            //                         }
            //                     }
            //                     tarSheet = sheets[(int)numericUpDown1.Value];//Worksheet:table::raw::cell
            //                     XmlNode table = null;
            //                     foreach (XmlNode node in tarSheet)
            //                     {
            //                         if (node.Name == "Table")
            //                         {
            //                             table = node;
            //                             break;
            //                         }
            //                     }
            //                     int x = table.ChildNodes.Count;
            //                     int y = table.FirstChild.ChildNodes.Count;
            //                     string[][] content = new string[x][];
            //                     for (int i = 0; i < content.Length; i++)
            //                     {                        
            //                         content[i] = new string[y];
            //                     }
            //                     for (int i = 0; i < x; i++)
            //                     {
            //                         for (int j = 0; j < y; j++)
            //                         {
            //                             var v = table.ChildNodes[i].ChildNodes[j];
            //                             if (v != null)
            //                                 content[i][j] = v.FirstChild.InnerText;
            //                         }
            //                     }
            //                     //json                    
            //                     if (RemoveLine1.Checked)
            //                     {
            //                         var result = new string[content.Length - 1][];
            //                         for(int i=0;i<result.Length;i++)
            //                         {
            //                             result[i] = content[i + 1];
            //                         }
            //                         content = result;
            //                     }
            //                     string s = LitJson.JsonMapper.ToJson(content);
            //                     SaveToFile(s);
            //                 }
            //                 catch
            //                 {
            //                     button3.Show();
            //                 }
            //                 return;
            //             }
            #endregion
            //xls

            if (!File.Exists(pathExcel))
                return;

            saveBtn.Hide();
            using (FileStream fs = File.Open(pathExcel, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fs))
                {
                    var result = reader.AsDataSet();
                    var table = result.Tables[sheetPage - 1];

                    List<List<object>> datas = new List<List<object>>();
                    if (table.Rows.Count <= 0 || table.Rows.Count <= 1 && reamoveHead)
                        return;
                    for (int i = reamoveHead? 1:0; i < table.Rows.Count; i++)
                    {
                        List<object> temp = new List<object>();
                        for(int c = 0;c<table.Columns.Count;c++)
                        {
                            var v = table.Rows[i][c];
                            if(v is double)
                            {
                                v =Convert.ToSingle(v);
                                if((Convert.ToInt32(v)) == Convert.ToSingle(v))
                                {
                                    v = Convert.ToInt32(v);
                                }
                            }                            
                            if (v == DBNull.Value)
                                v = string.Empty;
                            temp.Add(v);
                        }
                        datas.Add(temp);
                    }
                    string jsonContent = LitJson.JsonMapper.ToJson(datas);
                    SaveToFile(jsonContent, pathSaveFile);
                }
            }
            saveBtn.Show();

            #region 用的Microsoft.Office.Interop.Excel.Application,已经废弃
            //             Microsoft.Office.Interop.Excel.Application app = null;
            // 
            //             //open excel
            //             app = new Microsoft.Office.Interop.Excel.Application();
            //             try
            //             {
            //                 app.Visible = false;
            //                 try
            //                 {
            //                     app.Workbooks.Open(pathExcel);
            //                 }
            //                 catch
            //                 {
            //                     MessageBox.Show("大哥,Excel路径是不是错了");
            //                     app.Quit();
            //                     saveBtn.Show();
            //                     return;
            //                 }
            //                 //read excel
            //                 _Worksheet ws = app.Sheets[sheetPage];
            //                 int x = ws.UsedRange.Rows.Count;
            //                 int y = ws.UsedRange.Columns.Count;
            //                 ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                 string[][] content = new string[x][];
            //                 for (int i = 0; i < x; i++)
            //                 {
            //                     content[i] = new string[y];
            //                 }
            //                 for (int i = 0; i < x; i++)
            //                 {
            //                     for (int j = 0; j < y; j++)
            //                     {
            //                         var v = ws.Cells[i + 1, j + 1];
            //                         if (v != null)
            //                             content[i][j] = ((Range)v).Text;
            //                     }
            //                 }
            //                 ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                 app.Quit();
            //                 if (reamveHead)
            //                 {
            //                     var result = new string[content.Length - 1][];
            //                     for (int i = 0; i < result.Length; i++)
            //                     {
            //                         result[i] = content[i + 1];
            //                     }
            //                     content = result;
            //                 }
            //                 string jsonContent = LitJson.JsonMapper.ToJson(content);
            //                 SaveToFile(jsonContent, pathSaveFile);
            //             }
            //             catch
            //             {
            //                 // if (app != null)
            //                 app.Quit();
            //                 saveBtn.Show();
            //                 MessageBox.Show("Failed");
            //             }
            //             saveBtn.Text = btnRawText;
            #endregion
        }

        void SaveToFile(string jsonContent, string path)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(path, false);
            }
            catch
            {
                MessageBox.Show("Save path error!");
                if (sw != null)
                    sw.Close();
                saveBtn.Show();
                return;
            }
            sw.Write(jsonContent);
            sw.Flush();
            sw.Close();
            //             var m = MessageBox.Show("已保存,是否打开查看?", "保存成功", MessageBoxButtons.YesNo);
            //             if (m == DialogResult.Yes)
            //             {
            //                 System.Diagnostics.Process.Start(pathFileSave1.Text);
            //             }
            saveBtn.Show();
        }
    }
}
