using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
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
            ps.excel = pathExcel1.Text;
            ps.saveFile = pathFileSave1.Text;
            ps.removedHead = (int)ignoreLines.Value;
            ps.sheets = (int)sheet1.Value;
            ps.isCompress = isUseCompress.Checked;
            ps.unCompressSrcPath = unCompressSrcPath.Text;
            ps.unCompressSavePath = unCompressSavePath.Text;

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
                pathExcel1.Text = ps.excel;
                pathFileSave1.Text = ps.saveFile;
                ignoreLines.Value = ps.removedHead < ignoreLines.Maximum && ps.removedHead >= ignoreLines.Minimum ? ps.removedHead : ignoreLines.Minimum;
                sheet1.Value = ps.sheets >= sheet1.Minimum && ps.sheets <= sheet1.Maximum ? ps.sheets : sheet1.Minimum;
                isUseCompress.Checked = ps.isCompress;
                unCompressSrcPath.Text = ps.unCompressSrcPath;
                unCompressSavePath.Text = ps.unCompressSavePath;
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


        private void sheet1_ValueChanged(object sender, EventArgs e)
        {
            RememberPaths();
        }


        private void unCompressSrcPath_TextChanged(object sender, EventArgs e)
        {
            RememberPaths();
        }

        private void unCompressSavePath_TextChanged(object sender, EventArgs e)
        {
            RememberPaths();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            ReadAndSave(pathExcel1.Text, (int)sheet1.Value, pathFileSave1.Text, (int)ignoreLines.Value);
            MessageBox.Show("完成");
        }

        private void isUseCompress_CheckedChanged(object sender, EventArgs e)
        {
            RememberPaths();
        }

        private void unCompressBtn_Click(object sender, EventArgs e)
        {
            ClickBtn2SetPath(unCompressSrcPath, false);
        }

        private void unCompressSaveBtn_Click(object sender, EventArgs e)
        {
            ClickBtn2SetPath(unCompressSavePath, false);
        }

        private void SaveUncompressBtn_Click(object sender, EventArgs e)
        {
            ReadAndSave_UnCompress(unCompressSrcPath.Text, unCompressSavePath.Text);
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

        void ReadAndSave(string pathExcel, int sheetPage, string pathSaveFile, int removeHead)
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

            if (!File.Exists(pathExcel))
                return;

            saveBtn.Hide();
            using (FileStream fs = File.Open(pathExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fs))
                {
                    DataSet dataset = reader.AsDataSet();
                    if (dataset.Tables.Count == 0)
                        return;

                    int _sheetPage = sheetPage - 1;
                   
                    //读取
                    {
                        string allSheetsListJson = string.Empty;
                        for (int _sheetIndex = 0; _sheetIndex < dataset.Tables.Count; _sheetIndex++)
                        {
                            allSheetsListJson += GetSheetJsonContent(dataset, _sheetIndex, removeHead) + (_sheetIndex == dataset.Tables.Count-1?"":",");                           
                        }
                        allSheetsListJson = AddBigMark(allSheetsListJson); 
                        if (isUseCompress.Checked)
                        {
                            try
                            {
                                byte[] jsonB = Encoding.UTF8.GetBytes(allSheetsListJson);
                                byte[] _result = ZipHelper.GZipCompress(jsonB);
                                SaveToFile(_result, pathSaveFile);
                                return;
                            }
                            catch
                            {
                                MessageBox.Show("压缩失败");
                            }
                        }
                        SaveToFile(allSheetsListJson, pathSaveFile);
                    }
                }
                saveBtn.Show();
            }
        }

        char dyMark = '"';
        string AddDQMark(string s)
        {
            return dyMark + s + dyMark;
        }
        string AddBigMark(string s)
        {
            return "{" + s + "}";
        }

        string GetSheetJsonContent(DataSet _dataset, int _sheetPage, int _removeHead)
        {
            var table = _dataset.Tables[_sheetPage];
            if (table.Rows.Count <= _removeHead)
                return string.Empty;
            if (table.Columns.Count < 1)
                return string.Empty;
            string tableName = table.TableName;

            string dicContentJson = string.Empty;

            List<string> tempKey = new List<string>();
            for (int _c = _removeHead + 1; _c < table.Rows.Count; _c++)
            {
                var v = table.Rows[_c][0];
                if (v != DBNull.Value)
                {
                    tempKey.Add(v.ToString());
                }
            }

            List<string> valuesStr = new List<string>();

            List<SupportTypeRecord> valuesType = new List<SupportTypeRecord>();            
            List<string> membersNames = new List<string>();
            for (int columnId = 1; columnId < table.Columns.Count; columnId++)
            {
                string s = (table.Rows[_removeHead][columnId]).ToString();
                string[] split = s.Split(',');
                if (split == null || split.Length != 2)
                {
                    MessageBox.Show("split error " + columnId);
                    return string.Empty;
                }
                membersNames.Add(split[1]);

                if (split[0] == "string")
                {
                    valuesType.Add(new SupportTypeRecord(false,true));
                }
                else if (split[0] == "int")
                {
                    valuesType.Add(new SupportTypeRecord(false, false));
                }
                else if (split[0] == "float")
                {
                    valuesType.Add(new SupportTypeRecord(false, false));
                }
                else if (split[0] == "list<string>")
                {
                    valuesType.Add(new SupportTypeRecord(true, true));
                }
                else if (split[0] == "list<int>")
                {
                    valuesType.Add(new SupportTypeRecord(true, false));
                }
                else if (split[0] == "list<float>")
                {
                    valuesType.Add(new SupportTypeRecord(true, false));
                }
            }

            for (int i = _removeHead + 1; i < table.Rows.Count; i++)
            {
                string str = string.Empty;
                for (int valuesId = 1; valuesId < table.Columns.Count; valuesId++)
                {
                    var rv = table.Rows[i][valuesId];
                    string realContent = rv.ToString();
                    SupportTypeRecord tr = valuesType[valuesId - 1];

                    if (rv == DBNull.Value)
                    {
                        if (!tr.isList && !tr.isString)
                            realContent = "0";
                    }                    

                    if (tr.isList)
                    {
                        if (tr.isString)
                        {
                            string[] strs = realContent.Split(',');
                            for(int sId = 0; sId < strs.Length; sId++)
                            {
                                strs[sId] = AddDQMark(strs[sId]);
                            }
                            realContent = string.Empty;
                            for (int sId = 0; sId < strs.Length; sId++)
                            {
                                realContent += strs[sId] + (sId == strs.Length - 1 ? "" : ",");
                            }
                        }
                        realContent = "[" + realContent + "]";
                    }
                    else if (tr.isString)
                    {
                        realContent = AddDQMark(realContent);
                    }
                    str += AddDQMark(membersNames[valuesId - 1])+":" + realContent + (valuesId == table.Columns.Count-1? "":",");                                     
                }
                valuesStr.Add(str);
            }

            for (int i = 0; i < tempKey.Count; i++)
            {
                dicContentJson += (AddDQMark(tempKey[i]) + ":" + AddBigMark(valuesStr[i]) + (i == tempKey.Count - 1 ? "" : ","));
            }

            string result = AddDQMark(tableName) + ":" + AddBigMark(dicContentJson); //放最后
           
            return result;
        }

        void ReadAndSave_UnCompress(string pathCompressFile, string pathToSave)
        {
            try
            {
                FileStream srcFs = new FileStream(pathCompressFile, FileMode.OpenOrCreate);
                byte[] _compressedBytes = new byte[srcFs.Length];
                srcFs.Read(_compressedBytes, 0, _compressedBytes.Length);
                srcFs.Close();
                string unCompressErr;
                byte[] _readResult = ZipHelper.GZipDecompress(_compressedBytes, out unCompressErr);
                SaveToFile(_readResult, pathToSave);
            }
            catch
            {
                MessageBox.Show("UnCompress err");
            }
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
            saveBtn.Show();
            //             var m = MessageBox.Show("已保存,是否打开查看?", "保存成功", MessageBoxButtons.YesNo);
            //             if (m == DialogResult.Yes)
            //             {
            //                 System.Diagnostics.Process.Start(pathFileSave1.Text);
            //             }
        }
        void SaveToFile(byte[] jsonContent, string path)
        {
            FileStream fs = null;
            try
            {
                fs = new FileStream(path, FileMode.OpenOrCreate);
            }
            catch
            {
                MessageBox.Show("Save path error!");
                if (fs != null)
                    fs.Close();
                saveBtn.Show();
                return;
            }
            fs.SetLength(jsonContent.Length);
            fs.Write(jsonContent, 0, jsonContent.Length);
            fs.Close();
            saveBtn.Show();
            //             var m = MessageBox.Show("已保存,是否打开查看?", "保存成功", MessageBoxButtons.YesNo);
            //             if (m == DialogResult.Yes)
            //             {
            //                 System.Diagnostics.Process.Start(pathFileSave1.Text);
            //             }

        }

        private void ignoreLines_ValueChanged(object sender, EventArgs e)
        {
            RememberPaths();
        }
    }
}

public class SupportTypeRecord
{
    public bool isList;
    public bool isString;

    public SupportTypeRecord(bool _isList,bool _isStr)
    {
        isList = _isList;
        isString = _isStr;
    }
}

