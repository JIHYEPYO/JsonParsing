using System;
using System.Collections.Generic;
using System.Data;
using Newtonsoft.Json;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;

namespace JsonParsing
{
    public partial class Form1 : Form
    {
        static string fileName;
        static string fileFullName;
        static string filePath;
        static List<string> name;
        DataTable dataTable;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ShowFileOpenDialog();
            ConvertJsonToDataTable();
            comboBox1.Visible = false;
        }

        private static void ShowFileOpenDialog()
        {
            //파일오픈창 생성 및 설정
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Json 파일 선택";
            ofd.Filter = "json 파일 (*.json) | *.json; | 모든 파일 (*.*) | *.*";

            //파일 오픈창 로드
            DialogResult dr = ofd.ShowDialog();

            //OK버튼 클릭시
            if (dr == DialogResult.OK)
            {
                fileName = ofd.SafeFileName;
                fileFullName = ofd.FileName;
                filePath = fileFullName.Replace(fileName, "");
            }
        }

        private void ConvertJsonToDataTable()
        {
            name = new List<string>();

            try
            {
                string jsonString = File.ReadAllText(fileFullName);

                if (!String.IsNullOrWhiteSpace(jsonString))
                {
                    dynamic dynObj = JsonConvert.DeserializeObject(jsonString);

                    dataTable = new DataTable();
                    dataTable.Columns.Add("num", typeof(string));
                    dataTable.Columns.Add("userName", typeof(string));
                    dataTable.Columns.Add("time", typeof(string));
                    dataTable.Columns.Add("tag", typeof(string));
                    dataTable.Columns.Add("filename", typeof(string));
                    dataTable.Columns.Add("location", typeof(string));
                    dataTable.Columns.Add("mode", typeof(string));
                    dataTable.Columns.Add("text", typeof(string));
                    dataTable.Columns.Add("wpm", typeof(string));
                    dataTable.Columns.Add("delete", typeof(string));
                    dataTable.Columns.Add("gap", typeof(string));
                    dataTable.Columns.Add("activityName", typeof(string));
                    dataTable.Columns.Add("pushButton", typeof(string));
                    dataTable.Columns.Add("value", typeof(string));

                    AddRows(dynObj.ButtonMenu, "ButtonMenu");
                    AddRows(dynObj.ButtonTraining, "ButtonTraining");
                    AddRows(dynObj.ButtonViewer, "ButtonViewer");
                    AddRows(dynObj.File, "File");
                    AddRows(dynObj.Mode, "Mode");
                    AddRows(dynObj.Training, "Training");
                    AddRows(dynObj.Viewer, "Viewer");

                    if (dataTable.Rows.Count > 0)
                    {
                        dataGridView.DataSource = dataTable;
                        dataGridView.Columns["num"].Visible = false;
                        dataGridView.Columns["userName"].HeaderText = "User Name";
                        dataGridView.Columns["time"].HeaderText = "Date";
                        dataGridView.Columns["tag"].HeaderText = "TAG";
                        dataGridView.Columns["filename"].HeaderText = "File Name";
                        dataGridView.Columns["location"].HeaderText = "Location";
                        dataGridView.Columns["mode"].HeaderText = "Mode";
                        dataGridView.Columns["text"].HeaderText = "Text";
                        dataGridView.Columns["wpm"].HeaderText = "WPM";
                        dataGridView.Columns["delete"].HeaderText = "Delete";
                        dataGridView.Columns["gap"].HeaderText = "Gap";
                        dataGridView.Columns["activityName"].HeaderText = "Activity Name";
                        dataGridView.Columns["pushButton"].HeaderText = "Push Button";
                        dataGridView.Columns["value"].HeaderText = "Value";

                        dataTable.DefaultView.Sort = "time";
                        dataGridView.ClearSelection();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SortTable()
        {
            try
            {
                string jsonString = File.ReadAllText(fileFullName);

                if (!String.IsNullOrWhiteSpace(jsonString))
                {
                    dynamic dynObj = JsonConvert.DeserializeObject(jsonString);

                    dataTable = new DataTable();
                    dataTable.Columns.Add("num", typeof(string));
                    dataTable.Columns.Add("userName", typeof(string));
                    dataTable.Columns.Add("time", typeof(string));
                    dataTable.Columns.Add("tag", typeof(string));
                    dataTable.Columns.Add("filename", typeof(string));
                    dataTable.Columns.Add("location", typeof(string));
                    dataTable.Columns.Add("mode", typeof(string));
                    dataTable.Columns.Add("text", typeof(string));
                    dataTable.Columns.Add("wpm", typeof(string));
                    dataTable.Columns.Add("delete", typeof(string));
                    dataTable.Columns.Add("gap", typeof(string));
                    dataTable.Columns.Add("activityName", typeof(string));
                    dataTable.Columns.Add("pushButton", typeof(string));
                    dataTable.Columns.Add("value", typeof(string));

                    SortAddRows(dynObj.ButtonMenu, "ButtonMenu");
                    SortAddRows(dynObj.ButtonTraining, "ButtonTraining");
                    SortAddRows(dynObj.ButtonViewer, "ButtonViewer");
                    SortAddRows(dynObj.File, "File");
                    SortAddRows(dynObj.Mode, "Mode");
                    SortAddRows(dynObj.Training, "Training");
                    SortAddRows(dynObj.Viewer, "Viewer");

                    if (dataTable.Rows.Count > 0)
                    {
                        dataGridView.DataSource = dataTable;
                        dataGridView.Columns["num"].Visible = false;
                        dataGridView.Columns["userName"].HeaderText = "User Name";
                        dataGridView.Columns["time"].HeaderText = "Date";
                        dataGridView.Columns["tag"].HeaderText = "TAG";
                        dataGridView.Columns["filename"].HeaderText = "File Name";
                        dataGridView.Columns["location"].HeaderText = "Location";
                        dataGridView.Columns["mode"].HeaderText = "Mode";
                        dataGridView.Columns["text"].HeaderText = "Text";
                        dataGridView.Columns["wpm"].HeaderText = "WPM";
                        dataGridView.Columns["delete"].HeaderText = "Delete";
                        dataGridView.Columns["gap"].HeaderText = "Gap";
                        dataGridView.Columns["activityName"].HeaderText = "Activity Name";
                        dataGridView.Columns["pushButton"].HeaderText = "Push Button";
                        dataGridView.Columns["value"].HeaderText = "Value";

                        dataTable.DefaultView.Sort = "time";
                        dataGridView.ClearSelection();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void AddRows(dynamic rows, string st)
        {
            bool flag = false;
            foreach (var row in rows)
            {
                string bt1 = Convert.ToString(row);
                string[] RowData = Regex.Split(bt1.Replace("{", "").Replace("}", ""), "\",");
                DataRow nr = dataTable.NewRow();

                foreach (string rowData in RowData)
                {
                    flag = false;
                    try
                    {
                        int idx = rowData.IndexOf(":");
                        int idx2 = 0;

                        if (rowData.StartsWith("\"-K"))
                        {
                            idx2 = idx;
                            idx += rowData.Substring(idx + 1).IndexOf(":");
                        }
                        string RowColumns = rowData.Substring(0, idx - 1).Replace("\"", "").Trim();
                        if (idx2 != 0)
                        {
                            RowColumns = rowData.Substring(idx2 + 7, rowData.Substring(idx2 + 7).IndexOf(":") - 1);
                        }

                        string RowDataString = rowData.Substring(idx + 2).Replace("\"", "");
                        if (RowDataString.Contains("\r") || RowDataString.Contains("\n"))
                        {
                            RowDataString = RowDataString.Replace("\r", "");
                            RowDataString = RowDataString.Replace("\n", "");
                        }
                        nr[RowColumns] = RowDataString;
                        nr["tag"] = st;

                        if (name == null || (!name.Contains(RowDataString)) && RowColumns.Equals("userName"))
                        {
                            name.Add(RowDataString);
                        }
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }
                }
                if (!flag)
                {
                    dataTable.Rows.Add(nr);
                }
            }
        }

        private void SortAddRows(dynamic rows, string st)
        {
            bool flag = false;
            foreach (var row in rows)
            {
                string bt1 = Convert.ToString(row);
                string[] RowData = Regex.Split(bt1.Replace("{", "").Replace("}", ""), "\",");
                DataRow nr = dataTable.NewRow();

                foreach (string rowData in RowData)
                {
                    try
                    {
                        int idx = rowData.IndexOf(":");
                        int idx2 = 0;

                        if (rowData.StartsWith("\"-K"))
                        {
                            idx2 = idx;
                            idx += rowData.Substring(idx + 1).IndexOf(":");
                        }

                        string RowColumns = rowData.Substring(0, idx - 1).Replace("\"", "").Trim();
                        if (idx2 != 0)
                        {
                            RowColumns = rowData.Substring(idx2 + 7, rowData.Substring(idx2 + 7).IndexOf(":") - 1);
                        }

                        string RowDataString = rowData.Substring(idx + 2).Replace("\"", "");

                        if (RowDataString.Contains(comboBox1.SelectedItem.ToString()))
                        {
                            flag = true;
                        }

                        nr[RowColumns] = RowDataString;
                        nr["tag"] = st;
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }
                }
                if (flag)
                {
                    dataTable.Rows.Add(nr);
                    flag = false;
                }
            }
        }

        private void ComboBoxSet()
        {
            foreach (var cb in name)
            {
                comboBox1.Items.Add(cb);
            }
        }

        private void originToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            ConvertJsonToDataTable();
            comboBox1.Visible = false;
        }

        private void sortingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            ComboBoxSet();
            comboBox1.Visible = true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SortTable();
        }

        private void tAGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 form = new JsonParsing.Form2();
            form.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            XLWorkbook wb = new XLWorkbook();

            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "Please select a storage path";
            dialog.OverwritePrompt = true;
            dialog.Filter = "Excel 파일 (*.xlsx) | *.xlsx";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                wb.Worksheets.Add(dataTable,"Log Data");
                wb.SaveAs(dialog.FileName);
                MessageBox.Show("SAVE!!");
            }
            
        }
    }
}
