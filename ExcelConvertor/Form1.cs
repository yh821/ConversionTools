using Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Simplexcel;
using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace ExcelConvertor
{
    public partial class Form1 : Form
    {
        private class DataSource
        {
            public string extension;
            public string name;
            public int startRow;
            public int startCol;
            public DataTable source;
        }

        private IExcelConvertor xlsxConvertor;
        private IExcelConvertor xlsConvertor;

        private string importPath;
        private string exportPath;

        private ArrayList dataSources;
        private ArrayList errorSources;
        
        public Form1()
        {
            InitializeComponent();
            xlsxConvertor = new XLSXConvertor();
            xlsConvertor = new XLSConvertor();
        }

        private void btnImportBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                importPath = folderBrowserDialog.SelectedPath;
                lblImportPath.Text = importPath;
            }
        }

        private void btnExportBrowse_Click(object sender, EventArgs e)
        {
             FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

             if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
             {
                 exportPath = folderBrowserDialog.SelectedPath;
                 lblExportPath.Text = exportPath;
             }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (exportPath == null)
            {
                MessageBox.Show("请设置导出路径！！");
                return ;
            }


            xlsConvertor.CellWidth = Convert.ToInt32(txtCellWidth.Text);
            xlsxConvertor.CellWidth = Convert.ToInt32(txtCellWidth.Text);

            dataSources = new ArrayList();
            errorSources = new ArrayList();
            string[] paths = Directory.GetFiles(importPath, "*.xls");
            foreach (string path in paths)
            {
                try
                {
                    GetFileData(path);
                }
                catch (Exception ex)
                {
                    errorSources.Add(Path.GetFileName(path));
                }
            }

            int successNum = 0;
            foreach (DataSource data in dataSources)
            {
                try
                {
                    if (radioButton1.Checked)
                        SaveExcel(data, "SERVER", false);
                    else if (radioButton2.Checked)
                        SaveExcel(data, "CLIENT", false);
                    else if (radioButton3.Checked)
                        SaveExcel(data, "SERVER", true);
                    else if (radioButton4.Checked)
                        SaveExcel(data, "CLIENT", true);
                    else if (radioButton5.Checked)
                        ConvertXML(data, false);
                    else if (radioButton6.Checked)
                        ConvertXML(data, true);
                    else if (radioButton7.Checked)
                        ConvertJSON(data);

                    successNum++;
                }
                catch
                {
                    if (errorSources.IndexOf(data.name) == -1)
                        errorSources.Add(data.name);
                }
            }

            string message = "转换成功 " + successNum + " 个\n\n";
            message += "转换失败 " + errorSources.Count + " 个\n";
            if (errorSources.Count > 0)
            {
                foreach (string err in errorSources)
                {
                    message += err + "\n";
                }
            }

            MessageBox.Show(message);
            System.Diagnostics.Process.Start("explorer.exe", exportPath);
        }

        private void GetFileData(string path)
        {
            DataSource dataSource = new DataSource();
            dataSource.extension = Path.GetExtension(path);
            dataSource.name = Path.GetFileName(path);

            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
            if (dataSource.extension == ".xlsx")
                dataSource.source = xlsxConvertor.ReadExcel(stream);
            else
                dataSource.source = xlsConvertor.ReadExcel(stream);

            bool isFindControl = false;
            int rows = dataSource.source.Rows.Count;
            int columns = dataSource.source.Columns.Count;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    string value = dataSource.source.Rows[i][j].ToString();
                    if (value == "控制列")
                    {
                        dataSource.startRow = i;
                        dataSource.startCol = j;
                        isFindControl = true;
                    }
                }
            }

            if (!isFindControl && !radioButton5.Checked && !radioButton6.Checked && !radioButton7.Checked)
                throw new Exception();

            dataSources.Add(dataSource);
        }

        private void SaveExcel(DataSource data, string control, bool isTraditionalChinese)
        {
            int controlRowIndex, formatRowIndex, ignoreRowIndex;
            if (control == "SERVER")
            {
                controlRowIndex = data.startRow + 3;
                formatRowIndex = data.startRow + 2;
                ignoreRowIndex = data.startRow + 1;
            }
            else
            {
                controlRowIndex = data.startRow + 1;
                formatRowIndex = data.startRow + 2;
                ignoreRowIndex = data.startRow + 3;
            }

            IExcelConvertor convertor = data.extension == ".xlsx" ? xlsxConvertor : xlsConvertor;
            DataRow controlRow = data.source.Rows[controlRowIndex];
            DataRow formatRow = data.source.Rows[formatRowIndex];
            int rows = data.source.Rows.Count;
            int columns = data.source.Columns.Count;
            int offsetRow = data.startRow;
            int offsetCol = data.startCol + 1;
            convertor.CreateSheet("sheet1");
            for (int i = data.startRow; i < rows; i++)
            {
                if (i == ignoreRowIndex)
                {
                    offsetRow += 1;
                    continue;
                }
                offsetCol = data.startCol + 1;
                for (int j = data.startCol + 1; j < columns; j++)
                {
                    if (controlRow[j].ToString() != "")
                    {
                        string value = data.source.Rows[i][j].ToString();
                        if (isTraditionalChinese)
                            value = Microsoft.VisualBasic.Strings.StrConv(value, Microsoft.VisualBasic.VbStrConv.TraditionalChinese, 0);

                        try
                        {
                            if (i > data.startRow + 3 && value != "" && formatRow[j].ToString() == "int")
                                convertor.WriteCell(i - offsetRow, j - offsetCol, Convert.ToInt32(value));
                            else
                                convertor.WriteCell(i - offsetRow, j - offsetCol, value);
                        }
                        catch
                        {
                            convertor.WriteCell(i - offsetRow, j - offsetCol, value);
                        }
                    }
                    else
                    {
                        offsetCol++;
                    }
                }
            }

            convertor.SaveFile(exportPath + "\\" + data.name);
        }

        private void SaveXML(DataSource data, string control)
        {
            int controlRowIndex = control == "SERVER" ? data.startRow + 3 : data.startRow + 1;
            string fileName = Path.ChangeExtension(data.name, ".xml");
            StreamWriter streamWriter = new StreamWriter(exportPath + "\\" + fileName);
            XmlTextWriter xmlWriter = new XmlTextWriter(streamWriter);
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement("dataset");

            DataRow controlRow = data.source.Rows[controlRowIndex];
            int rows = data.source.Rows.Count;
            int columns = data.source.Columns.Count;

            for (int i = data.startRow + 4; i < rows; i++)
            {
                xmlWriter.WriteStartElement("data");
                for (int j = data.startCol + 1; j < columns; j++)
                {
                    if (controlRow[j].ToString() != "")
                    {
                        string value = data.source.Rows[i][j].ToString();
                        xmlWriter.WriteAttributeString(controlRow[j].ToString(), value);
                    }
                }
                xmlWriter.WriteEndElement();
            }

            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndDocument();
            xmlWriter.Close();
        }

        private void ConvertXML(DataSource data, bool isTraditionalChinese)
        {
            string fileName = Path.ChangeExtension(data.name, ".xml");
            StreamWriter streamWriter = new StreamWriter(exportPath + "\\" + fileName);
            XmlTextWriter xmlWriter = new XmlTextWriter(streamWriter);
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement("dataset");

            DataRow controlRow = data.source.Rows[data.startRow+1];
            int rows = data.source.Rows.Count;
            int columns = data.source.Columns.Count;

            for (int i = data.startRow + 3; i < rows; i++)
            {
                xmlWriter.WriteStartElement("data");
                for (int j = data.startCol; j < columns; j++)
                {
                    string value = data.source.Rows[i][j].ToString();
                    if (isTraditionalChinese)
                        value = Microsoft.VisualBasic.Strings.StrConv(value, Microsoft.VisualBasic.VbStrConv.TraditionalChinese, 0);
                    xmlWriter.WriteAttributeString(controlRow[j].ToString(), value);
                   
                }
                xmlWriter.WriteEndElement();
            }

            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndDocument();
            xmlWriter.Close();
        }

        private void ConvertJSON(DataSource data)
        {
            string fileName = Path.ChangeExtension(data.name, ".json");
            FileStream fileStream = new FileStream(exportPath + "\\" + fileName, FileMode.Create);
            StreamWriter streamWriter = new StreamWriter(fileStream);
            streamWriter.WriteLine("{\"data\":[");

            DataRow controlRow = data.source.Rows[data.startRow + 1];
            int rows = data.source.Rows.Count;
            int columns = data.source.Columns.Count;
            bool isLineStart = true;
            string stringFormat1 = "\"{0}\":\"{1}\"";
            string stringFormat2 = ", \"{0}\":\"{1}\"";

            for (int i = data.startRow + 3; i < rows; i++)
            {
                isLineStart = true;
                streamWriter.Write("\t{ ");
                for (int j = data.startCol; j < columns; j++)
                {
                    if (isLineStart)
                        streamWriter.Write(string.Format(stringFormat1, controlRow[j], data.source.Rows[i][j]));
                    else
                        streamWriter.Write(string.Format(stringFormat2, controlRow[j], data.source.Rows[i][j]));
                    isLineStart = false;
                }
                if (i == rows - 1)
                    streamWriter.Write(" }\n");
                else
                    streamWriter.Write(" },\n");
            }
            streamWriter.WriteLine("]}");

            streamWriter.Close();
            fileStream.Close();
        }
    }
}
