using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LanguageConvertor
{
    public partial class Form1 : Form
    {
        private const string SRC_PATH = "src";
        private const string DES_PATH = "des";
        private const string JSON_PATH = "json";
        private const string EXCEL_PATH = "excel";
        private const string EXCEL_NAME = "language.xlsx";

        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            IProgress<int> progress = new Progress<int>((progressValue) => { progressBar1.Value = progressValue; });
                
            progressBar1.Value = 0;
            if (radioButton1.Checked) //抽取中文
            {
                await Task.Run(() => ExtractChinese(progress));
            }
            else if (radioButton2.Checked)//替换中文
            {
                button1.Enabled = false;
                await Task.Run(() => ReplaceChinese(progress));
                button1.Enabled = true;
            }
            else if (radioButton3.Checked)//导出json
            {
                button1.Enabled = false;
                await Task.Run(() => ExportJson(progress));
                button1.Enabled = true;
            }
        }

        private void ExtractChinese(IProgress<int> progress)
        {
            var currentDir = Directory.GetCurrentDirectory();
            var srcDir = Path.Combine(currentDir, SRC_PATH);
            if (!Directory.Exists(srcDir))
                Directory.CreateDirectory(srcDir);

            var data = new ArrayList();
            try
            {
                var files = Directory.GetFiles(srcDir, "*.xml");
                for (int i = 0; i < files.Length; i++)
                {
                    GetXmlFileData(files[i], data);
                    progress.Report((i + 1) * 100 / files.Length - 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("提取数据出错:" + ex.Message);
            }

            try
            {
                var excelDir = Path.Combine(currentDir, EXCEL_PATH);
                if (!Directory.Exists(excelDir))
                    Directory.CreateDirectory(excelDir);

                var saveDir = Path.Combine(excelDir, EXCEL_NAME);
                SaveExcelFileData(saveDir, data);

                progress.Report(100);
                MessageBox.Show("抽取成功!");
                System.Diagnostics.Process.Start("explorer.exe", excelDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存数据出错:" + ex.Message);
            }
        }

        /// <summary>
        /// 获取xml里面的中文数据
        /// </summary>
        /// <param name="path"></param>
        /// <param name="data"></param>
        private void GetXmlFileData(string path, ArrayList data)
        {
            var text = File.ReadAllText(path);
            Regex reg1 = new Regex("(?<=\").*?(?=\")");     //提取双引号里面的内容
            Regex reg2 = new Regex("[\u4E00-\u9FFF]+");     //提取中文内容
            MatchCollection mc1 = reg1.Matches(text);
            foreach (Match m1 in mc1)
            {
                MatchCollection mc2 = reg2.Matches(m1.Value);
                if (mc2.Count == 0) continue;
                string format = m1.Value;
                if (data.IndexOf(format) == -1)
                    data.Add(format);
            }
        }

        /// <summary>
        /// 把xml里面的中文数据保存到excel
        /// </summary>
        /// <param name="path"></param>
        /// <param name="data"></param>
        private void SaveExcelFileData(string path, ArrayList data)
        {
            var xlsxConvertor = new XLSXConvertor();
            xlsxConvertor.CellWidth = 50;
            xlsxConvertor.CreateSheet("sheet1");
            xlsxConvertor.WriteCell(0, 0, "key");
            xlsxConvertor.WriteCell(0, 1, "value");
            for (int i = 0; i < data.Count; i++)
            {
                xlsxConvertor.WriteCell(i + 1, 0, data[i].ToString());
            }
            xlsxConvertor.SaveFile(path);
        }

        private void ReplaceChinese(IProgress<int> progress)
        {
            var currentDir = Directory.GetCurrentDirectory();
            var languageDir = Path.Combine(currentDir, EXCEL_PATH, EXCEL_NAME);
            if (!File.Exists(languageDir))
            {
                MessageBox.Show("文件不存在:" + EXCEL_NAME);
                return;
            }

            var map = new Dictionary<string, string>();
            try
            {
                GetExcelFileData(languageDir, map);
                progress.Report(5);
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取" + EXCEL_NAME + "数据失败:" + ex.Message);
            }

            try
            {
                var srcDir = Path.Combine(currentDir, SRC_PATH);
                if (!Directory.Exists(srcDir))
                    Directory.CreateDirectory(srcDir);

                var desDir = Path.Combine(currentDir, DES_PATH);
                if (!Directory.Exists(desDir))
                    Directory.CreateDirectory(desDir);

                var files = Directory.GetFiles(srcDir, "*.xml");
                for (int i=0; i<files.Length; i++)
                {
                    ReplaceXmlFileData(files[i], desDir, map);
                    progress.Report(5 + ((i+1)*100/files.Length - 5));
                }

                progress.Report(100);
                MessageBox.Show("替换成功!");
                System.Diagnostics.Process.Start("explorer.exe", desDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show("替换xml数据失败:" + ex.Message);
            }
        }

        /// <summary>
        /// 获取excel里面的语言包数据
        /// </summary>
        /// <param name="path"></param>
        /// <param name="map"></param>
        private void GetExcelFileData(string path, Dictionary<string, string> map)
        {
            var xlsxConvertor = new XLSXConvertor();
            var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            var source = xlsxConvertor.ReadExcel(stream);
            int rows = source.Rows.Count;
            int columns = source.Columns.Count;
            int keyRow = 0;
            int keyCol = 0;
            int valueCol = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    string value = source.Rows[i][j].ToString();
                    if (value == "key")
                    {
                        keyRow = i;
                        keyCol = j;
                    }
                    else if (value == "value")
                    {
                        valueCol = j - keyCol;
                        break ;
                    }
                }
            }

            for (int i = keyRow + 1; i < rows; i++)
            {
                string key = source.Rows[i][keyCol].ToString();
                string value = source.Rows[i][valueCol].ToString();
                map.Add(key, value);
            }
        }

        /// <summary>
        /// 替换xml里面的中文数据
        /// </summary>
        /// <param name="fileDir"></param>
        /// <param name="saveDir"></param>
        /// <param name="map"></param>
        private void ReplaceXmlFileData(string fileDir, string saveDir, Dictionary<string, string> map)
        {
            var text = File.ReadAllText(fileDir);
            var reg1 = new Regex("(?<=\").*?(?=\")");     //提取双引号里面的内容
            var reg2 = new Regex("[\u4E00-\u9FFF]+");     //提取中文内容
            var mc1 = reg1.Matches(text);
            var needReplace = new List<string>();
            foreach (Match m1 in mc1)
            {
                MatchCollection mc2 = reg2.Matches(m1.Value);
                if (mc2.Count == 0) continue;
                if (map.ContainsKey(m1.Value))
                    needReplace.Add(m1.Value);
            }
            
            var fileName = Path.GetFileName(fileDir);
            var savePath = Path.Combine(saveDir, fileName);
            if (needReplace.Count > 0)
            {
                needReplace.Sort(new StringComparer());
                for (int i = 0; i < needReplace.Count; i++)
                {
                    text = text.Replace(needReplace[i], map[needReplace[i]]);
                }
                File.WriteAllText(savePath, text, Encoding.UTF8);
            }
            else
            {
                File.Copy(fileDir, savePath, true);
            }
        }

        /// <summary>
        /// 导出json文件
        /// </summary>
        /// <param name="progress"></param>
        private void ExportJson(IProgress<int> progress)
        {
            var currentDir = Directory.GetCurrentDirectory();
            var srcDir = Path.Combine(currentDir, SRC_PATH);
            if (!Directory.Exists(srcDir))
                Directory.CreateDirectory(srcDir);

            var jsonDir = Path.Combine(currentDir, JSON_PATH);
            if (!Directory.Exists(jsonDir))
                Directory.CreateDirectory(jsonDir);

            try
            {
                var files = Directory.GetFiles(srcDir, "*.xml");
                for (int i = 0; i < files.Length; i++)
                {
                    ExportJsonFromXml(files[i], jsonDir);
                    progress.Report((i + 1) * 100 / files.Length - 1);
                }
                MessageBox.Show("导出成功!");
                System.Diagnostics.Process.Start("explorer.exe", jsonDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出json出错:" + ex.Message);
            }
        }

        private void ExportJsonFromXml(string path, string output)
        {
            string text = File.ReadAllText(path);
            text = Regex.Replace(text, @"<\?.*?\?>", "");
            text = Regex.Replace(text, @"<!--.*?-->", "");
            text = text.Replace("\" ", "\", \"");
            text = text.Replace("<data ", "{ \"");
            text = text.Replace(", \"/>", " },");
            text = text.Replace("/>", "},");
            text = text.Replace("=", "\":");
            text = text.Replace("<dataset>","{\"data\":[");
            text = text.Replace("</dataset>", "");
            int len = text.LastIndexOf(',');
            text = text.Remove(len);
            text += "\n]}";

            var savePath = Path.Combine(output, Path.GetFileNameWithoutExtension(path) + ".json");
            File.WriteAllText(savePath, text, Encoding.UTF8);
        }
    }

    public class StringComparer : IComparer<string>
    {
        public int Compare(string valueA, string valueB)
        {
            return valueB.Length.CompareTo(valueA.Length);
        }
    }
}
