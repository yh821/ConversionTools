using Excel;
using Simplexcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LanguageConvertor
{
    class XLSXConvertor : IExcelConvertor
    {
        private Worksheet sheet;
        private int width;
        public XLSXConvertor()
        {
        }

        public int CellWidth
        {
            set
            {
                width = value;
            }
        }

        public void CreateSheet(string name)
        {
            sheet = new Worksheet(name);
        }

        public void WriteCell(int row, int col, string value)
        {
            sheet[row, col] = value;
            sheet.ColumnWidths[col] = width;
        }

        public void WriteCell(int row, int col, int value)
        {
            sheet[row, col] = value;
            sheet.ColumnWidths[col] = width;
        }

        public DataTable ReadExcel(Stream stream)
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            return result.Tables[0];
        }

        public void SaveFile(string path)
        {
            Workbook workbook = new Workbook();
            workbook.Add(sheet);
            workbook.Save(path);
        }
    }
}
