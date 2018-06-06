using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConvertor
{
    class XLSConvertor : IExcelConvertor
    {
        private HSSFWorkbook workbook;
        private ISheet sheet;
        private int width;

        public XLSConvertor()
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
            workbook = new HSSFWorkbook();
            sheet = workbook.CreateSheet("sheet1");
            sheet.DefaultColumnWidth = width;
        }

        public void WriteCell(int row, int col, string value)
        {
            IRow rowSheet = sheet.GetRow(row);
            if (rowSheet == null)
                rowSheet = sheet.CreateRow(row);

            ICell cell = rowSheet.CreateCell(col);
            cell.SetCellValue(value);
        }

        public void WriteCell(int row, int col, int value)
        {
            IRow rowSheet = sheet.GetRow(row);
            if (rowSheet == null)
                rowSheet = sheet.CreateRow(row);

            ICell cell = rowSheet.CreateCell(col);
            cell.SetCellValue(value);
        }

        public DataTable ReadExcel(Stream stream)
        {
            HSSFWorkbook workbook = new HSSFWorkbook(stream);
            ISheet sheet = workbook.GetSheetAt(0);
            IEnumerator rowEnumerator = sheet.GetRowEnumerator();

            int maxCellNum = 0;
            int rowNum = sheet.LastRowNum;
            for (int i = 0; i < rowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    int cellNum = sheet.GetRow(i).LastCellNum;
                    if (cellNum > maxCellNum)
                        maxCellNum = cellNum;
                }
            }

            DataTable source = new DataTable();
            for (int i =0; i<maxCellNum; i++)
            {
                DataColumn column = new DataColumn();
                source.Columns.Add(column);
            }

            while (rowEnumerator.MoveNext())
            {
                IRow row = (IRow)rowEnumerator.Current;
                int cellCount = row.LastCellNum;
                DataRow dataRow = source.NewRow();
                for (int i = 0; i < cellCount; i++)
                {
                    ICell cell = row.GetCell(i);
                    if (cell != null)
                        dataRow[i] = cell.ToString();
                    else
                        dataRow[i] = null;
                }
                source.Rows.Add(dataRow);
            }
            return source;
        }

        public void SaveFile(string path)
        {
            FileStream file = File.OpenWrite(path);
            workbook.Write(file);
            file.Flush();
            file.Close();
        }
    }
}
