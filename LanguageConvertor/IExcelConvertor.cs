using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LanguageConvertor
{
    interface IExcelConvertor
    {
        int CellWidth { set; }
        void CreateSheet(string name);
        void WriteCell(int row, int col, string value);
        void WriteCell(int row, int col, int value);
        DataTable ReadExcel(Stream stream);
        void SaveFile(string path);
    }
}
