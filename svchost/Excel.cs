using Microsoft.Office.Interop.Excel;
using System.IO;
using _Excel = Microsoft.Office.Interop.Excel;

namespace svchost
{
    internal class Excel
    {
        private string path = "";
        private _Application excel = new _Excel.Application();
        private Workbook wb;
        private Worksheet ws;

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(Path.Combine(System.Environment.CurrentDirectory, "source.xlsx"));
            ws = wb.Worksheets[Sheet];

        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
            {
                return System.Convert.ToString(ws.Cells[i, j].Value2);
            }
            else
            {
                return "";
            }
        }

        public void Close()
        {
            wb.Close();
        }


    }
}
