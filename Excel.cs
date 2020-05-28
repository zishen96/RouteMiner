using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace RouteMiner
{
    public class Excel
    {
        public string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        //n number of rows
        public int wsRowCount = 0;
        //should be 6 columns: [streetNum] [streetName] [apt #] [city] [state] [zip]
        //optional apt # will be "" empty string if excel field is empty
        public int wsColumnCount = 6;
        public string[,] dataMatrix;

        public Excel(string path, int sheetNum)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheetNum];
            wsRowCount = ws.UsedRange.Rows.Count;
            //wsColumnCount = ws.UsedRange.Columns.Count;   //Non varying columns required for exact/specific fields to line up (6 in this case)
        }

        public void Close()
        {
            wb.Close(false);
        }

        public void ReadFile()
        {
            Debug.WriteLine("Reading file...");
            Range wsRange = (Range)ws.Range[ws.Cells[1, 1], ws.Cells[wsRowCount, wsColumnCount]];
            object[,] wsRangeContent = wsRange.Value2;
            string[,] wsStringMatrix = new string[wsRowCount, wsColumnCount];

            for (int i = 1; i <= wsRowCount; ++i)
            {
                for (int j = 1; j <= wsColumnCount; ++j)
                {
                    //Excel worksheet index starts at (1,1)
                    //wsStringMatrix index start at 0 though, so a [-1] offset needed
                    //If field is empty (null), fill it in with ""
                    //If field has value, fill with Value2 of cell converted to string
                    if (ws.Cells[i, j] != null)
                    {
                        wsStringMatrix[i - 1, j - 1] = Convert.ToString(wsRangeContent[i, j]);
                    }
                    else
                        wsStringMatrix[i - 1, j - 1] = "";
                }
            }
            dataMatrix = wsStringMatrix;
            Debug.WriteLine("Read file complete");
        }
    }
}
