using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Xml.Serialization;
using System.IO;
using OfficeOpenXml;
using System.Windows.Forms;

namespace RouteMiner
{
    public abstract class Creator
    {
        public abstract Product Create(object data, Excel excelSource);
    }

    public class ExporterOne : Creator
    {
        private ReportOne reportOne;
        public override Product Create(object data, Excel excelSource)
        {
            reportOne = new ReportOne();
            List<string[]> response = new List<string[]>();
            response = (List<string[]>)data;
            var ws = reportOne.excelPackage.Workbook.Worksheets["Report1"];

            for (int x = 2; x <= excelSource.wsRowCount + 1; ++x)
            {
                ws.Cells[x, 1].Value = $"{response[x - 2][0]}";
                ws.Cells[x, 2].Value = $"{response[x - 2][1]}";
            }
            return reportOne;
        }
    }

    public class ExporterTwo : Creator
    {
        private ReportTwo reportTwo;
        public override Product Create(object data, Excel excelSource)
        {
            reportTwo = new ReportTwo();
            Dictionary<string, int> tally = new Dictionary<string, int>();
            tally = (Dictionary<string, int>)data;
            var ws = reportTwo.excelPackage.Workbook.Worksheets["Report2"];

            for (int x = 0; x < tally.Count; ++x)
            {
                ws.Cells[x + 2, 1].Value = tally.ElementAt(x).Key;
                ws.Cells[x + 2, 2].Value = tally.ElementAt(x).Value;
            }
            return reportTwo;
        }
    }

    public abstract class Product
    {
        public ExcelPackage excelPackage;
        public Product()
        {
            excelPackage = new ExcelPackage();
        }

        public void Save()
        {
            SaveFileDialog destination = new SaveFileDialog();
            destination.DefaultExt = "*.xlsx";
            destination.Filter = "Excel File (*.xlsx)|*.xlsx";
            if (destination.ShowDialog() == DialogResult.OK)
            {
                FileInfo excelFile = new FileInfo(destination.FileName);
                excelPackage.SaveAs(excelFile);
            }
        }
    }

    public class ReportOne : Product
    {
        public ReportOne()
        {
            excelPackage.Workbook.Worksheets.Add("Report1");

            var ws = excelPackage.Workbook.Worksheets["Report1"];

            ws.Cells[1, 1].Value = "Address";
            ws.Cells[1, 2].Value = "Carrier Route";
        }
    }

    public class ReportTwo : Product
    {
        public ReportTwo()
        {
            excelPackage.Workbook.Worksheets.Add("Report2");

            var ws = excelPackage.Workbook.Worksheets["Report2"];

            ws.Cells[1, 1].Value = "Carrier Route";
            ws.Cells[1, 2].Value = "Num. of Address in Route";
        }
    }
}
