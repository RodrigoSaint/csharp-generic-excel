using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace Kaizen.Generic.Excel
{
    public class ExcelFile
    {
        public ExcelPackage File { get; protected set; }
        protected Dictionary<string, ExcelSheet> ExcelSheetCollection { get; set; }

        public ExcelFile()
        {
            this.File = new ExcelPackage();
            this.ExcelSheetCollection = new Dictionary<string, ExcelSheet>();
        }

        public void AddSheet<T>(string name, ICollection<T> rows)
        {
            var sheet = this.File.Workbook.Worksheets.Add(name);
            this.ExcelSheetCollection.Add(name, new ExcelSheetImplementation<T>(sheet, rows));
        }

        public ExcelSheet GetSheet(string name)
        {
            return this.ExcelSheetCollection[name];
        }

        public byte[] GetFileStream()
        {
            return this.File.GetAsByteArray();
        }

    }
}