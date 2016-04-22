using OfficeOpenXml;
using System.Drawing;

namespace Kaizen.Generic.Excel
{
    public interface ExcelSheet
    {
        ExcelWorksheet ExcelWorksheet { get; set; }
        ExcelSheet addAlternativeBackgroundColor(Color color, Color alternativeColor);
        ExcelSheet addBackgroundColor(Color color);
        ExcelSheet addHeaderBackgroundColor(Color color);
    }
}