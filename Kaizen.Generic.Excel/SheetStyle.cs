using OfficeOpenXml.Style;
using System.Drawing;
using System;

namespace Kaizen.Generic.Excel
{
    public class SheetStyle
    {
        public Color BackgroundColor { get; set; }
        public Color FontColor { get; set; }
        public double FontSize { get; set; }
        public bool FontBold { get; set; }
        public ExcelBorderStyle Border { get; set; }
        public ExcelRange ExcelRange { get; set; }
        public ExcelSheet ExcelSheet { get; set; }

        public SheetStyle(ExcelRange excelRange, ExcelSheet excelSheet)
        {
            this.ExcelRange = excelRange;
            this.ExcelSheet = excelSheet;
        }

        public void ExecuteStyle()
        {
            this.AddBackgroundColor();
            this.AddFontStyle();
            this.AddBorderStyle();
        }

        public void AddAlternativeColor(Color alternativeColor)
        {
            for (int rowIndex = ExcelRange.RowStartPosition; rowIndex <= ExcelRange.RowEndPosition; rowIndex++)
            {
                var rowRange = new ExcelRange(rowIndex, ExcelRange);
                var selectedColor = rowIndex % 2 == 0 ? this.BackgroundColor : alternativeColor;
                this.AddBackgroundColor(rowRange.GetRange(this.ExcelSheet), selectedColor);
            }
        }

        private void AddBorderStyle()
        {
            var range = this.ExcelRange.GetRange(this.ExcelSheet);
            range.Style.Border.BorderAround(this.Border);
        }

        private void AddFontStyle()
        {
            var range = this.ExcelRange.GetRange(this.ExcelSheet);
            AddFontWeigth(range);
            AddFontColor(range);
        }

        private void AddFontWeigth(OfficeOpenXml.ExcelRange range)
        {
            range.Style.Font.Bold = this.FontBold;
        }

        private void AddFontColor(OfficeOpenXml.ExcelRange range)
        {
            range.Style.Font.Color.SetColor(this.FontColor);
        }

        private void AddBackgroundColor()
        {
            var range = this.ExcelRange.GetRange(this.ExcelSheet);
            this.AddBackgroundColor(range);
        }

        private void AddBackgroundColor(OfficeOpenXml.ExcelRange range)
        {
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(this.BackgroundColor);
        }

        private void AddBackgroundColor(OfficeOpenXml.ExcelRange range, Color alternativeColor)
        {
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(alternativeColor);
        }
    }
}
