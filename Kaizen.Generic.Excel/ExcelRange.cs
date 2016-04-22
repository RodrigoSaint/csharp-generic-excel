using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kaizen.Generic.Excel
{
    public class ExcelRange
    {
        protected const int MARGIN_SIZE_ROW_TABLE = 2;
        protected const int MARGIN_SIZE_ROW_HEADER = 1;
        protected const int MARGIN_SIZE_COLUMN = 1;
        protected const int MARGIN_SIZE_BOTTOM_COLUMN = 1;

        public int RowStartPosition { get; protected set; }
        public int RowEndPosition { get; protected set; }
        public int ColumnStartPosition { get; protected set; }
        public int ColumnEndPosition { get; protected set; }

        public ExcelRange(int columnEndPosition)
        {
            this.RowStartPosition = ExcelRange.MARGIN_SIZE_ROW_HEADER;
            this.RowEndPosition = ExcelRange.MARGIN_SIZE_ROW_HEADER;
            this.ColumnStartPosition = ExcelRange.MARGIN_SIZE_COLUMN;
            this.ColumnEndPosition = columnEndPosition;
        }

        public ExcelRange(int rowEndPosition, int columnEndPosition)
        {
            this.RowStartPosition = ExcelRange.MARGIN_SIZE_ROW_TABLE;
            this.ColumnStartPosition = ExcelRange.MARGIN_SIZE_COLUMN;
            this.RowEndPosition = rowEndPosition + ExcelRange.MARGIN_SIZE_BOTTOM_COLUMN;
            this.ColumnEndPosition = columnEndPosition;
        }

        public ExcelRange(int rowStartPosition, int rowEndPosition, int columnStartPosition, int columnEndPosition) 
            : this(rowEndPosition, columnEndPosition)
        {
            this.RowStartPosition = rowStartPosition + ExcelRange.MARGIN_SIZE_ROW_TABLE;
            this.ColumnStartPosition = columnStartPosition + ExcelRange.MARGIN_SIZE_COLUMN;
        }

        public ExcelRange(int rowPosition, ExcelRange range)
        {
            this.RowStartPosition = rowPosition;
            this.ColumnStartPosition = range.ColumnStartPosition;
            this.RowEndPosition = rowPosition;
            this.ColumnEndPosition = range.ColumnEndPosition;
        }

        public OfficeOpenXml.ExcelRange GetRange(ExcelSheet excelSheet)
        {
            return excelSheet.ExcelWorksheet.Cells[this.RowStartPosition, this.ColumnStartPosition,
                this.RowEndPosition, this.ColumnEndPosition];
        }

    }
}
