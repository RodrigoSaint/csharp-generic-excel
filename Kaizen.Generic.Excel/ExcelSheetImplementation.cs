using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;

namespace Kaizen.Generic.Excel
{
    public class ExcelSheetImplementation<T> : ExcelSheet
    {
        protected ICollection<T> _collection;
        public ExcelWorksheet ExcelWorksheet { get; set; }
        protected List<PropertyInfo> _propertyCollection;
        protected ExcelRange _tableDataRange;
        protected ExcelRange _headerRange;
        protected SheetStyle _headerStyle;
        protected SheetStyle _tableDataStyle;

        public ExcelSheetImplementation(ExcelWorksheet excelWorksheet, ICollection<T> collection)
        {
            this.ExcelWorksheet = excelWorksheet;
            this._collection = collection;
            this._propertyCollection = ExcelSheetImplementation<T>.GetWriteblePropertyCollection();
            this.AddExcelData();
            this.AddExcelHeaders();
            this._headerStyle = new SheetStyle(this._headerRange, this);
            this._tableDataStyle = new SheetStyle(this._tableDataRange, this);
            this.SetExcelTableDefinition();
        }

        private void SetExcelTableDefinition()
        {
            var address = new ExcelAddressBase(1, 1, this._collection.Count + 1, this._propertyCollection.Count);
            this.ExcelWorksheet.Tables.Add(address, this.ExcelWorksheet.Name);
        }

        #region Data
        protected static List<PropertyInfo> GetWriteblePropertyCollection()
        {
            return typeof(T).GetProperties()
                .Where(x => x.PropertyType.Name == "String" ||
                (!x.PropertyType.IsClass && !x.PropertyType.IsGenericType))
                .ToList();
        }

        protected ExcelSheet AddExcelData()
        {
            for (int rowNumber = 0; rowNumber < this._collection.Count; rowNumber++)
            {
                var columnNumber = 0;
                foreach (var property in this._propertyCollection)
                {
                    this.ExcelWorksheet.Cells[rowNumber + 2, columnNumber + 1].Value =
                        GetPropertyValue(this._collection.ElementAt(rowNumber), property);
                    columnNumber++;
                }
            }
            AddExcelDateRange();
            return this;
        }

        protected ExcelSheet AddExcelHeaders()
        {
            for (int propertyNumber = 0; propertyNumber < this._propertyCollection.Count; propertyNumber++)
            {
                var property = this._propertyCollection[propertyNumber];
                var displayName = property.GetCustomAttributes(typeof(DisplayNameAttribute), true).FirstOrDefault() as DisplayNameAttribute;
                this.ExcelWorksheet.Cells[1, propertyNumber + 1].Value = displayName != null ? displayName.DisplayName : property.Name;
            }
            AddHeaderRange();
            return this;
        }

        protected ExcelSheet AddExcelDateRange()
        {
            this._tableDataRange = new ExcelRange(this._collection.Count, this._propertyCollection.Count);
            return this;
        }

        protected object GetPropertyValue(T row, PropertyInfo property)
        {
            object value = null;
            value = row
                .GetType()
                .GetProperty(property.Name)
                .GetValue(row);
            return value;
        }

        protected ExcelSheet AddHeaderRange()
        {
            this._headerRange = new ExcelRange(this._propertyCollection.Count);
            return this;
        } 
        #endregion

        #region Style


        public ExcelSheet addBackgroundColor(Color color)
        {
            this._tableDataStyle.BackgroundColor = color;
            this._tableDataStyle.ExecuteStyle();
            return this;
        }

        public ExcelSheet addHeaderBackgroundColor(Color color)
        {
            this._headerStyle.BackgroundColor = color;
            this._headerStyle.ExecuteStyle();
            return this;
        }

        public ExcelSheet addAlternativeBackgroundColor(Color color, Color alternativeColor)
        {
            this._tableDataStyle.BackgroundColor = color;
            this._tableDataStyle.AddAlternativeColor(alternativeColor);
            return this;
        }
        #endregion

    }
}
