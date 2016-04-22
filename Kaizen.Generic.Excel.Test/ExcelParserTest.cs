using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Kaizen.Generic.Excel.Test.Model;
using System.Collections.Generic;
using System.IO;
using System.Drawing;

namespace Kaizen.Generic.Excel.Test
{
    [TestClass]
    public class ExcelParserTest
    {
        [TestMethod]
        public void ConvertExcelToExcel()
        {
            ExcelFile excelFile = CreateExcelFile();

            var buffer = excelFile.GetFileStream();

            var streamWriter = new StreamWriter(string.Format(@"{0}\User.xlsx", Environment.CurrentDirectory));
            streamWriter.BaseStream.Write(buffer, 0, buffer.Length);
            streamWriter.Close();
        }

        private static ExcelFile CreateExcelFile()
        {
            var excelFile = new ExcelFile();
            AddExcelData(excelFile);
            AddExcelStyle(excelFile);
            return excelFile;
        }

        private static void AddExcelStyle(ExcelFile excelFile)
        {
            excelFile.GetSheet("User")
                .addAlternativeBackgroundColor(Color.Aquamarine, Color.Aqua)
                .addHeaderBackgroundColor(Color.Red);
            excelFile.GetSheet("Product").addBackgroundColor(Color.Red);
        }

        private static void AddExcelData(ExcelFile excelFile)
        {
            var userCollection = new List<User>()
            {
                new User("Rodrigo", "rodrigo.saint01@live.com"),
                new User("Jéssica", "jessica@live.com"),
                new User("Jéssica", "jessica@live.com")
            };
            excelFile.AddSheet<User>("User", userCollection);

            var productCollection = new List<Product>()
            {
                new Product("Bread", 10.5d),
                new Product("Cheese", 3.2d)
            };
            excelFile.AddSheet<Product>("Product", productCollection);
        }
    }
}
