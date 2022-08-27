using ExcelDataReader;
using ReadExcelData.Repository.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;

namespace ReadExcelData.Repository
{
    public class ExcelData : IExcelData
    {
        public DataTable ReadDataFromExcel<T>(string FilePath)
        {
            try
            {
                IExcelDataReader excelDataReader = ExcelDataas(FilePath);
                DataSet result = excelDataReader.AsDataSet();
                DataTable dt = result.Tables[0];
                Type temp = typeof(T);
                int i = 0;
                foreach(PropertyInfo pro in temp.GetProperties())
                {
                    var z = "Column" + i.ToString();
                    i++;
                    dt.Columns[z].ColumnName = pro.Name;
                }

                return dt;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
           // throw new NotImplementedException();
        }

        private IExcelDataReader ExcelDataas(string filePath)
        {
            try
            {
                Console.WriteLine("Reading data from "+filePath);
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelDataReader;

                if(Path.GetExtension(filePath).ToUpper()==".XLSX")
                {
                    excelDataReader = ExcelReaderFactory.CreateReader(stream);
                }
                else
                {
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    excelDataReader = ExcelReaderFactory.CreateReader(stream);
                }

                Console.WriteLine("data Reading Completed");
                return excelDataReader;

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
            //throw new NotImplementedException();
        }
    }
}
