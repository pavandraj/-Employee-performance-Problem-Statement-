using ReadExcelData.Model;
using ReadExcelData.Repository.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelData.Business
{
    public class ConvertExcelToJson : IConvertExcelToJson
    {
        private IExcelData excelData;

        public ConvertExcelToJson(IExcelData _excelData)
        {
            excelData = _excelData;
        }
        public async Task<string> IConvertExcelToJson.ConvertExcelToJson()
        {
            var Converteddata = GetConvertedexceldata();

            //throw new NotImplementedException();
        }

        private List<ExcelHeaders> GetConvertedexceldata()
        {
            List<ExcelHeaders> excelHeaders = excelData.ReadDataFromExcel<ExcelHeaders>("https://docs.google.com/spreadsheets/d/1OziYYgstXPdyrZwWteejN6C3KZdXZppD/edit#gid=844951412").AsEnumerable().Select(s=> new ExcelHeaders
            {
                Date = (DateTime)s["Date"],
                ProjectName = s["ProjectName"].ToString().Trim(),
                Hours = (Double)s["Hours"],
                Owner = s["Owner"].ToString().Trim(),
                BillingStatus = s["BillingStatus"].ToString().Trim()
            }).ToList();
            // throw new NotImplementedException();
            return excelHeaders;
        }
    }
}
