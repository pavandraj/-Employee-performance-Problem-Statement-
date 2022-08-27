using ReadExcelData.Business;
using ReadExcelData.Interface;
using System;
using System.Collections.Generic;
using System.Text;

namespace ReadExcelData
{
    public class ReadExcel : IReadExcelData
    {
        private readonly IConvertExcelToJson readExcelDataBusiness;

        public ReadExcel(IConvertExcelToJson _readExcelDataBusiness)
        {
            readExcelDataBusiness = _readExcelDataBusiness;
        }
        public void ReadExcelData()
        {
            Console.WriteLine("Started Converting Excel data to Json");
            var result = readExcelDataBusiness.ConvertExcelToJson().Result;
            //throw new NotImplementedException();
        }
    }
}
