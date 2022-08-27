using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelData.Business
{
    public interface IConvertExcelToJson
    {
        Task<string> ConvertExcelToJson();
    }
}
