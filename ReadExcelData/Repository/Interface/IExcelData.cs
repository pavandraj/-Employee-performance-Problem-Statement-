using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ReadExcelData.Repository.Interface
{
    public interface IExcelData
    {
        DataTable ReadDataFromExcel<T>(string FileName);
    }
}
