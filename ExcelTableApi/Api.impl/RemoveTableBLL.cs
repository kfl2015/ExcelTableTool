using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTableApi.Api.service;
using TestPlatformManager;

namespace ExcelTableApi.Api.impl
{
  public  class RemoveTableBLL : IRemoveTableBLL
    {
      public DataTable GetExcelTable(String fileName, bool convertColumn)
      {
          return ExcelHelper.GetDataTableFromExcel(fileName, convertColumn);
      }
    }
}
