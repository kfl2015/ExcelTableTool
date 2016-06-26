using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ExcelTableApi.Api.service
{
    public interface  IRemoveTableBLL
    {
        /// <summary>
       /// 获取datatable
       /// </summary>
       /// <param name="s"> </param>
       /// <returns></returns>
         DataTable GetExcelTable(String filName,bool convertColumn);
         /// <summary>
       /// 分组拆表
       /// </summary>
         /// <returns></returns>
         List<DataTable> GetGroupSendMemberTodt(DataTable dt, out int currentCount);

        /// <summary>
        /// 分组拆表
        /// </summary>
        /// <returns></returns>
        /// 
        bool ExportToExcel(List<DataTable> tableList,out int filecount);
    }
}
