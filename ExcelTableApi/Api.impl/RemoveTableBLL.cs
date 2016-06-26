using System;
using System.Collections.Generic;
using System.Data;
using System.Net.Mime;
using System.Text;
using ExcelTableApi.Api.common;
using System.Threading.Tasks;
using ExcelTableApi.Api.service;


namespace ExcelTableApi.Api.impl
{
  public  class RemoveTableBLL : IRemoveTableBLL
    {
      public DataTable GetExcelTable(String fileName, bool convertColumn)
      {
          return ExcelHelper.GetDataTableFromExcel(fileName, convertColumn);
      }

      public List<DataTable> GetGroupSendMemberTodt(DataTable dt, out int currentCount)
      {
          List<DataTable> dtList = new List<DataTable>();
          Boolean[] flag = new Boolean[dt.Rows.Count];
          currentCount = 0;

          for (int j = 0; j < dt.Rows.Count; j++)
          {
              DataTable dtNew;
              if (j==0)
              {
                   dtNew = dt.Clone();
                  //初始化，取到的数据是第一行的判断逻辑
                  
             
                      dtNew.ImportRow(dt.Rows[j]);
                      flag[0] = true;
                      currentCount++;
              }
              else 
              {
                  if (flag[j] == false)
                  {
                      dtNew = dt.Clone();
                      currentCount++;
                  }
                  else
                  {
                      currentCount++;
                      continue;
                  }
              }
                   for (int i = 1; i < dt.Rows.Count; i++)
                   {
                        if (dt.Rows[j]["收/派件员"].ToString().Trim().Equals(dt.Rows[i]["收/派件员"].ToString().Trim()) && flag[i] == false)
                           {
                               dtNew.ImportRow(dt.Rows[i]);
                               flag[i] = true;
                           }
                           else
                           {
                               if (i + 1 == dt.Rows.Count)
                               {
                                   dtList.Add(dtNew);
                               }
                               continue;
                           }
                           if (i + 1 == dt.Rows.Count)
                           {
                               dtList.Add(dtNew);
                           }
                       }
              }
          return dtList;
      }

      public bool ExportToExcel(List<DataTable> tableList, out int filecount)
      {
          bool isSuccess = true;
         
           filecount = 0;
          for (int i = 0; i < tableList.Count; i++)
          {
              ExcelHelper.ExportToExcel(tableList[i], @"..\\Debug\\拆表结果\\" + tableList[i].Rows[0]["收/派件员"].ToString());
              filecount++;
          }
          if (filecount == tableList.Count)
          {
              return isSuccess;
          }
          else
          {
              isSuccess = false;
              return isSuccess;
          }
           
      }
    }
}
