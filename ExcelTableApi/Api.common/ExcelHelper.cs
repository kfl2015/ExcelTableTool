using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
//using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;//2007
namespace ExcelTableApi.Api.common
{
    public class ExcelHelper
    {
        /// <summary> 
        /// 获取窗体进程的唯一标识符.
        /// </summary>
        /// <param name="hwnd">窗体句柄.</param>
        /// <param name="ID">(输出)窗体进程的唯一标识符.</param>
        /// <returns></returns>
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        public ExcelHelper()
        { }
        public static void CreateExcel()
        {
            throw new Exception("不支持这个函数了哦。");
            //建立Excel对象
            //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //excel.Application.Workbooks.Add(true);
            //excel.Visible = true;
        }
        public static void CreateWord()
        {

        }
        /// <summary>
        /// 导出数据到指定文件并返回该文件完整路径名。如果不指定位置，则默认导出到临时目录。
        /// </summary>
        /// <param name="srcDataTable">要导出的数据源</param>
        /// <param name="fileName">要导出的数据源</param>
        /// <returns></returns>
        public static string ExportToExcel(DataTable srcDataTable, string fileName)
        {
            if (srcDataTable == null)
                return "";
            if (srcDataTable.Rows.Count == 0)
                return "";
            DataSet ds = new DataSet();
            ds.Tables.Add(srcDataTable);
            ExcelWriter ew = new ExcelWriter(ds);
            ew.WriteTo(fileName);
            return fileName;
        }

        /// <summary>
        /// 导出数据到指定文件并返回该文件完整路径名。如果不指定位置，则默认导出到临时目录。
        /// </summary>
        /// <param name="srcDataTable">要导出的数据源</param>
        /// <param name="fileName">要导出的数据源</param>
        /// <returns></returns>
        public static string ExportToExcelForPurchase(DataTable srcDataTable, string fileName, string pact, string contract, string plan, string aved, string money, string cnt, string ctime, string bid)
        {
            throw new Exception("不支持这个函数了哦。");

            #region 备注哦

            //if (srcDataTable == null)
            //    return "";
            //if (srcDataTable.Rows.Count == 0)
            //    return "";

            ////建立Excel对象
            //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //excel.Application.Workbooks.Add(true);
            //excel.Visible = false;
            //try
            //{
            //    Microsoft.Office.Interop.Excel.Range rg = excel.Range[excel.Cells[1, 1], excel.Cells[1, 9]];
            //    rg.Merge(rg.MergeCells);
            //    ((Microsoft.Office.Interop.Excel.Range)excel.Rows["1:1", Type.Missing]).RowHeight = 25;
            //    excel.Cells[1, 1] = "入库验收单";
            //    excel.Range[excel.Cells[1, 1], excel.Cells[1, 1]].HorizontalAlignment = 3;
            //    //生成表格列头名称
            //    for (int i = 0; i < srcDataTable.Columns.Count; i++)
            //    {
            //        // 合并单元格
            //        if (i == 2)
            //        {
            //            Microsoft.Office.Interop.Excel.Range rgg = excel.Range["C2", "D2"];
            //            rgg.Merge(rgg.MergeCells);
            //        }
            //        excel.Cells[2, i + (i > 2 ? 2 : 1)] = srcDataTable.Columns[i].ColumnName;
            //    }

            //    // 设置表格列头样式
            //    excel.Range[excel.Cells[2, 1], excel.Cells[1, srcDataTable.Columns.Count]].Select();
            //    excel.Range[excel.Cells[2, 1], excel.Cells[1, srcDataTable.Columns.Count]].Font.Bold = true;
            //    excel.Range[excel.Cells[2, 1], excel.Cells[1, srcDataTable.Columns.Count]].HorizontalAlignment = 3;
            //    excel.Range[excel.Cells[2, 1], excel.Cells[1, srcDataTable.Columns.Count]].VerticalAlignment = 2;
            //    ((Microsoft.Office.Interop.Excel.Range)excel.Rows["2:2", Type.Missing]).RowHeight = 25;
            //    decimal qm = 0.0M, mm = 0.0M;
            //    int qnum = 0, mnum = 0;
            //    int nowrows = 0;

            //    //填充表格数据
            //    for (int i = 0; i < srcDataTable.Rows.Count; i++)
            //    {
            //        // 修改表格数据，计算金额
            //        decimal tmpmoney = 0.0M;
            //        // 计算正品金额
            //        try
            //        {
            //            tmpmoney = decimal.Parse(srcDataTable.Rows[i]["正品数量"].ToString()) * decimal.Parse(srcDataTable.Rows[i]["采购价"].ToString());
            //            srcDataTable.Rows[i]["正品金额"] = tmpmoney.ToString();
            //            qm += tmpmoney;
            //        }
            //        catch (Exception ex)
            //        {
            //            Common.WriteLog(ex);
            //            srcDataTable.Rows[i]["正品金额"] = "0.00";
            //        }

            //        try
            //        {
            //            qnum += int.Parse(srcDataTable.Rows[i]["正品数量"].ToString());
            //        }
            //        catch (Exception ex)
            //        {
            //            Common.WriteLog(ex);
            //        }

            //        // 计算残品金额
            //        try
            //        {
            //            tmpmoney = decimal.Parse(srcDataTable.Rows[i]["残品数量"].ToString()) * decimal.Parse(srcDataTable.Rows[i]["采购价"].ToString());
            //            srcDataTable.Rows[i]["残品金额"] = tmpmoney.ToString();
            //            mm += tmpmoney;

            //        }
            //        catch (Exception ex)
            //        {
            //            Common.WriteLog(ex);
            //            srcDataTable.Rows[i]["残品金额"] = "0.00";
            //        }
            //        try
            //        {
            //            mnum += int.Parse(srcDataTable.Rows[i]["残品数量"].ToString());
            //        }
            //        catch (Exception ex)
            //        {
            //            Common.WriteLog(ex);
            //        }

            //        for (int j = 0; j < srcDataTable.Columns.Count; j++)
            //        {
            //            // 合并单元格
            //            if (j == 2)
            //            {
            //                rg = excel.Range["C" + (i + 3).ToString(), "D" + (i + 3).ToString()];
            //                rg.Merge(rg.MergeCells);
            //            }
            //            excel.Cells[i + 3, j + (j > 2 ? 2 : 1)] = srcDataTable.Rows[i][j].ToString();
            //        }

            //        nowrows = i + 4;
            //    }

            //    // 写入合计 行

            //    excel.Cells[nowrows, 1] = "合计";
            //    excel.Cells[nowrows, 7] = qnum.ToString();
            //    excel.Cells[nowrows, 8] = qm.ToString().Contains(".") ? qm.ToString() : qm.ToString() + ".00";
            //    excel.Cells[nowrows, 9] = mnum.ToString();
            //    excel.Cells[nowrows, 10] = mm.ToString().Contains(".") ? mm.ToString() : mm.ToString() + ".00";

            //    //实现插入行
            //    Microsoft.Office.Interop.Excel.Range rag = ((Microsoft.Office.Interop.Excel.Range)excel.Rows["2:2", Type.Missing]);
            //    rag.Select();
            //    rag.Rows.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            //    rag.Rows.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            //    rag.Rows.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            //    rag.Rows.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            //    // 设置表格列头颜色
            //    excel.Cells[2, 1] = "LieBo采购单 " + bid;
            //    excel.Cells[3, 1] = "供应商：" + pact;
            //    excel.Cells[4, 4] = "合同编号：" + contract;
            //    excel.Cells[3, 5] = "总件数：" + cnt;
            //    excel.Cells[3, 8] = "总金额：" + money;
            //    excel.Cells[4, 1] = "审核时间：" + ctime;
            //    excel.Cells[4, 5] = "计划到货时间：" + plan;
            //    excel.Cells[4, 8] = "实际到货时间：" + aved;

            //    //rag = ((Microsoft.Office.Interop.Excel.Range)excel.Rows["2:2", Type.Missing]);
            //    //rag.Select();
            //    //rag.EntireColumn.AutoFit();
            //}
            //catch (Exception ex)
            //{
            //    Common.WriteLog(ex);

            //}
            //if (string.IsNullOrEmpty(fileName))
            //{
            //    fileName = Path.GetTempFileName().Trim().Trim('\\') + "\\" + ".xls";
            //}
            //try
            //{
            //    ((Microsoft.Office.Interop.Excel.Worksheet)excel.ActiveSheet).SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel7, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //}
            //catch (Exception ex)
            //{
            //    Common.WriteLog(ex); fileName = "";
            //}
            //finally
            //{
            //    try
            //    {
            //        excel.Application.Quit();
            //        //excel.Quit();
            //        excel = null;
            //        IntPtr t = new IntPtr(excel.Hwnd);
            //        int k = 0;
            //        GetWindowThreadProcessId(t, out k);
            //        if (k > 10)
            //        {
            //            Process p = Process.GetProcessById(k);
            //            p.Kill();
            //        }
            //        GC.Collect();
            //    }
            //    catch (Exception ex)
            //    {
            //        Common.WriteLog(ex);
            //    }
            //}
            //return fileName;
            #endregion
        }

        /// <summary>
        /// 从给定的Excel文件获取Datatable对象。默认从Excel的Sheet1中读取。
        /// </summary>
        /// <param name="fileName">Excel文件路径</param>
        /// <exception cref="FileNotFoundException">给定的Excel文件不存在。</exception>
        /// <exception cref="ArgumentException">给定的文件不是Excel格式。</exception>
        /// <returns></returns>
        public static DataTable GetDataTableFromExcel(string fileName)
        {
            if (!File.Exists(fileName))
            {
                throw new FileNotFoundException("给定的Excel文件不存在。");
            }
            string ext = Path.GetExtension(fileName).Trim('.').ToLower();
            if (!".xls.xlsx".Contains(ext))
            {
                throw new ArgumentException("给定的文件不是Excel格式。");
            }
            DataTable dt = new DataTable();
            try
            {
                dt = ImportFromXls(fileName);
            }
            catch (Exception ex)
            {
                //LogNet.Log.Error(ex.Message, ex);
                throw ex;
            }
            return dt;
        }

        /// <summary>
        /// 从给定的Excel文件获取Datatable对象。默认从Excel的Sheet1中读取。
        /// </summary>
        /// <param name="fileName">Excel文件路径</param>
        /// <param name="convertColumn">是否从第一行提取为列头</param>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException">给定的Excel文件不存在。</exception>
        /// <exception cref="ArgumentException">给定的文件不是Excel格式。</exception>
        public static DataTable GetDataTableFromExcel(string fileName, bool convertColumn)
        {
            DataTable table = GetDataTableFromExcel(fileName);
            if (convertColumn && table != null && table.Rows.Count > 0)
            {
                int i = 1;
                foreach (DataColumn dc in table.Columns)
                {
                    string dcName = table.Rows[0][dc.ColumnName].ToString();
                    dc.ColumnName = string.IsNullOrWhiteSpace(dcName) ? ("Default" + (i++).ToString()) : dcName;
                }
                try
                {
                    table.Rows.RemoveAt(0);
                }
                catch (Exception ex)
                {
                    //LogNet.Log.Error(ex.Message, ex);
                }
            }
            return table;
        }

        /// <summary>
        /// 天啊，万能的读取Excel啊
        /// </summary>
        /// <param name="path"></param>
        /// <remarks>yumi</remarks>
        /// <returns></returns>
        private static DataTable ImportFromXls(string path)
        {
            IWorkbook workbook = null;
            if (path.ToLower().EndsWith(".xls"))
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    workbook = new HSSFWorkbook(file);
                }
            }
            else
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(file);
                }
            }
            ISheet sheet = workbook.GetSheetAt(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            DataTable dt = new DataTable();
            for (int j = 0; j < 1; j++)
            {
                dt.Columns.Add("C1");
            }
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                int cCount = row.LastCellNum > row.Cells.Count ? row.LastCellNum : row.Cells.Count;
                int cy = cCount - dt.Columns.Count;
                int nowCount = dt.Columns.Count;
                if (cy > 0)
                {
                    for (int j = nowCount + 1; j < cCount + 1; j++)
                    {
                        dt.Columns.Add("C" + j);
                    }
                }
                DataRow dr = dt.NewRow();
                for (int i = 0; i < row.LastCellNum; i++)
                {
                    ICell cell = row.GetCell(i);
                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        dr[i] = cell.ToString();
                    }
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }

    /// <summary>
    /// Excel写帮助
    /// </summary>
    public class ExcelWriter
    {
        DataSet _excelDataSet = null;

        public DataSet ExcelDataSet
        {
            get { return _excelDataSet; }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("ExcelDataSet为空。");
                }
                _excelDataSet = value;
            }
        }

        public ExcelWriter(DataSet excelDataSet)
        {
            if (excelDataSet == null)
            {
                throw new ArgumentNullException("excelDataSet为空。");
            }
            _excelDataSet = excelDataSet;
        }

        public void WriteTo(string fileName)
        {
            if (_excelDataSet == null)
            {
                throw new ArgumentNullException("ExcelDataSet为空。");
            }
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentNullException("fileName为空。");
            }
            else if (!fileName.ToLower().EndsWith(".xls") && !fileName.ToLower().EndsWith(".xlsx"))
            {
                fileName = fileName + ".xls";
            }
            string path = Path.GetDirectoryName(fileName);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            using (FileStream swExcel = new FileStream(fileName, FileMode.Create))
            {
                byte[] bys = Encoding.UTF8.GetBytes(GetExcelHeader());
                swExcel.Write(bys, 0, bys.Length);
                if (_excelDataSet.Tables.Count <= 0)
                {
                    bys = Encoding.UTF8.GetBytes(GetWorksheetHeader(-1));
                    swExcel.Write(bys, 0, bys.Length);
                    bys = Encoding.UTF8.GetBytes(GetRowData(new List<string>() { "" }));
                    swExcel.Write(bys, 0, bys.Length);
                    bys = Encoding.UTF8.GetBytes(GetWorksheetFooter());
                    swExcel.Write(bys, 0, bys.Length);
                }
                else
                {
                    for (int i = 0; i < _excelDataSet.Tables.Count; i++)
                    {
                        bys = Encoding.UTF8.GetBytes(GetWorksheetHeader(i));
                        swExcel.Write(bys, 0, bys.Length);
                        //swExcel.Write(GetWorksheetHeader(i));
                        bys = Encoding.UTF8.GetBytes(GetWorksheetColumnsSet(i));
                        swExcel.Write(bys, 0, bys.Length);
                        //swExcel.Write(GetWorksheetColumnsSet(i));
                        bys = Encoding.UTF8.GetBytes(GetWorksheetColumns(i));
                        swExcel.Write(bys, 0, bys.Length);
                        //swExcel.Write(GetWorksheetColumns(i));
                        foreach (DataRow row in _excelDataSet.Tables[i].Rows)
                        {
                            bys = Encoding.UTF8.GetBytes(GetRowData(row));
                            swExcel.Write(bys, 0, bys.Length);
                            //swExcel.Write(GetRowData(row));
                        }
                        bys = Encoding.UTF8.GetBytes(GetWorksheetFooter());
                        swExcel.Write(bys, 0, bys.Length);
                        //swExcel.Write(GetWorksheetFooter());
                    }
                }
                bys = Encoding.UTF8.GetBytes(GetExcelFooter());
                swExcel.Write(bys, 0, bys.Length);
                //swExcel.Write(GetExcelFooter());
                swExcel.Flush();
                swExcel.Close();
            }
        }

        public string GetExcelContent()
        {
            if (_excelDataSet == null)
            {
                throw new ArgumentNullException("ExcelDataSet为空。");
            }
            StringBuilder sbExcel = new StringBuilder();
            sbExcel.Append(GetExcelHeader());
            if (_excelDataSet.Tables.Count <= 0)
            {
                sbExcel.Append(GetWorksheetHeader(-1));
                sbExcel.Append(GetRowData(new List<string>() { "" }));
                sbExcel.Append(GetWorksheetFooter());
            }
            else
            {
                for (int i = 0; i < _excelDataSet.Tables.Count; i++)
                {
                    sbExcel.Append(GetWorksheetHeader(i));
                    sbExcel.Append(GetWorksheetColumnsSet(i));
                    foreach (DataRow row in _excelDataSet.Tables[i].Rows)
                    {
                        sbExcel.Append(GetRowData(row));
                    }
                    sbExcel.Append(GetWorksheetFooter());
                }
            }
            sbExcel.Append(GetExcelFooter());
            return sbExcel.ToString();
        }

        private string GetExcelHeader()
        {
            return string.Format(@"
<?xml version={0}1.0{0} encoding={0}utf-8{0}?>
<?mso-application progid={0}Excel.Sheet{0}?>
<Workbook xmlns:ss={0}urn:schemas-microsoft-com:office:spreadsheet{0} xmlns={0}urn:schemas-microsoft-com:office:spreadsheet{0}>
  <OfficeDocumentSettings xmlns={0}urn:schemas-microsoft-com:office:office{0} />
<DocumentProperties xmlns={0}urn:schemas-microsoft-com:office:office{0}>
    <Author>yumi</Author>
    <Created>{1}</Created>
  </DocumentProperties>
  <ExcelWorkbook xmlns={0}urn:schemas-microsoft-com:office:excel{0}><WindowWidth>1024</WindowWidth><WindowHeight>768</WindowHeight><ProtectStructure>False</ProtectStructure><ProtectWindows>False</ProtectWindows></ExcelWorkbook>
  <Styles>
    <Style ss:ID={0}Default{0}>
      <Font ss:FontName={0}宋体{0} ss:Size={0}10{0} ss:Color={0}#000000{0}/>
    </Style>
  </Styles>
", "\"", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        }

        private string GetWorksheetHeader(int dataTableIndex)
        {
            int columnCount = 1, rowCount = 1;
            if (dataTableIndex >= 0)
            {
                columnCount = _excelDataSet.Tables[dataTableIndex].Columns.Count;
                rowCount = _excelDataSet.Tables[dataTableIndex].Rows.Count + 1;
            }
            return string.Format(@"<ss:Worksheet ss:Name={0}{1}{0}>
<Table>
", "\"", dataTableIndex >= 0 ? _excelDataSet.Tables[dataTableIndex].TableName : "Sheet1");
        }

        private string GetWorksheetColumnsSet(int dataTableIndex)
        {
            StringBuilder sbColumns = new StringBuilder();
            foreach (DataColumn dc in _excelDataSet.Tables[dataTableIndex].Columns)
            {
                sbColumns.AppendFormat("<Column ss:AutoFit={0}0{0} ss:Width={0}75{0} />", "\"");
            }
            return sbColumns.ToString();
        }

        private string GetWorksheetColumns(int dataTableIndex)
        {
            var items = from f in _excelDataSet.Tables[dataTableIndex].Columns.Cast<DataColumn>()
                        select f.ColumnName;
            return GetRowData(items.ToList());
        }

        private string GetRowData(List<string> cellsContent)
        {
            StringBuilder sbRow = new StringBuilder();
            cellsContent.ForEach(item =>
            {
                sbRow.Append(GetCellString(item));
            });
            return string.Format("<Row ss:AutoFitHeight=\"0\">{0}</Row>", sbRow.ToString());
        }

        private string GetRowData(DataRow rowData)
        {
            StringBuilder sbRow = new StringBuilder();
            foreach (var item in rowData.ItemArray)
            {
                sbRow.Append(GetCellString(item));
            }
            return string.Format("<Row ss:AutoFitHeight=\"0\">{0}</Row>", sbRow.ToString());
        }

        private string GetCellString(object cellContent)
        {
            return string.Format("<Cell>{0}</Cell>", GetCellData(cellContent));
        }

        private string GetCellData(object content)
        {
            if (content == null)
                return "<Data ss:Type=\"String\"></Data>";
            return string.Format("<Data ss:Type=\"String\">{0}</Data>",
                content.ToString().Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("'", "&apos;")
                .Replace("\"", "&quot;")
                .Replace(" ", "&nbsp;")
                .Replace("©", "&copy;")
                .Replace("®", "&reg;"));
        }

        private string GetWorksheetFooter()
        {
            return string.Format(@"
</Table>
</ss:Worksheet>
", "\"");
        }

        private string GetExcelFooter()
        {
            return "</Workbook>";
        }
    }
}