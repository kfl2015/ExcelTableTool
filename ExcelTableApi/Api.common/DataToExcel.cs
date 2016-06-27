using System;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Collections;
using System.Diagnostics;
using System.IO;

using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;


namespace ExcelTableApi.Api.common
{
    /// <summary>
    /// ����EXCEL�������ݱ�������
    /// </summary>
    public class DataToExcel
    {
        public DataToExcel()
        {
        }

        const string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=Excel 8.0;";

        /// <summary>
        /// ����EXECL ����
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="excelPath"></param>
        /// <returns></returns>
        public static string DataTableToExcel(System.Data.DataTable dt, string excelPath)
        {


            if (dt == null)
            {
                return "DataTable����Ϊ��";
            }

            //����sheet1����
            dt.TableName = "Sheet1";

            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;
            StringBuilder sb;
            string connString;

            if (rows == 0)
            {
                return "û������";
            }

            sb = new StringBuilder();
            connString = string.Format(ConnectionString, excelPath);

            //���ɴ������Ľű�
            sb.Append("CREATE TABLE ");
            sb.Append(dt.TableName + " ( ");

            for (int i = 0; i < cols; i++)
            {
                if (i < cols - 1)
                    sb.Append(string.Format("{0} nvarchar,", dt.Columns[i].ColumnName));
                else
                    sb.Append(string.Format("{0} nvarchar)", dt.Columns[i].ColumnName));
            }

            using (OleDbConnection objConn = new OleDbConnection(connString))
            {
                OleDbCommand objCmd = new OleDbCommand();
                objCmd.Connection = objConn;

                objCmd.CommandText = sb.ToString();

                try
                {
                    objConn.Open();
                    objCmd.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    return "��Excel�д�����ʧ�ܣ�������Ϣ��" + e.Message;
                }

                #region ���ɲ������ݽű�
                sb.Remove(0, sb.Length);
                sb.Append("INSERT INTO ");
                sb.Append(dt.TableName + " ( ");

                for (int i = 0; i < cols; i++)
                {
                    if (i < cols - 1)
                        sb.Append(dt.Columns[i].ColumnName + ",");
                    else
                        sb.Append(dt.Columns[i].ColumnName + ") values (");
                }

                for (int i = 0; i < cols; i++)
                {
                    if (i < cols - 1)
                        sb.Append("@" + dt.Columns[i].ColumnName + ",");
                    else
                        sb.Append("@" + dt.Columns[i].ColumnName + ")");
                }
                #endregion


                //�������붯����Command
                objCmd.CommandText = sb.ToString();
                OleDbParameterCollection param = objCmd.Parameters;

                for (int i = 0; i < cols; i++)
                {
                    param.Add(new OleDbParameter("@" + dt.Columns[i].ColumnName, OleDbType.VarChar));
                }

                //����DataTable�����ݲ����½���Excel�ļ���
                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < param.Count; i++)
                    {
                        param[i].Value = row[i];
                    }

                    objCmd.ExecuteNonQuery();
                }

                return "�����ѳɹ�����Excel";
            }//end using
        }


        /// <summary> 
        /// ��ȡExcel�ĵ� 
        /// </summary> 
        /// <param name="Path">�ļ�����</param> 
        /// <returns>����һ�����ݼ�</returns> 
        public DataSet ExcelToDS(string Path)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\";";
            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            System.Data.OleDb.OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select * from [Sheet1$]";
            myCommand = new System.Data.OleDb.OleDbDataAdapter(strExcel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            conn.Close();
            return ds;
        }

        #region ����EXCEL��һ����(��ҪExcel.dll֧��)

        private int titleColorindex = 15;
        /// <summary>
        /// ���ⱳ��ɫ
        /// </summary>
        public int TitleColorIndex
        {
            set { titleColorindex = value; }
            get { return titleColorindex; }
        }

        private DateTime beforeTime;			//Excel����֮ǰʱ��
        private DateTime afterTime;				//Excel����֮��ʱ��

        #region ����һ��Excelʾ��
        /// <summary>
        /// ����һ��Excelʾ��
        /// </summary>
        public void CreateExcel()
        {
            //Excel.Application excel = new Excel.Application();
            //excel.Application.Workbooks.Add(true);
            //excel.Cells[1, 1] = "��1�е�1��";
            //excel.Cells[1, 2] = "��1�е�2��";
            //excel.Cells[2, 1] = "��2�е�1��";
            //excel.Cells[2, 2] = "��2�е�2��";
            //excel.Cells[3, 1] = "��3�е�1��";
            //excel.Cells[3, 2] = "��3�е�2��";

            ////����
            //excel.ActiveWorkbook.SaveAs("./tt.xls", XlFileFormat.xlExcel9795, null, null, false, false, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
            ////����ʾ
            //excel.Visible = true;
            ////			excel.Quit();
            ////			excel=null;            
            ////			GC.Collect();//��������
        }
        #endregion

        #region ��DataTable�����ݵ�����ʾΪ����
        /// <summary>
        /// ��DataTable�����ݵ�����ʾΪ����
        /// </summary>
        /// <param name="dt">Ҫ����������</param>
        /// <param name="strTitle">���������ı���</param>
        /// <param name="FilePath">�����ļ���·��</param>
        /// <returns></returns>
        //public string OutputExcel(System.Data.DataTable dt, string strTitle, string FilePath)
        //{
        //    beforeTime = DateTime.Now;

        //    Excel.Application excel;
        //    Excel._Workbook xBk;
        //    Excel._Worksheet xSt;

        //    int rowIndex = 4;
        //    int colIndex = 1;

        //    excel = new Excel.ApplicationClass();
        //    xBk = excel.Workbooks.Add(true);
        //    xSt = (Excel._Worksheet)xBk.ActiveSheet;

        //    //ȡ���б���			
        //    foreach (DataColumn col in dt.Columns)
        //    {
        //        colIndex++;
        //        excel.Cells[4, colIndex] = col.ColumnName;

        //        //���ñ����ʽΪ���ж���
        //        xSt.get_Range(excel.Cells[4, colIndex], excel.Cells[4, colIndex]).Font.Bold = true;
        //        xSt.get_Range(excel.Cells[4, colIndex], excel.Cells[4, colIndex]).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
        //        xSt.get_Range(excel.Cells[4, colIndex], excel.Cells[4, colIndex]).Select();
        //        xSt.get_Range(excel.Cells[4, colIndex], excel.Cells[4, colIndex]).Interior.ColorIndex = titleColorindex;//19;//����Ϊǳ��ɫ��������56��
        //    }


        //    //ȡ�ñ����е�����			
        //    foreach (DataRow row in dt.Rows)
        //    {
        //        rowIndex++;
        //        colIndex = 1;
        //        foreach (DataColumn col in dt.Columns)
        //        {
        //            colIndex++;
        //            if (col.DataType == System.Type.GetType("System.DateTime"))
        //            {
        //                excel.Cells[rowIndex, colIndex] = (Convert.ToDateTime(row[col.ColumnName].ToString())).ToString("yyyy-MM-dd");
        //                xSt.get_Range(excel.Cells[rowIndex, colIndex], excel.Cells[rowIndex, colIndex]).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;//���������͵��ֶθ�ʽΪ���ж���
        //            }
        //            else
        //                if (col.DataType == System.Type.GetType("System.String"))
        //                {
        //                    excel.Cells[rowIndex, colIndex] = "'" + row[col.ColumnName].ToString();
        //                    xSt.get_Range(excel.Cells[rowIndex, colIndex], excel.Cells[rowIndex, colIndex]).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;//�����ַ��͵��ֶθ�ʽΪ���ж���
        //                }
        //                else
        //                {
        //                    excel.Cells[rowIndex, colIndex] = row[col.ColumnName].ToString();
        //                }
        //        }
        //    }

        //    //����һ���ϼ���			
        //    int rowSum = rowIndex + 1;
        //    int colSum = 2;
        //    excel.Cells[rowSum, 2] = "�ϼ�";
        //    xSt.get_Range(excel.Cells[rowSum, 2], excel.Cells[rowSum, 2]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        //    //����ѡ�еĲ��ֵ���ɫ			
        //    xSt.get_Range(excel.Cells[rowSum, colSum], excel.Cells[rowSum, colIndex]).Select();
        //    //xSt.get_Range(excel.Cells[rowSum,colSum],excel.Cells[rowSum,colIndex]).Interior.ColorIndex =Assistant.GetConfigInt("ColorIndex");// 1;//����Ϊǳ��ɫ��������56��

        //    //ȡ�����������ı���			
        //    excel.Cells[2, 2] = strTitle;

        //    //�������������ı����ʽ			
        //    xSt.get_Range(excel.Cells[2, 2], excel.Cells[2, 2]).Font.Bold = true;
        //    xSt.get_Range(excel.Cells[2, 2], excel.Cells[2, 2]).Font.Size = 22;

        //    //���ñ�������Ϊ����Ӧ����			
        //    xSt.get_Range(excel.Cells[4, 2], excel.Cells[rowSum, colIndex]).Select();
        //    xSt.get_Range(excel.Cells[4, 2], excel.Cells[rowSum, colIndex]).Columns.AutoFit();

        //    //�������������ı���Ϊ���о���			
        //    xSt.get_Range(excel.Cells[2, 2], excel.Cells[2, colIndex]).Select();
        //    xSt.get_Range(excel.Cells[2, 2], excel.Cells[2, colIndex]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenterAcrossSelection;

        //    //���Ʊ߿�			
        //    xSt.get_Range(excel.Cells[4, 2], excel.Cells[rowSum, colIndex]).Borders.LineStyle = 1;
        //    xSt.get_Range(excel.Cells[4, 2], excel.Cells[rowSum, 2]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;//��������߼Ӵ�
        //    xSt.get_Range(excel.Cells[4, 2], excel.Cells[4, colIndex]).Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;//�����ϱ��߼Ӵ�
        //    xSt.get_Range(excel.Cells[4, colIndex], excel.Cells[rowSum, colIndex]).Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;//�����ұ��߼Ӵ�
        //    xSt.get_Range(excel.Cells[rowSum, 2], excel.Cells[rowSum, colIndex]).Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;//�����±��߼Ӵ�



        //    afterTime = DateTime.Now;

        //    //��ʾЧ��			
        //    //excel.Visible=true;			
        //    //excel.Sheets[0] = "sss";

        //    ClearFile(FilePath);
        //    string filename = DateTime.Now.ToString("yyyyMMddHHmmssff") + ".xls";
        //    excel.ActiveWorkbook.SaveAs(FilePath + filename, Excel.XlFileFormat.xlExcel9795, null, null, false, false, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);

        //    //wkbNew.SaveAs strBookName;
        //    //excel.Save(strExcelFileName);

        //    #region  ����Excel����

        //    //��Ҫ��Excel��DCOM�����������:dcomcnfg


        //    //excel.Quit();
        //    //excel=null;            

        //    xBk.Close(null, null, null);
        //    excel.Workbooks.Close();
        //    excel.Quit();


        //    //ע�⣺�����õ�������Excel����Ҫִ����������������������Excel����
        //    //			if(rng != null)
        //    //			{
        //    //				System.Runtime.InteropServices.Marshal.ReleaseComObject(rng);
        //    //				rng = null;
        //    //			}
        //    //			if(tb != null)
        //    //			{
        //    //				System.Runtime.InteropServices.Marshal.ReleaseComObject(tb);
        //    //				tb = null;
        //    //			}
        //    if (xSt != null)
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xSt);
        //        xSt = null;
        //    }
        //    if (xBk != null)
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xBk);
        //        xBk = null;
        //    }
        //    if (excel != null)
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        //        excel = null;
        //    }
        //    GC.Collect();//��������
        //    #endregion

        //    return filename;

        //}
        #endregion

        #region Kill Excel����

        /// <summary>
        /// ����Excel����
        /// </summary>
        public void KillExcelProcess()
        {
            Process[] myProcesses;
            DateTime startTime;
            myProcesses = Process.GetProcessesByName("Excel");

            //�ò���Excel����ID����ʱֻ���жϽ�������ʱ��
            foreach (Process myProcess in myProcesses)
            {
                startTime = myProcess.StartTime;
                if (startTime > beforeTime && startTime < afterTime)
                {
                    myProcess.Kill();
                }
            }
        }
        #endregion

        #endregion

        #region ��DataTable�����ݵ�����ʾΪ����(��ʹ��Excel����ʹ��COM.Excel)

        #region ʹ��ʾ��
        /*ʹ��ʾ����
         * DataSet ds=(DataSet)Session["AdBrowseHitDayList"];
            string ExcelFolder=Assistant.GetConfigString("ExcelFolder");
            string FilePath=Server.MapPath(".")+"\\"+ExcelFolder+"\\";
			
            //�����е����Ķ�Ӧ��
            Hashtable nameList = new Hashtable();
            nameList.Add("ADID", "������");
            nameList.Add("ADName", "�������");
            nameList.Add("year", "��");
            nameList.Add("month", "��");
            nameList.Add("browsum", "��ʾ��");
            nameList.Add("hitsum", "�����");
            nameList.Add("BrowsinglIP", "����IP��ʾ");
            nameList.Add("HitsinglIP", "����IP���");
            //����excel����
            DataToExcel dte=new DataToExcel();
            string filename="";
            try
            {			
                if(ds.Tables[0].Rows.Count>0)
                {
                    filename=dte.DataExcel(ds.Tables[0],"����",FilePath,nameList);
                }
            }
            catch
            {
                //dte.KillExcelProcess();
            }
			
            if(filename!="")
            {
                Response.Redirect(ExcelFolder+"\\"+filename,true);
            }
         * 
         * */

        #endregion

         ///<summary>
         ///��DataTable�����ݵ�����ʾΪ����(��ʹ��Excel����)
         ///</summary>
         ///<param name="dt">����DataTable</param>
         ///<param name="strTitle">����</param>
         ///<param name="FilePath">�����ļ���·��</param>
         ///<param name="nameList"></param>
         ///<returns></returns>
        public string DataExcel(System.Data.DataTable dt, string strTitle, string FilePath, Hashtable nameList)
        {
            COM.Excel.cExcelFile excel = new COM.Excel.cExcelFile();
            ClearFile(FilePath);
            string filename = DateTime.Now.ToString("yyyyMMddHHmmssff") + ".xls";
            excel.CreateFile(FilePath + filename);
            excel.PrintGridLines = false;

            COM.Excel.cExcelFile.MarginTypes mt1 = COM.Excel.cExcelFile.MarginTypes.xlsTopMargin;
            COM.Excel.cExcelFile.MarginTypes mt2 = COM.Excel.cExcelFile.MarginTypes.xlsLeftMargin;
            COM.Excel.cExcelFile.MarginTypes mt3 = COM.Excel.cExcelFile.MarginTypes.xlsRightMargin;
            COM.Excel.cExcelFile.MarginTypes mt4 = COM.Excel.cExcelFile.MarginTypes.xlsBottomMargin;

            double height = 1.5;
            excel.SetMargin(ref mt1, ref height);
            excel.SetMargin(ref mt2, ref height);
            excel.SetMargin(ref mt3, ref height);
            excel.SetMargin(ref mt4, ref height);

            COM.Excel.cExcelFile.FontFormatting ff = COM.Excel.cExcelFile.FontFormatting.xlsNoFormat;
            string font = "����";
            short fontsize = 9;
            excel.SetFont(ref font, ref fontsize, ref ff);

            byte b1 = 1,
                b2 = 12;
            short s3 = 12;
            excel.SetColumnWidth(ref b1, ref b2, ref s3);

            string header = "ҳü";
            string footer = "ҳ��";
            excel.SetHeader(ref header);
            excel.SetFooter(ref footer);


            COM.Excel.cExcelFile.ValueTypes vt = COM.Excel.cExcelFile.ValueTypes.xlsText;
            COM.Excel.cExcelFile.CellFont cf = COM.Excel.cExcelFile.CellFont.xlsFont0;
            COM.Excel.cExcelFile.CellAlignment ca = COM.Excel.cExcelFile.CellAlignment.xlsCentreAlign;
            COM.Excel.cExcelFile.CellHiddenLocked chl = COM.Excel.cExcelFile.CellHiddenLocked.xlsNormal;

            // ��������
            int cellformat = 1;
            //			int rowindex = 1,colindex = 3;					
            //			object title = (object)strTitle;
            //			excel.WriteValue(ref vt, ref cf, ref ca, ref chl,ref rowindex,ref colindex,ref title,ref cellformat);

            int rowIndex = 1;//��ʼ��
            int colIndex = 0;



            //ȡ���б���				
            foreach (DataColumn colhead in dt.Columns)
            {
                colIndex++;
                string name = colhead.ColumnName.Trim();
                object namestr = (object)name;
                IDictionaryEnumerator Enum = nameList.GetEnumerator();
                while (Enum.MoveNext())
                {
                    if (Enum.Key.ToString().Trim() == name)
                    {
                        namestr = Enum.Value;
                    }
                }
                excel.WriteValue(ref vt, ref cf, ref ca, ref chl, ref rowIndex, ref colIndex, ref namestr, ref cellformat);
            }

            //ȡ�ñ����е�����			
            foreach (DataRow row in dt.Rows)
            {
                rowIndex++;
                colIndex = 0;
                foreach (DataColumn col in dt.Columns)
                {
                    colIndex++;
                    if (col.DataType == System.Type.GetType("System.DateTime"))
                    {
                        object str = (object)(Convert.ToDateTime(row[col.ColumnName].ToString())).ToString("yyyy-MM-dd"); ;
                        excel.WriteValue(ref vt, ref cf, ref ca, ref chl, ref rowIndex, ref colIndex, ref str, ref cellformat);
                    }
                    else
                    {
                        object str = (object)row[col.ColumnName].ToString();
                        excel.WriteValue(ref vt, ref cf, ref ca, ref chl, ref rowIndex, ref colIndex, ref str, ref cellformat);
                    }
                }
            }
            int ret = excel.CloseFile();

            //			if(ret!=0)
            //			{
            //				//MessageBox.Show(this,"Error!");
            //			}
            //			else
            //			{
            //				//MessageBox.Show(this,"����ļ�c:\\test.xls!");
            //			}
            return filename;

        }

        
        #endregion


        #region ��DataTable����ΪExcel(OleDb ��ʽ������


        /// <summary>
        /// ��DataTable����ΪExcel(OleDb ��ʽ������
        /// </summary>
        /// <param name="dataTable">��</param>
        /// <param name="fileName">����Ĭ���ļ���</param>
        public static void DataSetToExcel(DataTable dataTable, string fileName)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "xls files (*.xls)|*.xls";
            saveFileDialog.FileName = fileName;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = saveFileDialog.FileName;
                if (File.Exists(fileName))
                {
                    try
                    {
                        File.Delete(fileName);
                    }
                    catch
                    {
                        MessageBox.Show("���ļ�����ʹ����,�ر��ļ����������������ļ�����!");
                        return;
                    }
                }
                OleDbConnection oleDbConn = new OleDbConnection();
                OleDbCommand oleDbCmd = new OleDbCommand();
                string sSql = "";
                try
                {
                    oleDbConn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + fileName + @";Extended ProPerties=""Excel 8.0;HDR=Yes;""";
                    oleDbConn.Open();
                    oleDbCmd.CommandType = CommandType.Text;
                    oleDbCmd.Connection = oleDbConn;
                    sSql = "CREATE TABLE sheet1 (";
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        // �ֶ����Ƴ��ֹؼ��ֻᵼ�´���
                        if (i < dataTable.Columns.Count - 1)
                            sSql += "[" + dataTable.Columns[i].Caption + "] TEXT(100) ,";
                        else
                            sSql += "[" + dataTable.Columns[i].Caption + "] TEXT(200) )";
                    }
                    oleDbCmd.CommandText = sSql;
                    oleDbCmd.ExecuteNonQuery();
                    for (int j = 0; j < dataTable.Rows.Count; j++)
                    {
                        sSql = "INSERT INTO sheet1 VALUES('";
                        for (int i = 0; i < dataTable.Columns.Count; i++)
                        {
                            if (i < dataTable.Columns.Count - 1)
                                sSql += dataTable.Rows[j][i].ToString() + " ','";
                            else
                                sSql += dataTable.Rows[j][i].ToString() + " ')";
                        }
                        oleDbCmd.CommandText = sSql;
                        oleDbCmd.ExecuteNonQuery();
                    }
                    MessageBox.Show("����EXCEL�ɹ�");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("����EXCELʧ��:" + ex.Message);
                }
                finally
                {
                    oleDbCmd.Dispose();
                    oleDbConn.Close();
                    oleDbConn.Dispose();
                }
            }
        }
        #endregion
        #region �����ݵ�����Excel�ļ�

        /// <summary>
        /// �����ݵ�����Excel�ļ�
        /// </summary>
        /// <param name="Table">DataTable����</param>
        /// <param name="ExcelFilePath">Excel�ļ�·��</param>
        public static bool OutputToExcel(DataTable Table, string ExcelFilePath)
        {
            if (File.Exists(ExcelFilePath))
            {
                throw new Exception("���ļ��Ѿ����ڣ�");
            }

            if ((Table.TableName.Trim().Length == 0) || (Table.TableName.ToLower() == "table"))
            {
                Table.TableName = "Sheet1";
            }

            //���ݱ�������
            int ColCount = Table.Columns.Count;

            //���ڼ�����ʵ��������ʱ�����
            int i = 0;

            //��������
            OleDbParameter[] para = new OleDbParameter[ColCount];

            //�������ṹ��SQL���
            string TableStructStr = @"Create Table " + Table.TableName + "(";

            //�����ַ���
            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ExcelFilePath + ";Extended Properties=Excel 8.0;";
            OleDbConnection objConn = new OleDbConnection(connString);

            //�������ṹ
            OleDbCommand objCmd = new OleDbCommand();

            //�������ͼ���
            ArrayList DataTypeList = new ArrayList();
            DataTypeList.Add("System.Decimal");
            DataTypeList.Add("System.Double");
            DataTypeList.Add("System.Int16");
            DataTypeList.Add("System.Int32");
            DataTypeList.Add("System.Int64");
            DataTypeList.Add("System.Single");

            //�������ݱ��������У����ڴ������ṹ
            foreach (DataColumn col in Table.Columns)
            {
                //��������������У������ø��е���������Ϊdouble
                if (DataTypeList.IndexOf(col.DataType.ToString()) >= 0)
                {
                    para[i] = new OleDbParameter("@" + col.ColumnName, OleDbType.Double);
                    objCmd.Parameters.Add(para[i]);

                    //��������һ��
                    if (i + 1 == ColCount)
                    {
                        TableStructStr += col.ColumnName + " double)";
                    }
                    else
                    {
                        TableStructStr += col.ColumnName + " double,";
                    }
                }
                else
                {
                    para[i] = new OleDbParameter("@" + col.ColumnName, OleDbType.VarChar);
                    objCmd.Parameters.Add(para[i]);

                    //��������һ��
                    if (i + 1 == ColCount)
                    {
                        TableStructStr += col.ColumnName + " varchar)";
                    }
                    else
                    {
                        TableStructStr += col.ColumnName + " varchar,";
                    }
                }
                i++;
            }

            //����Excel�ļ����ļ��ṹ
            try
            {
                objCmd.Connection = objConn;
                objCmd.CommandText = TableStructStr;

                if (objConn.State == ConnectionState.Closed)
                {
                    objConn.Open();
                }
                objCmd.ExecuteNonQuery();
            }
            catch (Exception exp)
            {
                throw exp;
            }

            //�����¼��SQL���
            string InsertSql_1 = "Insert into " + Table.TableName + " (";
            string InsertSql_2 = " Values (";
            string InsertSql = "";

            //���������У����ڲ����¼���ڴ˴��������¼��SQL���
            for (int colID = 0; colID < ColCount; colID++)
            {
                if (colID + 1 == ColCount)  //���һ��
                {
                    InsertSql_1 += Table.Columns[colID].ColumnName + ")";
                    InsertSql_2 += "@" + Table.Columns[colID].ColumnName + ")";
                }
                else
                {
                    InsertSql_1 += Table.Columns[colID].ColumnName + ",";
                    InsertSql_2 += "@" + Table.Columns[colID].ColumnName + ",";
                }
            }

            InsertSql = InsertSql_1 + InsertSql_2;

            //�������ݱ�������������
            for (int rowID = 0; rowID < Table.Rows.Count; rowID++)
            {
                for (int colID = 0; colID < ColCount; colID++)
                {
                    if (para[colID].DbType == DbType.Double && Table.Rows[rowID][colID].ToString().Trim() == "")
                    {
                        para[colID].Value = 0;
                    }
                    else
                    {
                        para[colID].Value = Table.Rows[rowID][colID].ToString().Trim();
                    }
                }
                try
                {
                    objCmd.CommandText = InsertSql;
                    objCmd.ExecuteNonQuery();
                }
                catch (Exception exp)
                {
                    string str = exp.Message;
                }
            }
            try
            {
                if (objConn.State == ConnectionState.Open)
                {
                    objConn.Close();
                }
            }
            catch (Exception exp)
            {
                throw exp;
            }
            return true;
        }
        #endregion



        #region  ������ʱ��Excel�ļ�

        private void ClearFile(string FilePath)
        {
            String[] Files = System.IO.Directory.GetFiles(FilePath);
            if (Files.Length > 10)
            {
                for (int i = 0; i < 10; i++)
                {
                    try
                    {
                        System.IO.File.Delete(Files[i]);
                    }
                    catch
                    {
                    }

                }
            }
        }
        #endregion

          
    /// <summary>
    /// ����Excel
    /// </summary>
    public static void CreateExcels(DataTable dt, string path)
    {
      HSSFWorkbook workbook = new HSSFWorkbook();
      ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName);//����������

      
      IRow row = sheet.CreateRow(0);//�ڹ�����������һ��
      for (int i = 0; i < dt.Columns.Count; i++)
      {
        ICell cell = row.CreateCell(i);//����������һ��
        cell.SetCellValue(dt.Columns[i].ColumnName);//�����е�����	 
      }
    
      for (int i = 1; i <= dt.Rows.Count; i++)//����DataTable��
      {
        DataRow dataRow = dt.Rows[i - 1];
        row = sheet.CreateRow(i);//�ڹ�����������һ��

        for (int j = 0; j < dt.Columns.Count; j++)//����DataTable��
        {
          ICell cell = row.CreateCell(j);//����������һ��
          cell.SetCellValue(dataRow[j].ToString());//�����е�����	 
        }
      }
  
      MemoryStream ms = new MemoryStream();
      workbook.Write(ms);

      using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
      {
        byte[] bArr = ms.ToArray();
        fs.Write(bArr, 0, bArr.Length);
        fs.Flush();
      }
  }

    /// <summary>  
    /// ��DataTable���ݵ�����Excel�ļ���(xls)  
    /// </summary>  
    /// <param name="dt"></param>  
    /// <param name="file"></param>  
    public static void TableToExcelForXLS(DataTable dt, string file)
    {
        HSSFWorkbook hssfworkbook = new HSSFWorkbook();
        ISheet sheet = hssfworkbook.CreateSheet("Test");

        //��ͷ  
        IRow row = sheet.CreateRow(0);
        for (int i = 0; i < dt.Columns.Count; i++)
        {
            ICell cell = row.CreateCell(i);
            cell.SetCellValue(dt.Columns[i].ColumnName);
        }

        //����  
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            IRow row1 = sheet.CreateRow(i + 1);
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = row1.CreateCell(j);
                cell.SetCellValue(dt.Rows[i][j].ToString());
            }
        }

        //תΪ�ֽ�����  
        MemoryStream stream = new MemoryStream();
        hssfworkbook.Write(stream);
        var buf = stream.ToArray();

        //����ΪExcel�ļ�  
        using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
        {
            fs.Write(buf, 0, buf.Length);
            fs.Flush();
        }
    }



    }
}