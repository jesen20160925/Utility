using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Jesen.Utility.Excel
{
    public enum ExportExcelFormat
    {
        Excel2003,
        Excel2007,
        ExcelBoth
    }
    class x2003
    {
        #region Excel2003
        /// <summary>
        /// 将Excel文件中的数据读出到DataTable中(xls)
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static DataTable ExcelToTableForXLS(string file)
        {
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook(fs);
                ISheet sheet = hssfworkbook.GetSheetAt(0);

                //表头
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueTypeForXLS(header.GetCell(i) as HSSFCell);
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        //continue;
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueTypeForXLS(sheet.GetRow(i).GetCell(j) as HSSFCell);
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 将DataTable数据导出到Excel文件中(xls)
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file"></param>
        public static void TableToExcelForXLS(DataTable dt, string file)
        {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            ISheet sheet = hssfworkbook.CreateSheet("Sheet1");

            //表头
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组
            MemoryStream stream = new MemoryStream();
            hssfworkbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }

        /// <summary>
        /// 将DataGridView数据导出到Excel文件中(xls)
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file"></param>
        public static void DataGridViewToExcelForXLS(DataGridView dgv, string file)
        {
            try
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook();
                ISheet sheet = hssfworkbook.CreateSheet(file.Substring(file.LastIndexOf("\\") + 1, file.LastIndexOf(".") - file.LastIndexOf("\\")));
                int column = 0;
                //表头
                IRow row = sheet.CreateRow(0);
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (dgv.Columns[i].Visible == true)
                    {
                        ICell cell = row.CreateCell(column);
                        cell.SetCellValue(dgv.Columns[i].HeaderText);
                        column++;
                    }

                }

                //数据

                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    //if (dgv.Rows[i].Cells[0] == null || dgv.Rows[i].Cells[0].Value == null || dgv.Rows[i].Cells[0].Value.ToString().Trim() == "")
                    //{
                    //    break;
                    //}


                    if (dgv.Rows[i].Visible == true)
                    {
                        IRow row1 = sheet.CreateRow(i + 1);
                        column = 0;
                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            if (dgv.Columns[j].Visible == true)
                            {
                                ICell cell = row1.CreateCell(column);
                                cell.SetCellValue(dgv.Rows[i].Cells[j].Value.ToString());
                                ++column;
                            }
                        }
                    }

                }

                //转为字节数组
                MemoryStream stream = new MemoryStream();
                hssfworkbook.Write(stream);
                var buf = stream.ToArray();

                //保存为Excel文件
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 获取单元格类型(xls)
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueTypeForXLS(HSSFCell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                //case CellType.BLANK: //BLANK:
                //    return null;
                //case CellType.BOOLEAN: //BOOLEAN:
                //    return cell.BooleanCellValue;
                //case CellType.NUMERIC: //NUMERIC:
                //    return cell.NumericCellValue;
                //case CellType.STRING: //STRING:
                //    return cell.StringCellValue;
                //case CellType.ERROR: //ERROR:
                //    return cell.ErrorCellValue;
                //case CellType.FORMULA: //FORMULA:
                default:
                    return "=" + cell.CellFormula;
            }
        }
        #endregion
    }

    class x2007
    {
        #region Excel2007
        /// <summary>
        /// 将Excel文件中的数据读出到DataTable中(xlsx)
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static DataTable ExcelToTableForXLSX(string file)
        {
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);
                ISheet sheet = xssfworkbook.GetSheetAt(0);

                //表头
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueTypeForXLSX(header.GetCell(i) as HSSFCell);
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        //continue;
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueTypeForXLSX(sheet.GetRow(i).GetCell(j) as HSSFCell);
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 将DataTable数据导出到Excel文件中(xlsx)
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file"></param>
        public static void TableToExcelForXLSX(DataTable dt, string file)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            ISheet sheet = xssfworkbook.CreateSheet(file.Substring(file.LastIndexOf("\\") + 1, file.LastIndexOf(".") - file.LastIndexOf("\\")));

            //表头
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组
            MemoryStream stream = new MemoryStream();
            xssfworkbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }

        /// <summary>
        /// 将DataGridView数据导出到Excel文件中(xlsx)
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file"></param>
        public static void DataGridViewToExcelForXLSX(DataGridView dgv, string file)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            ISheet sheet = xssfworkbook.CreateSheet(file.Substring(file.LastIndexOf("\\") + 1, file.LastIndexOf(".") - file.LastIndexOf("\\")));
            int column = 0;
            //表头
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                if (dgv.Columns[i].Visible == true)
                {
                    ICell cell = row.CreateCell(column);
                    cell.SetCellValue(dgv.Columns[i].HeaderText);
                    column++;
                }

            }

            //数据
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (dgv.Rows[i].Visible == true)
                {
                    IRow row1 = sheet.CreateRow(i + 1);
                    column = 0;
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        if (dgv.Columns[j].Visible == true)
                        {
                            ICell cell = row1.CreateCell(column);
                            cell.SetCellValue(dgv.Rows[i].Cells[j].Value.ToString());
                            column++;
                        }
                    }
                }
            }

            //转为字节数组
            MemoryStream stream = new MemoryStream();
            xssfworkbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }


        /// <summary>
        /// 获取单元格类型(xlsx)
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueTypeForXLSX(HSSFCell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                //case CellType.BLANK: //BLANK:
                //    return null;
                //case CellType.BOOLEAN: //BOOLEAN:
                //    return cell.BooleanCellValue;
                //case CellType.NUMERIC: //NUMERIC:
                //    return cell.NumericCellValue;
                //case CellType.STRING: //STRING:
                //    return cell.StringCellValue;
                //case CellType.ERROR: //ERROR:
                //    return cell.ErrorCellValue;
                //case CellType.FORMULA: //FORMULA:
                default:
                    return "=" + cell.CellFormula;
            }
        }
        #endregion
    }

    public class ExcelHelper
    {
        #region 导出

        /// <summary>
        /// 将DataTable转换成Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filepath"></param>
        public static void TableToExcel(DataTable dt, ExportExcelFormat eefExcelFormat, string sExpectedFileName)
        {
            if (eefExcelFormat == ExportExcelFormat.Excel2003)
            {
                x2003.TableToExcelForXLS(dt, sExpectedFileName);
            }
            else
            {
                x2007.TableToExcelForXLSX(dt, sExpectedFileName);
            }
        }
    
        /// <summary>
        /// 将DataGridView转换成Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filepath"></param>
        public static void DataGridViewToExcel(DataGridView dgv, ExportExcelFormat eefExcelFormat, string sExpectedFileName)
        {
            if (eefExcelFormat == ExportExcelFormat.Excel2003)
            {
                x2003.DataGridViewToExcelForXLS(dgv, sExpectedFileName);
            }
            else
            {
                x2007.DataGridViewToExcelForXLSX(dgv, sExpectedFileName);
            }
        }
    
        /// <summary>
        /// 导出EXCEL,可以导出多个sheet
        /// </summary>
        /// <param name="dtSources">原始数据数组类型</param>
        /// <param name="strFileName">路径</param>
        public static void DataTableArrayToExcel2007(DataTable[] dtSources, string strFileName)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            for (int k = 0; k < dtSources.Length; k++)
            {
                ISheet sheet = workbook.CreateSheet(dtSources[k].TableName.ToString());

                //填充表头
                IRow dataRow = sheet.CreateRow(0);
                foreach (DataColumn column in dtSources[k].Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                }

                //填充内容
                for (int i = 0; i < dtSources[k].Rows.Count; i++)
                {
                    dataRow = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dtSources[k].Columns.Count; j++)
                    {
                        dataRow.CreateCell(j).SetCellValue(dtSources[k].Rows[i][j].ToString());
                    }
                }
            }

            //保存
            using (MemoryStream ms = new MemoryStream())
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
            }
        }

        /// <summary>
        /// 导出EXCEL,可以导出多个sheet
        /// </summary>
        /// <param name="dtSources">原始数据数组类型</param>
        /// <param name="strFileName">路径</param>
        public static void DataTableArrayToExcel2003(DataTable[] dtSources, string strFileName)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            for (int k = 0; k < dtSources.Length; k++)
            {
                ISheet sheet = workbook.CreateSheet(dtSources[k].TableName.ToString());

                //填充表头
                IRow dataRow = sheet.CreateRow(0);
                foreach (DataColumn column in dtSources[k].Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                }

                //填充内容
                for (int i = 0; i < dtSources[k].Rows.Count; i++)
                {
                    dataRow = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dtSources[k].Columns.Count; j++)
                    {
                        dataRow.CreateCell(j).SetCellValue(dtSources[k].Rows[i][j].ToString());
                    }
                }
            }

            //保存
            using (MemoryStream ms = new MemoryStream())
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
            }
        } 

        #endregion

        #region 导入

        /// <summary>
        /// 将Excel文件导入到DataTable
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        public static DataTable ExcelImportToDataTable(string filePath)
        {
            DataTable retValue = null;

            string fileExt = Path.GetExtension(filePath);
            IWorkbook workbook = null;
            try
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (fileExt == ".xls")
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                    else if (fileExt == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(file);
                    }
                    if (workbook != null)
                    {
                        ISheet sheet = workbook.GetSheetAt(0);
                        retValue = ImportDt(sheet, 0, true);
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }

            return retValue;
        }

        /// <summary>
        /// 将制定sheet中的数据导出到datatable中
        /// </summary>
        /// <param name="sheet">需要导出的sheet</param>
        /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
        /// <returns></returns>
        static DataTable ImportDt(ISheet sheet, int HeaderRowIndex, bool needHeader)
        {

            DataTable table = new DataTable();
            IRow headerRow;
            int cellCount;
            try
            {
                if (HeaderRowIndex < 0 || !needHeader)
                {
                    headerRow = sheet.GetRow(0) as HSSFRow;
                    cellCount = headerRow.LastCellNum;

                    for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                    {
                        DataColumn column = new DataColumn(Convert.ToString(i));
                        table.Columns.Add(column);
                    }
                }
                else
                {
                    headerRow = sheet.GetRow(HeaderRowIndex);
                    cellCount = headerRow.LastCellNum;

                    for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                    {
                        if (headerRow.GetCell(i) == null)
                        {
                            if (table.Columns.IndexOf(Convert.ToString(i)) > 0)
                            {
                                DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                                table.Columns.Add(column);
                            }
                            else
                            {
                                DataColumn column = new DataColumn(Convert.ToString(i));
                                table.Columns.Add(column);
                            }

                        }
                        else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0)
                        {
                            DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                            table.Columns.Add(column);
                        }
                        else
                        {
                            DataColumn column = new DataColumn(headerRow.GetCell(i).ToString());
                            table.Columns.Add(column);
                        }
                    }
                }
                int rowCount = sheet.LastRowNum;
                for (int i = (HeaderRowIndex + 1); i <= sheet.LastRowNum; i++)
                {
                    try
                    {
                        IRow row;
                        if (sheet.GetRow(i) == null)
                        {
                            row = sheet.CreateRow(i);
                        }
                        else
                        {
                            row = sheet.GetRow(i);
                        }

                        DataRow dataRow = table.NewRow();

                        for (int j = row.FirstCellNum; j <= cellCount; j++)
                        {
                            try
                            {
                                if (row.GetCell(j) != null)
                                {
                                    switch (row.GetCell(j).CellType)
                                    {
                                        case CellType.String:
                                            string str = row.GetCell(j).StringCellValue;
                                            if (str != null && str.Length > 0)
                                            {
                                                dataRow[j] = str.ToString();
                                            }
                                            else
                                            {
                                                dataRow[j] = null;
                                            }
                                            break;
                                        case CellType.Numeric:
                                            if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                            {
                                                dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                            }
                                            else
                                            {
                                                dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                            }
                                            break;
                                        case CellType.Boolean:
                                            dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                            break;
                                        case CellType.Error:
                                            dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                            break;
                                        case CellType.Formula:
                                            switch (row.GetCell(j).CachedFormulaResultType)
                                            {
                                                case CellType.String:
                                                    string strFORMULA = row.GetCell(j).StringCellValue;
                                                    if (strFORMULA != null && strFORMULA.Length > 0)
                                                    {
                                                        dataRow[j] = strFORMULA.ToString();
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = null;
                                                    }
                                                    break;
                                                case CellType.Numeric:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                    break;
                                                case CellType.Boolean:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                    break;
                                                case CellType.Error:
                                                    dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                    break;
                                                default:
                                                    dataRow[j] = "";
                                                    break;
                                            }
                                            break;
                                        default:
                                            dataRow[j] = "";
                                            break;
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                throw exception;
                            }
                        }
                        table.Rows.Add(dataRow);
                    }
                    catch (Exception exception)
                    {
                        throw exception;
                    }
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
            return table;
        }

        #endregion
    }
}
