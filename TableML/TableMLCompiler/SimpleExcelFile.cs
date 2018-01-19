using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TableML.Compiler
{
    // enum DATA_TYPE
    // {
    //     N = 0x1, //NULL 可以为空,产品可以不填
    //     D = 0x2, //dict,字典类型
    //     L = 0x4, //list,列表类型
    //     T = 0x8, //tuple,元组类型
    //     I = 0x10, //integer,整形
    //     F = 0x20, //float,浮点类型
    //     S = 0x40, //string,字符串类型
    //     A = D | L | T | I | F | S //any,任意类型,可能字符串,可能整形,可能浮点,可以是list,tuple,dict
    // }

    


    /// <summary>
    /// 对NPOI Excel封装, 支持xls, xlsx 和 tsv
    /// 带有头部、声明、注释
    /// </summary>
    public class SimpleExcelFile : ITableSourceFile
    {
        public Dictionary<string, int> ColName2Index { get; set; }
        public Dictionary<string, int> ColType2Index { get; set; }
        public Dictionary<int, string> Index2ColType { get; set; }
        public Dictionary<string, string> ColType2Statement { get; set; }
        public Dictionary<string, string> Fields2ColType { get; set; } //  string,or something
        public Dictionary<string, string> ColName2Comment { get; set; } // string comment
        public string ExcelFileName { get; set; }
        //NOTE by Nil 根据特殊的Excel格式定制,暂时修改成最普通的三行。
        /// <summary>
        /// Header, Statement, Comment, at lease 3 rows
        /// 预留行数
        /// </summary>
        public const int PreserverRowCount = 3;
        /// <summary>
        /// 从指定列开始读,默认是0,0行忽略不读
        /// </summary>
        public const int StartColumnIdx = 1;

        /// excel 是从 0，0 开始的

        /// <summary>
        /// 默认数据类型行数
        /// </summary>
        public const int DefDataTypeIdx = 1;

        /// <summary>
        /// 默认字段名行数
        /// </summary>
        public const int DefDataNameIdx = 2;

        /// <summary>
        /// 默认唯一ID位置
        /// </summary>
        public const int DefaultKeyIdx = 1;

        /// <summary>
        /// 默认唯一ID字段名
        /// </summary>
        public const string DefaultKeyName = "__KEY__";

        private string Path;
        private IWorkbook Workbook;
        private ISheet Worksheet;
        public bool IsLoadSuccess = true;
        private int _columnCount;
        private int sheetCount;
        public int SheetCount { get { return sheetCount; } private set { sheetCount = value; } }

        public SimpleExcelFile(string excelPath)
        {
            Path = excelPath;
            ColName2Index = new Dictionary<string, int>();
            ColType2Index = new Dictionary<string, int>();
            Index2ColType = new Dictionary<int, string>();
            Fields2ColType = new Dictionary<string, string>();
            ColType2Statement = new Dictionary<string, string>();
            ColName2Comment = new Dictionary<string, string>();
            ExcelFileName = System.IO.Path.GetFileName(excelPath);
            ParseExcel(excelPath);
        }

        /// <summary>
        /// Parse Excel file to data grid
        /// </summary>
        /// <param name="filePath"></param>
        private void ParseExcel(string filePath)
        {
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) // no isolation
            {
                try
                {
                    Workbook = WorkbookFactory.Create(file);
                }
                catch (Exception e)
                {
                    //                    throw new Exception(string.Format("无法打开Excel: {0}, 可能原因：正在打开？或是Office2007格式（尝试另存为）？ {1}", filePath, e.Message));
                    ConsoleHelper.Error(string.Format("无法打开Excel: {0}, 可能原因：正在打开？或是Office2007格式（尝试另存为）？ {1}", filePath, e.Message));
                    IsLoadSuccess = false;
                    return;
                }
            }

            if (Workbook == null)
            {
                //                    throw new Exception(filePath + " Null Workbook");
                ConsoleHelper.Error(filePath + " Null Workbook");
                return;
            }
            SheetCount = Workbook.NumberOfSheets;
            //for (int idx = 0; idx < sheetCount; idx++)
            {
                ParseSheet(filePath, 0);
            }

        }

        /// <summary>
        /// 检查excel是否符合输出规范
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool CheckRule(string filePath, int sheetIdx)
        {
            if (Workbook == null)
            {
                //                    throw new Exception(filePath + " Null Workbook");
                ConsoleHelper.Error(filePath + " Null Workbook");
                return false;
            }
            Worksheet = Workbook.GetSheetAt(sheetIdx);
            if (Worksheet == null)
            {
                //                    throw new Exception(filePath + " Null Worksheet");
                ConsoleHelper.Error(filePath + " Null Worksheet");
                return false;
            }

            var sheetRowCount = GetWorksheetCount();
            if (sheetRowCount < PreserverRowCount)
            {
                //                    throw new Exception(string.Format("{0} At lease {1} rows of this excel", filePath, sheetRowCount));
                ConsoleHelper.Error(string.Format("{0} At lease {1} rows of this excel", filePath, sheetRowCount));
                return false;

            }
            var row = Worksheet.GetRow(1);
            if (row == null || row.Cells.Count < 2)
            {
                //                throw new Exception(filePath + "第二行至少需要3列");
                ConsoleHelper.Error(filePath + "第二行至少需要3列");
                return false;
            }
            return true;
        }

        /// <summary>
        /// 解析excel的sheet
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetIdx"></param>
        public void ParseSheet(string filePath, int sheetIdx)
        {
            if (CheckRule(filePath, sheetIdx) == false) { return; }
            if (Worksheet == null) return;
            /**表头结构如下所示：
            *   编号 名称 CD时间
            *   int string int
            *   Id  Name    CDTime
            */

            // 第0行是描述，忽略
            //NOTE 从第1行开始读
            // (数据类型) 
            var headerRow = Worksheet.GetRow(DefDataTypeIdx);

            _columnCount = headerRow.LastCellNum;
            // 列总数保存
            int columnCount = GetColumnCount();

            //NOTE by Nil 列数据类型：从指定的列开始读取
            int emptyColumn = 0;
            for (int columnIndex = StartColumnIdx; columnIndex <= columnCount; columnIndex++)
            {
                var cell = headerRow.GetCell(columnIndex);
                var dataType = cell != null ? cell.ToString().Trim() : ""; // trim!
                var realIdx = columnIndex - StartColumnIdx;
                if (string.IsNullOrEmpty(dataType))
                {
                    //NOTE 如果列名是空，当作注释处理
                    emptyColumn += 1;
                    dataType = string.Concat("#Comment#", emptyColumn);
                }
                ColType2Index[dataType] = realIdx;
                Index2ColType[realIdx] = dataType;
            }
            // 表头声明(字段名称)
            var statementRow = Worksheet.GetRow(DefDataNameIdx);
            for (int columnIndex = StartColumnIdx; columnIndex <= columnCount; columnIndex++)
            {
                var realIdx = columnIndex - StartColumnIdx;
                if (Index2ColType.ContainsKey(realIdx) == false)
                {
                    continue;
                }
                var colType = Index2ColType[realIdx];
                var statementCell = statementRow.GetCell(columnIndex);
                var statementString = statementCell != null ? statementCell.ToString() : "";
                if (StartColumnIdx == DefaultKeyIdx)
                    statementString = DefaultKeyName;
                // 字段名对应类型
                Fields2ColType[statementString] = colType;
            }
        }

        public string CombieLine(string commentString, string lineStr)
        {
            if (commentString.Contains(lineStr))
            {
                var comments = commentString.Split(new string[] { lineStr }, StringSplitOptions.None);
                StringBuilder sb = new StringBuilder();
                sb.Append(comments[0]);
                for (int idx = 1; idx < comments.Length; idx++)
                {
                    if (string.IsNullOrEmpty(comments[idx])) continue;
                    sb.Append(string.Concat("\r\n", "       ///  ", comments[idx]));
                }
                commentString = sb.ToString();
            }
            return commentString;
        }

        /// <summary>
        /// 统一接口：获取单元格内容
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string GetCellString(ICell cell)
        {
            if (cell == null) return "";
            string result = string.Empty;
            switch (cell.CellType)
            {
                case CellType.Unknown:
                    result = cell.StringCellValue;
                    break;
                case CellType.Numeric:
                    result = cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                    break;
                case CellType.String:
                    result = cell.StringCellValue;
                    break;
                case CellType.Formula:
                    //NOTE 单元格为公式，分类型
                    switch (cell.CachedFormulaResultType)
                    {
                        //已测试的公式:SUM,& 
                        case CellType.Numeric:
                            result = cell.NumericCellValue.ToString();
                            break;
                        case CellType.String:
                            result = cell.StringCellValue;
                            break;
                    }
                    break;
                case CellType.Blank:
                    result = "";
                    break;
                case CellType.Boolean:
                    result = cell.BooleanCellValue ? "1" : "0";
                    break;
                case CellType.Error:
                    result = cell.ErrorCellValue.ToString();
                    break;
                default:
                    result = "未知类型";
                    break;
            }
            return result;
        }

        /// <summary>
        /// 是否存在列名
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public bool HasColumn(string columnName)
        {
            return ColName2Index.ContainsKey(columnName);
        }

        /// <summary>
        /// 清除行内容
        /// </summary>
        /// <param name="row"></param>
        public void ClearRow(int row)
        {
            if (Worksheet != null)
            {
                var theRow = Worksheet.GetRow(row);
                Worksheet.RemoveRow(theRow);
            }
        }

        public float GetFloat(string columnName, int row)
        {
            return float.Parse(GetString(columnName, row));
        }

        public int GetInt(string columnName, int row)
        {
            return int.Parse(GetString(columnName, row));
        }

        /// <summary>
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="dataRow">无计算表头的数据行数</param>
        /// <returns></returns>
        public string GetString(string columnName, int dataRow)
        {
            if (Worksheet == null) return null;
            dataRow += PreserverRowCount;

            var theRow = Worksheet.GetRow(dataRow);
            if (theRow == null)
                theRow = Worksheet.CreateRow(dataRow);

            var colIndex = ColName2Index[columnName] + SimpleExcelFile.StartColumnIdx;
            var cell = theRow.GetCell(colIndex);
            if (cell == null)
                cell = theRow.CreateCell(colIndex);

            return GetCellString(cell);
        }

        /// <summary>
        /// 不带预留头的数据总行数
        /// </summary>
        /// <returns></returns>
        public int GetRowsCount()
        {
            return GetWorksheetCount() - PreserverRowCount;
        }

        /// <summary>
        /// 工作表的总行数
        /// </summary>
        /// <returns></returns>
        private int GetWorksheetCount()
        {
            return Worksheet == null ? 0 : Worksheet.LastRowNum + 1;
        }

        private ICellStyle GreyCellStyleCache;

        public void SetRowGrey(int row)
        {
            if (Worksheet == null) { return; }
            var theRow = Worksheet.GetRow(row);
            foreach (var cell in theRow.Cells)
            {
                if (GreyCellStyleCache == null)
                {
                    var newStyle = Workbook.CreateCellStyle();
                    newStyle.CloneStyleFrom(cell.CellStyle);
                    //newStyle.FillBackgroundColor = colorIndex;
                    newStyle.FillPattern = FillPattern.Diamonds;
                    GreyCellStyleCache = newStyle;
                }

                cell.CellStyle = GreyCellStyleCache;
            }
        }

        public void SetRow(string columnName, int row, string value)
        {
            if (!ColName2Index.ContainsKey(columnName))
            {
                //                throw new Exception(string.Format("No Column: {0} of File: {1}", columnName, Path));
                ConsoleHelper.Error(string.Format("No Column: {0} of File: {1}", columnName, Path));
                return;
            }
            if (Worksheet == null) return;
            var theRow = Worksheet.GetRow(row);
            if (theRow == null)
                theRow = Worksheet.CreateRow(row);
            var cell = theRow.GetCell(ColName2Index[columnName]);
            if (cell == null)
                cell = theRow.CreateCell(ColName2Index[columnName]);

            if (value.Length > (1 << 14)) // if too long
            {
                value = value.Substring(0, 1 << 14);
            }
            cell.SetCellValue(value);
        }

        public void Save(string toPath)
        {
            /*for (var loopRow = Worksheet.FirstRowNum; loopRow <= Worksheet.LastRowNum; loopRow++)
        {
            var row = Worksheet.GetRow(loopRow);
            bool emptyRow = true;
            foreach (var cell in row.Cells)
            {
                if (!string.IsNullOrEmpty(cell.ToString()))
                    emptyRow = false;
            }
            if (emptyRow)
                Worksheet.RemoveRow(row);
        }*/
            //try
            {
                using (var memStream = new MemoryStream())
                {
                    Workbook.Write(memStream);
                    memStream.Flush();
                    memStream.Position = 0;

                    using (var fileStream = new FileStream(toPath, FileMode.Create, FileAccess.Write))
                    {
                        var data = memStream.ToArray();
                        fileStream.Write(data, 0, data.Length);
                        fileStream.Flush();
                    }
                }
            }
            //catch (Exception e)
            //{
            //    CDebug.LogError(e.Message);
            //    CDebug.LogError("是否打开了Excel表？");
            //}
        }

        public void Save()
        {
            Save(Path);
        }

        /// <summary>
        /// 获取列总数
        /// </summary>
        /// <returns></returns>
        public int GetColumnCount()
        {
            return _columnCount - StartColumnIdx;
        }



        /// <summary>
        /// 读表中的字段获取输出文件名
        /// 做好约定输出tml文件名在指定的单元格，不用遍历整表让解析更快
        /// </summary>
        /// <returns></returns>
        public static string GetSheetName(ISheet worksheet, string filePath)
        {
            var row = worksheet.GetRow(1);
            if (row == null || row.Cells.Count < 2)
            {
                //                throw new Exception(filePath + "第二行至少需要3列");
                ConsoleHelper.Error(filePath + "：表：" + worksheet.SheetName + "第二行至少需要3列");
                return "";
            }

            var idxName = worksheet.SheetName.IndexOf("_");
            string outFileName = "";

            if (idxName != -1)
            {
                outFileName = worksheet.SheetName.Substring(idxName + 1);
            }
            else
            {
                ConsoleHelper.Error(filePath + "没有找到匹配的配置表名称，excel表子表是否添加'_'？");
            }
            return outFileName;
        }

        
        /// <summary>
        /// 获得Excel 文件中 sheet 文件名列表
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<string> getOutFileNameList(string filePath)
        {
            var listName = new List<string>();

            IWorkbook workbook;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) // no isolation
            {
                try
                {
                    workbook = WorkbookFactory.Create(file);
                }
                catch (Exception e)
                {
                    //                    throw new Exception(string.Format("无法打开Excel: {0}, 可能原因：正在打开？或是Office2007格式（尝试另存为）？ {1}", filePath, e.Message));
                    ConsoleHelper.Error(string.Format("无法打开Excel: {0}, 可能原因：正在打开？或是Office2007格式（尝试另存为）？ {1}", filePath, e.Message));
                    return listName;
                }
            }

            var sheetCount = workbook.NumberOfSheets;

            if (sheetCount <= 0)
            {
                ConsoleHelper.Error(filePath + "Null Worksheet");
                return listName;
            }

            for (int idx = 0; idx < sheetCount; idx ++)
            {
                var worksheet = workbook.GetSheetAt(idx);
                var name = GetSheetName(worksheet, filePath);
                if (string.IsNullOrEmpty(name))
                {
                    listName.Add(name);
                }
            }

            return listName;
        }
    }
}