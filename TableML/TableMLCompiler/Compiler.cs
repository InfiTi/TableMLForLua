#region Copyright (c) 2015 KEngine / Kelly <http://github.com/mr-kelly>, All rights reserved.

// KEngine - Toolset and framework for Unity3D
// ===================================
// 
// Filename: SettingModuleEditor.cs
// Date:     2015/12/03
// Author:  Kelly
// Email: 23110388@qq.com
// Github: https://github.com/mr-kelly/KEngine
// 
// This library is free software; you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public
// License as published by the Free Software Foundation; either
// version 3.0 of the License, or (at your option) any later version.
// 
// This library is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
// Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public
// License along with this library.

#endregion
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TableML.Compiler
{

    /// <summary>
    /// Compile Excel to TSV
    /// </summary>
    public class Compiler
    {

        /// <summary>
        /// 编译时，判断格子的类型
        /// </summary>
        public enum CellType
        {
            Value,
            Comment,
            If,
            Endif
        }

        private readonly CompilerConfig _config;

        public Compiler()
            : this(new CompilerConfig()
            {
            })
        {
        }

        public Compiler(CompilerConfig cfg)
        {
            _config = cfg;
        }

        /// <summary>
        /// 生成tml文件内容
        /// </summary>
        /// <param name="path">源excel的相对路径</param>
        /// <param name="excelFile"></param>
        /// <param name="compileToFilePath">编译后的tml文件路径</param>
        /// <param name="compileBaseDir"></param>
        /// <param name="doCompile"></param>
        /// <returns></returns>
        private TableCompileResult DoCompilerExcelReader(string path, ITableSourceFile excelFile, string compileToFilePath = null, string compileBaseDir = null, bool doCompile = true)
        {
            var renderVars = new TableCompileResult();
            renderVars.ExcelFile = excelFile;
            renderVars.FieldsInternal = new List<TableColumnVars>();

            var tableBuilder = new StringBuilder();
            var rowBuilder = new StringBuilder();
            var ignoreColumns = new HashSet<int>();

            tableBuilder.Append("local \n");



            // Header Column
            foreach (var colNameStr in excelFile.ColName2Index.Keys)
            {
                if (string.IsNullOrEmpty(colNameStr))
                {
                    continue;
                }


            }
            tableBuilder.Append("\n");
            //以上是tml写入的第一行


            var fileName = Path.GetFileNameWithoutExtension(path);
            string exportPath;
            if (!string.IsNullOrEmpty(compileToFilePath))
            {
                exportPath = compileToFilePath;
            }
            else
            {
                // use default
                exportPath = fileName + _config.ExportTabExt;
            }
            // 是否写入文件
            if (doCompile)
                File.WriteAllText(exportPath, tableBuilder.ToString());

            return renderVars;
        }

        /// <summary>
        /// 获取#if A B语法的变量名，返回如A B数组
        /// </summary>
        /// <param name="cellStr"></param>
        /// <returns></returns>
        private string[] GetIfVars(string cellStr)
        {
            return cellStr.Replace("#if", "").Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }

        /// <summary>
        /// 检查一个表头名，是否是可忽略的注释
        /// 或检查一个字符串
        /// </summary>
        /// <param name="colNameStr"></param>
        /// <returns></returns>
        private CellType CheckCellType(string colNameStr)
        {
            if (colNameStr.StartsWith("#if"))
                return CellType.If;
            if (colNameStr.StartsWith("#endif"))
                return CellType.Endif;
            foreach (var commentStartsWith in _config.CommentStartsWith)
            {
                if (colNameStr.ToLower().Trim().StartsWith(commentStartsWith.ToLower()))
                {
                    return CellType.Comment;
                }
            }

            return CellType.Value;
        }

        /// <summary>
        /// Compile the specified path, auto change extension to config `ExportTabExt`
        /// </summary>
        /// <param name="path">Path.</param>
        public TableCompileResult Compile(string path)
        {
            var outputPath = System.IO.Path.ChangeExtension(path, this._config.ExportTabExt);
            return Compile(path, outputPath);
        }



        /// <summary>
        /// Compile a setting file, return a hash for template
        /// </summary>
        /// <param name="path"></param>
        /// <param name="compileToFilePath"></param>
        /// <param name="compileBaseDir"></param>
        /// <param name="doRealCompile">Real do, or just get the template var?</param>
        /// <returns></returns>
        public TableCompileResult Compile(string path, string compileToFilePath, string compileBaseDir = null, bool doRealCompile = true)
        {
            // 确保目录存在
            compileToFilePath = Path.GetFullPath(compileToFilePath);
            var compileToFileDirPath = Path.GetDirectoryName(compileToFilePath);

            if (!Directory.Exists(compileToFileDirPath))
                Directory.CreateDirectory(compileToFileDirPath);

            var ext = Path.GetExtension(path);

            ITableSourceFile sourceFile = new SimpleExcelFile(path);

            var hash = DoCompilerExcelReader(path, sourceFile, compileToFilePath, compileBaseDir, doRealCompile);
            return hash;

        }

        public List<TableCompileResult> Compile(string compileToFilePath, List<string> paths, string compileBaseDir = null, bool doRealCompile = true)
        {
            var lts = new List<TableCompileResult>();

            foreach (var path in paths)
            {
                lts.Add(Compile(path, compileToFilePath, compileBaseDir, doRealCompile));
            }

            return lts;
        }
    }
}