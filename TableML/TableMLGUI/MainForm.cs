﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using TableML.Compiler;

namespace TableMLGUI
{
    public partial class MainForm : Form
    {

        /// <summary>
        /// 输出tml文件路径
        /// </summary>
        public string GenTmlPath = "..\\client_setting";
        public string srcFullPath;

        /// <summary>
        /// 生成的代码路径
        /// </summary>
        public string GenCodePath = "..\\client_code\\";
        /// <summary>
        /// lua文件后缀
        /// </summary>
        public string tmlExtensions = ".lua";

        public string NameSpace = "AppSettings";

        public bool GenCSCode
        {
            get { return false; }
        }

        /// <summary>
        /// 框中的文件列表
        /// </summary>
        public string[] fileList
        {
            get
            {
                return tbFileList.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            }
        }
        public MainForm()
        {
            InitializeComponent();
            Init();
            //tbFileList.Text = @"e:\3dsn\plan\005ConfigTable\Src\NPC配置表.xlsx";
        }


        public void Init()
        {
            string useAbsolutePathStr = ConfigurationManager.AppSettings.Get("UseAbsolutePath").Trim().ToLower();
            bool useAbsolutePath = useAbsolutePathStr == "true" || useAbsolutePathStr == "1";

            //源始excel路径
            var srcExcelPath = ConfigurationManager.AppSettings.Get("srcExcelPath");
            var excelSrc = useAbsolutePath ? srcExcelPath : Path.GetFullPath(Application.StartupPath + srcExcelPath);
            this.tbSrcPath.Text = excelSrc;
            srcFullPath = Path.GetFullPath(Application.StartupPath + srcExcelPath);

            //tml文件格式
            var tmlFileEx = ConfigurationManager.AppSettings.Get("TmlExtensions");
            if (!string.IsNullOrEmpty(tmlFileEx))
            {
                tmlExtensions = tmlFileEx;
            }

            //tml路径
            var genTmlPath = ConfigurationManager.AppSettings.Get("GenTmlPath");
            GenTmlPath = useAbsolutePath ? genTmlPath : Path.GetFullPath(Application.StartupPath + genTmlPath);

            //代码路径
            var genCodePath = ConfigurationManager.AppSettings.Get("GenCodePath");
            GenCodePath = useAbsolutePath ? genCodePath : Path.GetFullPath(Application.StartupPath + genCodePath);

            //客户端代码路径
            var dstClientCodePath = ConfigurationManager.AppSettings.Get("dstClientCodePath");
            var dstClientCode = useAbsolutePath ? dstClientCodePath : Path.GetFullPath(Application.StartupPath + dstClientCodePath);
            this.txtCodePath.Text = dstClientCode;

            //客户端tml路径
            var dstClientTmlPath = ConfigurationManager.AppSettings.Get("dstClientTmlPath");
            var dstClientTml = useAbsolutePath ? dstClientTmlPath : Path.GetFullPath(Application.StartupPath + dstClientTmlPath);
            this.txtTmlPath.Text = dstClientTml;

            openFileDialog1.InitialDirectory = srcFullPath;
            openFileDialog1.DefaultExt = "*.xls,*.tsv,*.csv";
            openFileDialog1.Multiselect = true;
            openFileDialog1.SupportMultiDottedExtensions = true;
        }

        #region  文件拖拽到列表
        private void FileField_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void tbFileList_DragEnter(object sender, DragEventArgs e)
        {
            FileField_DragEnter(sender, e);
        }

        private void tbFileList_DragDrop(object sender, DragEventArgs e)
        {
            var files = ((System.Array)e.Data.GetData(DataFormats.FileDrop));
            foreach (var file in files)
            {
                this.tbFileList.Text += "\r\n" + file;
            }
        }

        private void tbFileDir_DragEnter(object sender, DragEventArgs e)
        {
            FileField_DragEnter(sender, e);
        }

        private void tbFileDir_DragDrop(object sender, DragEventArgs e)
        {
            var files = ((System.Array)e.Data.GetData(DataFormats.FileDrop));
            var dragDir = files.GetValue(0).ToString();
            if (Directory.Exists(dragDir) == false)
            {
                Console.WriteLine("Error !{0} 目录不存在或不是目录", dragDir);
                return;
            }
            this.tbSrcPath.Text = dragDir;
        }
        #endregion

        private void btnCompileSelect_Click(object sender, EventArgs e)
        {
            CompileSelect();
        }

        public void CompileSelect(bool msgResult = false)
        {
            CompileSelect(fileList, msgResult);
        }

        /// <summary>
        /// 编译选中的excel
        /// </summary>
        /// <param name="msgResult"></param>
        public void CompileSelect(string[] fullPaths, bool msgResult = false)
        {
            //编译选定的表
            List<string> tmlList = new List<string>();
            var startPath = Environment.CurrentDirectory;
            Console.WriteLine("当前目录：{0}", startPath);
            var compiler = new Compiler();
            Dictionary<string, string> dst2src = new Dictionary<string, string>();
            int comileCount = 0;
            foreach (var filePath in fullPaths)
            {
                if (string.IsNullOrEmpty(filePath))
                {
                    continue;
                }
                Console.WriteLine(filePath);
                List<string> savePath = new List<string>();
                var ext = Path.GetExtension(filePath).Trim().ToLower();
                if (ext == ".xls" || ext == ".xlsx")
                {
                    var sheetNameList = SimpleExcelFile.getOutFileNameList(filePath);
                    for(int idx = 0; idx < sheetNameList.Count; idx ++)
                    {
                        savePath.Add(GenTmlPath + "\\" + sheetNameList[idx] + tmlExtensions);

                    }
                }
                else
                {
                    continue;
                }
                //编译表时，生成代码
                TableCompileResult compileResult = compiler.Compile(filePath, savePath);
                tmlList.Add(Path.GetFullPath(savePath));
                var dstFileName = Path.GetFileNameWithoutExtension(savePath);
                if (dst2src.ContainsKey(dstFileName) == false)
                {
                    dst2src.Add(dstFileName, Path.GetFileName(filePath));
                }
                Console.WriteLine("编译结果:{0}---->{1}", filePath, savePath);
                Console.WriteLine();

                //生成代码
                //BatchCompiler batchCompiler = new BatchCompiler();
                ////NOTE 替换成相对路径(保证最后只有文件名)
                //string repStr = Directory.GetParent(compileResult.TabFileRelativePath).FullName + "\\";
                //compileResult.TabFileRelativePath = compileResult.TabFileRelativePath.Replace(repStr, "");
                //if (GenCSCode)
                //{
                //    batchCompiler.GenCodeFile(compileResult, DefaultTemplate.GenSingleClassCodeTemplate, GenCodePath, NameSpace, tmlExtensions, null, true);
                //}

                if (compileResult != null)
                {
                    comileCount += 1;
                }
            }
            BatchCompiler.SaveCompileResult(dst2src);

            if (msgResult) { ShowCompileResult(comileCount); }
        }

        /// <summary>
        /// 编译所有的excel
        /// </summary>
        /// <param name="msgResult">是否弹出编译结果</param>
        public void CompileAllExcel(bool msgResult = false)
        {
            //编译整个目录
            var startPath = Environment.CurrentDirectory;
            Console.WriteLine("当前目录：{0}", startPath);

            var srcDirectory = tbSrcPath.Text;

            var batchCompiler = new BatchCompiler();

            string templateString = DefaultTemplate.GenSingleClassCodeTemplate;

            var results = batchCompiler.CompileAll(srcDirectory, GenTmlPath, GenCodePath,
               templateString, NameSpace, tmlExtensions, null, !string.IsNullOrEmpty(GenCodePath), GenCSCode);
            if (msgResult) { ShowCompileResult(results.Count); }
        }


        public void ShowCompileResult(int count)
        {
            MessageBox.Show(string.Format("共编译{0}张表", count), "编译完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnUpdateCSSyntax_Click(object sender, EventArgs e)
        {
            ExcelHelper.UpdateAllTableSyntax();
        }

        private void btnSyncCode_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(txtCodePath.Text) == false)
            {
                Directory.CreateDirectory(txtCodePath.Text);
            }
            FileHelper.CopyFolder(GenCodePath, txtCodePath.Text);
            Console.WriteLine("copy {0} to \r\n {1}", GenCodePath, txtCodePath.Text);
            MessageBox.Show(string.Format("{0}\r\n同步完成", txtCodePath.Text), "同步完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSyncTml_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(txtTmlPath.Text) == false)
            {
                Directory.CreateDirectory(txtTmlPath.Text);
            }
            FileHelper.CopyFolder(GenTmlPath, txtTmlPath.Text);
            Console.WriteLine("copy {0} to \r\n {1}", GenTmlPath, txtTmlPath.Text);
            MessageBox.Show(string.Format("{0}\r\n同步完成", txtTmlPath.Text), "同步完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnCheckNameRepet_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(tbFileList.Text))
            {
                ExcelHelper.CheckNameRepet(fileList);
            }
        }

        private void btnOpenCodeDir_Click(object sender, EventArgs e)
        {
            FileHelper.OpenFolder(GenCodePath);
        }

        private void btnOpenTmlDir_Click(object sender, EventArgs e)
        {
            FileHelper.OpenFolder(GenTmlPath);
        }

        private void btnCheckNameEmpty_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(tbFileList.Text))
            {
                ExcelHelper.CheckNameEmpty(fileList);
            }
        }

        private void btnCheckCSKW_Click(object sender, EventArgs e)
        {

            if (!string.IsNullOrEmpty(tbFileList.Text))
            {
                ExcelHelper.CheckHasKeyWords(fileList);
            }
        }

        private void btnUpdateSelectCSSyntax_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(tbFileList.Text))
            {
                foreach (string filePath in fileList)
                {
                    ExcelHelper.ToCSharpSyntax(filePath);
                }
                Console.WriteLine("数据类型更新完成");
            }
        }

        private void btnCompileExcel_Click(object sender, EventArgs e)
        {
            CompileAllExcel();
        }

        /// <summary>
        /// 命令行模式编译并插入到sqlite中
        /// </summary>
        public void CMDCompile()
        {
            CompileAllExcel();
        }


        private void btnHelp_Click(object sender, EventArgs e)
        {
            HelpForm helpForm = new HelpForm();
            helpForm.Show();
        }

        private void btnClearConsole_Click(object sender, EventArgs e)
        {
            Console.Clear();
        }

        private void btnFileBrowser_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                var selectFiles = openFileDialog1.FileNames;
                foreach (string file in selectFiles)
                {
                    tbFileList.Text += file + "\r\n";
                }
            }
        }

    }
}