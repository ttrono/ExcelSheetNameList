// <copyright file="Form1.cs" company="Tetsuro Ono">
// Copyright (c) 2023 Tetsuro Ono. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>

namespace ExcelSheetNameList
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Windows.Forms;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Validation;

    /// <summary>
    /// 土台となるフォームクラス.
    /// </summary>
    public partial class Form1 : Form
    {
        private SpreadsheetDocument targetDoc;

        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitDataGrid();
        }

        private bool ValidateExcelFile(string fileName)
        {
            bool validateResult = true;
            string errMsg = null;

            try
            {
                targetDoc = SpreadsheetDocument.Open(fileName, true);
                var validator = new OpenXmlValidator();
                foreach (ValidationErrorInfo error in validator.Validate(targetDoc))
                {
                    // 1件でもエラーが見つかったら終了
                    validateResult = false;
                    break;
                }
            }
            catch (Exception ex)
            {
                // 既に開いているExcelファイルを読み込んだ時
                errMsg = ex.Message;
                validateResult = false;
                string extension = Path.GetExtension(fileName);

                if (extension != ".xlsx" && extension != ".xlsm" && extension != ".xltx" && extension != ".xltm")
                {
                    // テキストファイル等の非Excelファイル読み込み時
                    errMsg = Properties.Resources.ERRMSG003;
                }
                else if (extension == ".xls")
                {
                    // SpreadsheetDocument.Openが旧型式に対応していない
                    errMsg = Properties.Resources.ERRMSG004;
                }

                targetDoc?.Dispose();
            }

            if (!validateResult)
            {
                MessageBox.Show(errMsg, Properties.Resources.ERRMSG000, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return validateResult;
        }

        private Dictionary<string, string> GetSheetList()
        {
            var dict = new Dictionary<string, string>();

            try
            {
                Sheets sheets = targetDoc.WorkbookPart.Workbook.Sheets;
                foreach (Sheet sheet in sheets.Elements<Sheet>())
                {
                    dict.Add(sheet.SheetId, sheet.Name);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Properties.Resources.ERRMSG000, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                targetDoc?.Dispose();
            }

            return dict;
        }

        private void InitDataGrid()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].HeaderText = Properties.Resources.HEADER_COL000;
            dataGridView1.Columns[1].HeaderText = Properties.Resources.HEADER_COL001;
            dataGridView1.Columns[1].Width = 330;
            dataGridView1.RowHeadersWidth = 30; // 左端の選択列
        }

        private void ShowSheetList(Dictionary<string, string> dict)
        {
            InitDataGrid();
            foreach (KeyValuePair<string, string> item in dict)
            {
                dataGridView1.Rows.Add(item.Key, item.Value);
            }

            dataGridView1.ReadOnly = true;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (files.Length > 1)
            {
                MessageBox.Show(Properties.Resources.ERRMSG002, Properties.Resources.ERRMSG000, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var fileName = files[0];
            if (!ValidateExcelFile(fileName))
            {
                return;
            }

            textBox1.Text = fileName;
            var dict = GetSheetList();
            ShowSheetList(dict);
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            // マウスポインターの形状を変更する
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        /// <summary>
        /// リセットボタン押下時の処理.
        /// </summary>
        private void Button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
            dataGridView1.Rows.Clear();
        }

        private void Excelファイル選択OToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                if (!ValidateExcelFile(openFileDialog1.FileName))
                {
                    return;
                }

                textBox1.Text = openFileDialog1.FileName;
                var dict = GetSheetList();
                ShowSheetList(dict);
            }
        }

        private void 終了XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void バージョン情報AToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VersionDaialog dialog = new VersionDaialog();
            dialog.ShowDialog();
        }
    }
}
