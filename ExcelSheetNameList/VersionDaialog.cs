// <copyright file="VersionDaialog.cs" company="Tetsuro Ono">
// Copyright (c) 2023 Tetsuro Ono. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>

namespace ExcelSheetNameList
{
    using System;
    using System.Diagnostics;
    using System.Drawing;
    using System.Reflection;
    using System.Windows.Forms;

    /// <summary>
    /// Documentation for VersionDaialog.
    /// </summary>
    public partial class VersionDaialog : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="VersionDaialog"/> class.
        /// </summary>
        public VersionDaialog()
        {
            InitializeComponent();
            InitForm();
        }

        private void InitForm()
        {
            System.Reflection.AssemblyName assemblyName =
                System.Reflection.Assembly.GetExecutingAssembly().GetName();

            string productVersion =
                assemblyName.Version.Major + "." +
                assemblyName.Version.Minor + "." +
                assemblyName.Version.Build;

            FileVersionInfo fileInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
            this.label1.TextAlign = ContentAlignment.TopCenter;
            this.label1.Text = fileInfo.ProductName + Environment.NewLine +
                "Version:" + productVersion + Environment.NewLine +
                 fileInfo.LegalCopyright;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
