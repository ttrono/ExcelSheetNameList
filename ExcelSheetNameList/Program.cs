// <copyright file="Program.cs" company="Tetsuro Ono">
// Copyright (c) 2023 Tetsuro Ono. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>

namespace ExcelSheetNameList
{
    using System;
    using System.Windows.Forms;

    internal static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
