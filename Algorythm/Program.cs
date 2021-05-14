// Copyright © 2021 Gorokhov Vladislav Dmitrievich. All rights reserved.
// Contacts: <gorohov2017vladislav@yandex.ru>
// License: https://www.gnu.org/licenses/old-licenses/gpl-2.0.html

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Algorythm
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}