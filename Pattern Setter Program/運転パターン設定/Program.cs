using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows;
using System.Diagnostics;

namespace 運転パターン設定
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (Process.GetProcessesByName(
                    System.IO.Path.GetFileNameWithoutExtension(
                        System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1
                ) 
                {
                    MessageBox.Show("すでにプログラムが動作中です。", "多重起動エラー", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ProgramPattern());
        }

    }
}
