using KW_XdwLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KW_XdwSample02
{
    class Program
    {
        /// <summary>
        /// ファイルを必ず一つとるようにして、そのファイルまでの
        /// パスを取得する。
        /// Debug用として、指定フォルダ内に新規バインダーを作り、「スキャン文章」をとりこんで
        /// セーブするような動作にする。
        /// </summary>
        /// <param name="args">DocuWorks文章へのファイルパス</param>
        static void Main(string[] args)
        {
            try
            {
                using (var binder = new DWDocumentBinder())
                {

                    string dwFilePath = null;       // バインダーへ追加するDW文章のフルパス
                    string dwFolderPath = null;     // バインダーへ追加するDW文章があるフォルダーへのパス

                    if (args.Length > 0)
                    {
                        dwFilePath = Path.GetFullPath(args[0]);
                        dwFolderPath = Path.GetDirectoryName(dwFilePath);
                    }
                    else
                    {
                        dwFilePath = @"C:\Users\mizuy\OneDrive\Documents\DocuWorksDev\Test\スキャン文書.xdw";
                        dwFolderPath = Directory.GetCurrentDirectory();
                    }

                    string binderName = null;       // バインダー名

                    if (dwFilePath != null)
                    {
                        binderName = GetBinderName(Path.GetFileNameWithoutExtension(dwFilePath));
                    }

                    // binderNameがnullの時は、DW文章の名前がそのままbinder名となる
                    int api_result = binder.Create(dwFolderPath, binderName);

                    if (api_result < 0)
                    {
                        MessageBox.Show(DWError.GetErrorMessage(api_result, "binder.Create"));
                        
                        return;
                    }

                    api_result =  binder.Open();
                    if (api_result < 0)
                    {
                        MessageBox.Show(DWError.GetErrorMessage(api_result, "binder.Open"));
                        return;
                    }

                    api_result = binder.Add(dwFilePath);
                    if (api_result < 0)
                    {
                        MessageBox.Show(DWError.GetErrorMessage(api_result, "binder.Add"));
                        return;
                    }

                    // 見出し・ページ番号設定
                    api_result = binder.SetPageFormAttribute();
                    if (api_result < 0)
                    {
                        MessageBox.Show(DWError.GetErrorMessage(api_result, "binder.SetAttribute"));
                        return;
                    }

                    // 見出し・ページ番号更新
                    api_result = binder.Update();
                    if (api_result < 0)
                    {
                        MessageBox.Show(DWError.GetErrorMessage(api_result, "binder.Update"));
                        return;
                    }

                    api_result = binder.Save();
                    if (api_result < 0)
                    {
                        MessageBox.Show(DWError.GetErrorMessage(api_result, "binder.Save"));
                        return;
                    }

                    binder.Close();

                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
        }

        private static string GetBinderName(string dwFileName)
        {
            // ファイル名が大体 540050_0222_パネル～とかそういう名前になっているので
            // 最初の11文字をバインダー名として追加するようする

            if (dwFileName != null && dwFileName.Length > 12)
            {
                return dwFileName.Substring(0, 11);
            }

            return dwFileName;
        }
    }
}
