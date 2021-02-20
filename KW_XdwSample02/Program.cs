using KW_XdwLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace KW_XdwSample02
{
    class Program
    {
        /// <summary>
        /// ファイルを必ず一つとるようにして、そのファイルまでの
        /// パスを取得する
        /// </summary>
        /// <param name="args">DocuWorks文章へのファイルパス</param>
        static void Main(string[] args)
        {
            try
            {
                using (var binder = new DWDocumentBinder())
                {

                    string dwFilepath = null;       // バインダーへ追加するDW文章のフルパス
                    string dwFolderpath = null;     // バインダーへ追加するDW文章があるフォルダーへのパス

                    if (args.Length > 0)
                    {
                        dwFilepath = Path.GetFullPath(args[0]);
                        dwFolderpath = Path.GetDirectoryName(dwFilepath);
                    }

                    string binderName = null;       // バインダー名

                    if (dwFilepath != null)
                    {
                        binderName = GetBinderName(dwFilepath);
                    }

                    int api_result = binder.Create(dwFolderpath, binderName);

                    if (api_result < 0)
                    {
                        Console.WriteLine(DWError.GetErrorMessage(api_result));
                        return;
                    }

                    binder.Open();

                    binder.SetPageFormAttribute();

                    binder.Add(dwFilepath);

                    binder.Close();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());

            }
        }

        private static string GetBinderName(string dwFilepath)
        {
            // ファイル名が大体 540050_0222_パネル～とかそういう名前になっているので
            // 最初の11文字をバインダー名として追加するようする

            if (dwFilepath != null)
            {
                return dwFilepath.Substring(0, 11);
            }

            return null;
        }
    }
}
