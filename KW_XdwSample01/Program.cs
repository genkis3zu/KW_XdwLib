using KW_XdwLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KW_XdwSample01
{
    /// <summary>
    /// 指定のDWファイルへアノテーションを追加する。
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string filepath = null;

                if (args.Length > 0)
                {
                    filepath = args[0];
                }
                else
                {
                    filepath = @"C:\Users\mizuy\OneDrive\Documents\DocuWorksDev\Test\ヨット設計図.xdw";
                }

                using (var file = new DWDocumentFile(filepath))
                {
                    int api_result = file.Open();
                    if (api_result < 0)
                    {
                        DWError.GetErrorMessage(api_result, "file.Open");
                        return;
                    }

                    if (!file.IsDWDocument()) return;

                    api_result = file.AddStampAnnotation();

                    if (api_result < 0)
                    {
                        DWError.GetErrorMessage(api_result, "file.AddStampAnnotation");
                        return;
                    }

                    api_result = file.Save();

                    if (api_result < 0)
                    {
                        DWError.GetErrorMessage(api_result, "file.Save");
                        return;
                    }
                    api_result = file.Close();

                    if (api_result < 0)
                    {
                        DWError.GetErrorMessage(api_result, "file.Close");
                        return;
                    }
                }

                return;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }
    }
}
