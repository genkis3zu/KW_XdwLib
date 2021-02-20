using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FujiXerox.DocuWorks.Toolkit;

namespace KW_XdwLib
{
    public class DWDocumentFile : DWDocument, IDisposable
    {
        public DWDocumentFile() : base() { }

        /// <summary>
        /// DocuWorks文章を開く
        /// </summary>
        /// <param name="filepath">ファイル名を含む文章へのフルパス</param>
        /// <returns>1:正常終了, 0以下:各エラーコード対応</returns>
        public int Open(string filepath)
        {
            int api_result = Xdwapi.XDW_OpenDocumentHandle(filepath, ref _handle, _mode);

            if (api_result < 0)
            {
                return api_result;
            }

            return 1;
        }

        /// <summary>
        /// DocuWorks文章を閉じる
        /// </summary>
        public void Close()
        {
            if (_handle != null)
            {
                Xdwapi.XDW_CloseDocumentHandle(_handle);
            }
        }

    }
}
