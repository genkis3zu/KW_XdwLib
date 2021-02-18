using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FujiXerox.DocuWorks.Toolkit;

namespace KW_XdwLib
{
    /// <summary>
    /// DocuWorksバインダーの作成
    /// 
    /// 各メソッドは、int整数型を返す。
    /// 1:正常終了
    /// 0以下：各エラーコード参照
    /// </summary>
    public class Binder
    {
        private Xdwapi.XDW_DOCUMENT_HANDLE _handle;
        private Xdwapi.XDW_OPEN_MODE_EX _mode;
        private const string _extension = ".xbd";

        public Binder() {
            _handle = new Xdwapi.XDW_DOCUMENT_HANDLE();
            _mode = new Xdwapi.XDW_OPEN_MODE_EX();

            // とりあえず、現状はバインダーは新規作成されるものとして考える
            _mode.Option = Xdwapi.XDW_OPEN_UPDATE;
            _mode.AuthMode = Xdwapi.XDW_AUTH_NODIALOGUE;
        }

        /// <summary>
        /// ハンドルの解放
        /// </summary>
        ~Binder()
        {
            if (_handle != null)
            {
                Xdwapi.XDW_CloseDocumentHandle(_handle);
            }
        }

        /// <summary>
        /// バインダーを作成する
        /// </summary>
        /// <param name="filepath">バインダーを作成する場所へのパス</param>
        /// <param name="filename">作成するファイル名</param>
        /// <returns>1:正常終了, 0以下：各エラーコード参照</returns>
        public int Create(string filepath, string filename)
        {
            int api_result;

            if ((_handle != null) && (_mode != null))
            {
                api_result = Xdwapi.XDW_OpenDocumentHandle(filepath, ref _handle, _mode);
                if (api_result < 0)
                {
                    // オープンエラー
                    return api_result;
                }
            }
            else
            {
                throw new Exception("ハンドルか、またはモードの設定がされていません");
            }

            string fullpath = filepath + "\\" + filename + _extension;

            api_result = Xdwapi.XDW_CreateBinder(fullpath, null);
            if (api_result < 0)
            {
                // 作成エラー
                return api_result;
            }

            return 1;
        }

        /// <summary>
        /// バインダーの先頭へDocuWorksファイルを挿入する
        /// </summary>
        /// <param name="inputPath">挿入するDocuWorksファイルの絶対パス</param>
        /// <returns>1:正常終了, 0以下:各エラーコード対応</returns>
        public int Add(string inputPath)
        {
            int api_result = this.Add(1, inputPath);
            if (api_result < 0)
            {
                return api_result;
            }
            return 1;
        }

        /// <summary>
        /// 複数ファイルをバインダーへ挿入する
        /// </summary>
        /// <param name="inputPaths">ファイルパスのコレクション</param>
        /// <returns>1:正常終了, 0以下:各エラーコード対応</returns>
        public int Add(List<string> inputPaths)
        {
            foreach (string inputPath in inputPaths)
            {
                Xdwapi.XDW_DOCUMENT_INFO info = new Xdwapi.XDW_DOCUMENT_INFO();
                int api_result = Xdwapi.XDW_GetDocumentInformation(_handle, ref info);
                if (api_result < 0)
                {
                    return api_result;
                }

                int position = info.BinderSize + 1;

                api_result = Add(position, inputPath);

                if (api_result < 0)
                {
                    return api_result;
                }
            }

            return 1;
        }

        /// <summary>
        /// バインダーへDocuWorks文章を挿入する
        /// </summary>
        /// <param name="position">挿入位置</param>
        /// <param name="inputPath">挿入するDocuWorksファイルの絶対パス</param>
        /// <returns>1:正常終了, 0以下:各エラーコード対応</returns>
        public int Add(int position, string inputPath)
        {
            int api_result = Xdwapi.XDW_InsertDocumentToBinder(_handle, position, inputPath);

            if (api_result < 0)
            {
                return api_result;
            }

            return 1;
        }

        /// <summary>
        /// バインダーに見出し・ページ番号を設定する。
        /// </summary>
        /// <returns></returns>
        public int SetPageFormAttribute()
        {
            Xdwapi.XDW_DOCUMENT_INFO info = new Xdwapi.XDW_DOCUMENT_INFO();
            int api_result = Xdwapi.XDW_GetDocumentInformation(_handle, ref info);
            if (api_result < 0)
            {
                return api_result;
            }

            string outString = null;
            int alinment = 0;
            if (info.DocType == Xdwapi.XDW_DT_DOCUMENT)
            {
                outString = "DocuWorks文章のページ番号 # です。";
                // DocuWorks文章の場合、ページ番号を左下に設定
                alinment = Xdwapi.XDW_ALIGN_LEFT;
            }
            else
            {
                outString = "DocuWorksバインダーのページ番号 # です。";
                // バインダーの場合、ページ番号を右下に設定
                alinment = Xdwapi.XDW_ALIGN_RIGHT;
            }

            // ページ番号の書式を設定する
            api_result = Xdwapi.XDW_SetPageFormAttribute(_handle, Xdwapi.XDW_PAGEFORM_PAGENUMBER, Xdwapi.XDW_ATN_Text, Xdwapi.XDW_ATYPE_STRING, outString);
            if (api_result < 0)
            {
                return api_result;
            }

            // ページ番号の位置揃えを設定する
            api_result = Xdwapi.XDW_SetPageFormAttribute(_handle, Xdwapi.XDW_PAGEFORM_PAGENUMBER, Xdwapi.XDW_ATN_Alignment, Xdwapi.XDW_ATYPE_INT, alinment);
            if (api_result < 0)
            {
                return api_result;
            }

            return 1;
        }
    }
}
