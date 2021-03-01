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
    public class DWDocumentBinder : DWDocument
    {
        private const string _extension = ".xbd";

        public string Extention { get; }

        public DWDocumentBinder() { }

        /// <summary>
        /// バインダーを作成する
        /// 同名のバインダーがすでに存在する場合は失敗する。
        /// </summary>
        /// <param name="filepath">バインダーを作成する場所へのパス</param>
        /// <param name="filename">作成するファイル名</param>
        public void Create(string filepath, string filename)
        {
            int api_result = 0;

            string fullpath = filepath + "\\" + filename + _extension;

            // うーん・・・ここで入れるのはどうかなぁ・・
            _filepath = fullpath;

            api_result = Xdwapi.XDW_CreateBinder(fullpath, null);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentBinder::Create関数が失敗しました");
            }

            return;
        }


        /// <summary>
        /// バインダーの先頭へDocuWorksファイルを挿入する
        /// </summary>
        /// <param name="inputPath">挿入するDocuWorksファイルの絶対パス</param>
        public void Add(string inputPath)
        {
            this.Add(1, inputPath);
            return;
        }

        /// <summary>
        /// DocuWorksファイルに付与された見出し・ページ番号を更新する。
        /// </summary>
        public void Update()
        {
            int api_result = Xdwapi.XDW_UpdatePageForm(_handle, Xdwapi.XDW_PAGEFORM_STAY);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog( api_result);
                throw new Exception("DWDocumentBinder::GetDocumentInformation関数が失敗しました");
            }
            return;
        }

        /// <summary>
        /// バインダーへDocuWorks文章を挿入する
        /// </summary>
        /// <param name="position">挿入位置</param>
        /// <param name="inputPath">挿入するDocuWorksファイルの絶対パス</param>
        public void Add(int position, string inputPath)
        {
            int api_result = Xdwapi.XDW_InsertDocumentToBinder(_handle, position, inputPath);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentBinder::Add(position, inputPath)関数が失敗しました");
            }

            return;
        }

        /// <summary>
        /// 複数ファイルをバインダーへ挿入する
        /// </summary>
        /// <param name="inputPaths">ファイルパスのコレクション</param>
        public void Add(string[] inputPaths)
        {
            foreach (string inputPath in inputPaths)
            {
                Xdwapi.XDW_DOCUMENT_INFO info = new Xdwapi.XDW_DOCUMENT_INFO();
                int api_result = Xdwapi.XDW_GetDocumentInformation(_handle, ref info);
                if (api_result < 0)
                {
                    DWErrorLogService.APIErrorLog( api_result);
                    throw new Exception("DWDocumentBinder::Add::GetDocumentInformation関数が失敗しました");
                }

                int position = info.BinderSize + 1;

                Add(position, inputPath);

            }

            return;
        }

        /// <summary>
        /// バインダーに見出し・ページ番号を設定する。
        /// </summary>
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
