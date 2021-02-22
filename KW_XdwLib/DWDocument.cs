using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FujiXerox.DocuWorks.Toolkit;

namespace KW_XdwLib
{
    public class DWDocument : IDisposable
    {
        protected Xdwapi.XDW_DOCUMENT_HANDLE _handle = null;
        protected Xdwapi.XDW_OPEN_MODE_EX _mode = null;
        protected string _filepath = null;

        private bool disposedValue;

        public DWDocument()
        {
            _handle = new Xdwapi.XDW_DOCUMENT_HANDLE();
            _mode = new Xdwapi.XDW_OPEN_MODE_EX();

            // とりあえず、現状はバインダーは新規作成されるものとして考える
            _mode.Option = Xdwapi.XDW_OPEN_UPDATE;
            _mode.AuthMode = Xdwapi.XDW_AUTH_NODIALOGUE;
        }

        /// <summary>
        /// バインダーを開く
        /// </summary>
        /// <returns>1:正常終了, 0以下：各エラーコード参照</returns>
        public int Open()
        {
            int api_result = Xdwapi.XDW_OpenDocumentHandle(_filepath, ref _handle, _mode);

            if (api_result < 0)
            {
                return api_result;
            }

            return 1;
        }

        /// <summary>
        /// バインダーを閉じる
        /// </summary>
        /// <returns>1:正常終了, 0以下：各エラーコード参照</returns>
        public int Close()
        {
            int api_result = Xdwapi.XDW_CloseDocumentHandle(_handle);

            if (api_result < 0)
            {
                return api_result;
            }

            return 1;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: マネージド状態を破棄します (マネージド オブジェクト)
                    if (_handle != null)
                    {
                        Xdwapi.XDW_CloseDocumentHandle(_handle);
                    }

                }

                // TODO: アンマネージド リソース (アンマネージド オブジェクト) を解放し、ファイナライザーをオーバーライドします
                // TODO: 大きなフィールドを null に設定します
                disposedValue = true;
            }
        }

        // // TODO: 'Dispose(bool disposing)' にアンマネージド リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします
        // ~DWDocumentFile()
        // {
        //     // このコードを変更しないでください。クリーンアップ コードを 'Dispose(bool disposing)' メソッドに記述します
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // このコードを変更しないでください。クリーンアップ コードを 'Dispose(bool disposing)' メソッドに記述します
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
