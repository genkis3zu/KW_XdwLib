using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FujiXerox.DocuWorks.Toolkit;

namespace KW_XdwLib
{
    public abstract class DWDocument : IDisposable
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
        /// DocuWorksファイルを開く
        /// </summary>
        public void Open()
        {
            int api_result = Xdwapi.XDW_OpenDocumentHandle(_filepath, ref _handle, _mode);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(NLog.LogLevel.Error, api_result);
                throw new Exception("DWDocument::Open関数が失敗しました");
            }

            return;
        }

        /// <summary>
        /// DocuWorksファイルを閉じる
        /// </summary>
        public void Close()
        {
            int api_result = Xdwapi.XDW_CloseDocumentHandle(_handle);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(NLog.LogLevel.Error, api_result);
                throw new Exception("DWDocument::Close関数が失敗しました");
            }

            return;
        }

        /// <summary>
        /// DocuWorksファイルを保存する。
        /// </summary>
        public void Save()
        {
            int api_result = Xdwapi.XDW_SaveDocument(_handle);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(NLog.LogLevel.Error, api_result);
                throw new Exception("DWDocument::Save関数が失敗しました");
            }
            return;
        }

        /// <summary>
        /// ファイルを移動させる
        /// </summary>
        public void Move(string destPath)
        {
            if (_filepath == null)
            {
                return;
            }

            System.IO.File.Move(_filepath, destPath);

            _filepath = destPath;
        }

        #region 破棄パターン実装
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
        #endregion
    }
}
