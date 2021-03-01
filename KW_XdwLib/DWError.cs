using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FujiXerox.DocuWorks.Toolkit;
using NLog;

namespace KW_XdwLib
{
    /// <summary>
    /// DocuWorksAPIエラー一覧
    /// </summary>
    public static class DWErrorLogService
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// DocuWorksAPI エラーログ
        /// </summary>
        /// <param name="logLevel"></param>
        /// <param name="api_result"></param>
        public static void APIErrorLog(int api_result)
        {
            string logMessage = GetErrorCode(api_result);

            _logger.Error(logMessage + ":" + api_result.ToString());
        }

        /// <summary>
        /// NLogのInfoログ
        /// </summary>
        /// <param name="message"></param>
        public static void InfoLog(string infoMessage)
        {
            _logger.Info(infoMessage);
        }

        /// <summary>
        /// NLogのErrorログ
        /// </summary>
        /// <param name="v"></param>
        public static void ErrLog(string errMessage)
        {
            _logger.Error(errMessage);
        }
        /// <summary>
        /// 各APIからのエラーコード戻り値から、エラーを特定して返す。
        /// </summary>
        /// <param name="api_result">DocuWorksAPIからのエラーコード</param>
        /// <returns>エラーメッセージ</returns>
        private static string GetErrorCode(int api_result)
        {
            StringBuilder sb = new StringBuilder();
            switch (api_result)
            {
                case Xdwapi.XDW_E_ACCESSDENIED:
                    sb.Append("アクセス権がないため、アクセスが拒否されました");
                    break;
                case Xdwapi.XDW_E_APPLICATION_FAILED:
                    sb.Append("アプリケーションの起動に失敗");
                    break;
                case Xdwapi.XDW_E_BAD_NETPATH:
                    sb.Append("指定されたファイルを開くことができません");
                    break;
                case Xdwapi.XDW_E_BAD_FORMAT:
                    sb.Append("指定されたファイルは正しいフォーマットではありません");
                    break;
                case Xdwapi.XDW_E_CANCELED:
                    sb.Append("キャンセルされました");
                    break;
                case Xdwapi.XDW_E_DISK_FULL:
                    sb.Append("ディスク容量がいっぱいです");
                    break;
                case Xdwapi.XDW_E_FILE_EXISTS:
                    sb.Append("すでに同じ名前のファイルが存在します");
                    break;
                case Xdwapi.XDW_E_FILE_NOT_FOUND:
                    sb.Append("指定されたファイルが見つかりません");
                    break;
                case Xdwapi.XDW_E_INFO_NOT_FOUND:
                    sb.Append("問い合わせた情報が存在しません");
                    break;
                case Xdwapi.XDW_E_INSUFFICIENT_BUFFER:
                    sb.Append("バッファ領域不足");
                    break;
                case Xdwapi.XDW_E_INVALIDARG:
                    sb.Append("関数の引数に仕様外の値を渡した");
                    break;
                case Xdwapi.XDW_E_INVALID_ACCESS:
                    sb.Append("指定された操作をする権限がありません");
                    break;
                case Xdwapi.XDW_E_INVALID_NAME:
                    sb.Append("名前が正しくありません");
                    break;
                case Xdwapi.XDW_E_INVALID_OPERATION:
                    sb.Append("無効な操作です");
                    break;
                case Xdwapi.XDW_E_NEWFORMAT:
                    sb.Append("DocuWorksファイルのバージョンが新しいため開くことができません");
                    break;
                case Xdwapi.XDW_E_NOT_INSTALLED:
                    sb.Append("DocuWorksがインストールされていません");
                    break;
                case Xdwapi.XDW_E_OUTOFMEMORY:
                    sb.Append("メモリの確保に失敗tしました");
                    break;
                case Xdwapi.XDW_E_SHARING_VIOLATION:
                    sb.Append("ファイルが使用中です");
                    break;
                case Xdwapi.XDW_E_SIGNATURE_MODULE:
                    sb.Append("署名モジュールに固有のエラーが発生しました");
                    break;
                case Xdwapi.XDW_E_UNEXPECTED:
                    sb.Append("予期しないエラーが発生しました");
                    break;
                case Xdwapi.XDW_E_WRITE_FAULT:
                    sb.Append("書き込み中にエラーが発生しました");
                    break;
                default:
                    sb.Append("不明なエラーが発生しました:" + api_result.ToString());
                    break;
            }

            return sb.ToString();
        }



        /// <summary>
        /// エラーメッセージと、呼出し元の関数名
        /// </summary>
        /// <param name="api_result">エラーコード</param>
        /// <param name="callbackName">発生元の関数名</param>
        /// <returns></returns>
        public static string GetErrorMessage(int api_result, string callbackName)
        {
            string message = GetErrorCode(api_result);

            return message + "CallBack:" + callbackName;
        }
    }
}
