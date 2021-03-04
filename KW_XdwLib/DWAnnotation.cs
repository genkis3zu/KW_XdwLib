using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FujiXerox.DocuWorks.Toolkit;

namespace KW_XdwLib
{
    /// <summary>
    /// DocuWorksアノテーション用クラス
    /// 設定がだるすぎるので別個に作成
    /// </summary>
    public class DWAnnotation
    {
        private Xdwapi.XDW_DOCUMENT_HANDLE _handle = null;

        /// <summary>
        /// DocuWorks文章アノテーションクラス
        /// </summary>
        /// <param name="handle">アノテーションを貼り付ける対象のDocuWorks文章へのハンドル</param>
        public DWAnnotation(Xdwapi.XDW_DOCUMENT_HANDLE handle)
        {
            _handle = handle;
        }

        /// <summary>
        /// 矩形アノテーションの貼り付け
        /// </summary>
        /// <param name="page">ページ番号</param>
        /// <param name="horPos">横方向の位置</param>
        /// <param name="verPos">縦方向の位置</param>
        /// <returns>矩形アノテーションのハンドル</returns>
        public Xdwapi.XDW_ANNOTATION_HANDLE CreateRectAnnotation(int page, int horPos, int verPos, DWAnnRectAttribute rectAttr)
        {
            // 囲いの矩形
            int rectAnnType = Xdwapi.XDW_AID_RECTANGLE;

            Xdwapi.XDW_AA_RECT_INITIAL_DATA rectInitData = new Xdwapi.XDW_AA_RECT_INITIAL_DATA();

            rectInitData.Width = int.Parse(DWEnvIni.GetValue("DISTRIBUTION", "DIST_LABEL_DEFAULT_WIDTH", "1000"));
            rectInitData.Height = int.Parse(DWEnvIni.GetValue("DISTRIBUTION", "DIST_LABEL_DEFAULT_HEIGHT", "1000"));

            Xdwapi.XDW_ANNOTATION_HANDLE rectAnnHandle = new Xdwapi.XDW_ANNOTATION_HANDLE();

            int api_result = Xdwapi.XDW_AddAnnotation(_handle, rectAnnType, page, horPos, verPos, rectInitData, ref rectAnnHandle);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::AddAnnotation(XDW_AID_RECTANGLE)関数が失敗しました");
            }

            SetRectAttribute(rectAnnHandle, rectAttr);

            return rectAnnHandle;
        }

        /// <summary>
        /// 矩形アノテーション属性値設定
        /// </summary>
        /// <param name="rectAnnHanle">矩形アノテーションのハンドル</param>
        /// <param name="rectAttr">矩形アノテーション用設定クラス</param>
        private void SetRectAttribute(Xdwapi.XDW_ANNOTATION_HANDLE rectAnnHanle, DWAnnRectAttribute rectAttr)
        {
            // 線の表示・非表示
            int api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, rectAnnHanle, Xdwapi.XDW_ATN_BorderStyle, Xdwapi.XDW_ATYPE_INT, rectAttr.BorderSytle);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_BorderStyle)関数が失敗しました");
            }

            // 線の太さ
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, rectAnnHanle, Xdwapi.XDW_ATN_BorderWidth, Xdwapi.XDW_ATYPE_INT, rectAttr.BorderWidth);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_BorderWidth)関数が失敗しました");
            }

            // 線の色
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, rectAnnHanle, Xdwapi.XDW_ATN_BorderColor, Xdwapi.XDW_ATYPE_INT, rectAttr.BorderColor);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_BorderColor)関数が失敗しました");
            }
            
            // 矩形内の塗りつぶし
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, rectAnnHanle, Xdwapi.XDW_ATN_FillStyle, Xdwapi.XDW_ATYPE_INT, rectAttr.FillStyle);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_FillStyle)関数が失敗しました");
            }

            // 矩形内の色
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, rectAnnHanle, Xdwapi.XDW_ATN_FillColor, Xdwapi.XDW_ATYPE_INT, rectAttr.FillColor);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_FillColor)関数が失敗しました");
            }

            // 矩形内の透過
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, rectAnnHanle, Xdwapi.XDW_ATN_FillTransparent, Xdwapi.XDW_ATYPE_INT, rectAttr.FillTransparent);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_FillTransparent)関数が失敗しました");
            }
        }

        /// <summary>
        /// テキストアノテーションの貼り付け
        /// </summary>
        /// <param name="page">貼り付けページ</param>
        /// <param name="horPos">横方向の位置</param>
        /// <param name="verPos">縦方向の位置</param>
        /// <param name="txtAttr">テキストアノテーション設定用クラス</param>
        /// <returns></returns>
        public Xdwapi.XDW_ANNOTATION_HANDLE CreateTextAnnotation(int page, int horPos, int verPos, DWAnnTextAttribute txtAttr)
        {
            // テキストアノテーション
            int txtAnnType = Xdwapi.XDW_AID_TEXT;

            Xdwapi.XDW_ANNOTATION_HANDLE txtAnnHandle = new Xdwapi.XDW_ANNOTATION_HANDLE();

            int api_result = Xdwapi.XDW_AddAnnotation(_handle, txtAnnType, page, horPos, verPos, null, ref txtAnnHandle);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::XDW_AddAnnotation(XDW_AID_TEXT)関数が失敗しました");
            }

            SetTextAttribute(txtAnnHandle, txtAttr);

            return txtAnnHandle;
        }

        /// <summary>
        /// テキストアノテーション属性設定
        /// 今のところは、テキスト内容とテキスト色、フォントサイズのみ
        /// </summary>
        /// <param name="txtAnnHandle">テキストアノテーションのハンドル</param>
        /// <param name="txtAttr">テキストアノテーション設定用クラス</param>
        private void SetTextAttribute(Xdwapi.XDW_ANNOTATION_HANDLE txtAnnHandle, DWAnnTextAttribute txtAttr)
        {
            // テキスト内容設定
            int api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, txtAnnHandle, Xdwapi.XDW_ATN_Text, Xdwapi.XDW_ATYPE_STRING, txtAttr.Text);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::XDW_SetAnnotationAttribute(XDW_ATN_Text)関数が失敗しました");
            }

            // テキスト色設定
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, txtAnnHandle, Xdwapi.XDW_ATN_ForeColor, Xdwapi.XDW_ATYPE_INT, txtAttr.ForeColor);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::XDW_SetAnnotationAttribute(XDW_ATN_ForeColor)関数が失敗しました");
            }

            // テキストサイズ
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, txtAnnHandle, Xdwapi.XDW_ATN_FontSize, Xdwapi.XDW_ATYPE_INT, txtAttr.FontSize);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::XDW_SetAnnotationAttribute(XDW_ATN_FontSize)関数が失敗しました");
            }
        }

        /// <summary>
        /// 日付印アノテーションの追加
        /// </summary>
        /// <param name="date">日付印に入れる日付</param>
        public Xdwapi.XDW_ANNOTATION_HANDLE CreateStampAnnotation(DateTime date)
        {
            // 日付印
            int annType = Xdwapi.XDW_AID_STAMP;
            int page = 1;
            // 用紙のサイズを取得する。
            Xdwapi.XDW_PAGE_INFO_EX pageInfo = new Xdwapi.XDW_PAGE_INFO_EX();
            int api_result = Xdwapi.XDW_GetPageInformation(_handle, page, ref pageInfo);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::GetPageInfomation関数が失敗しました");
            }

            // horPos, verPosが小さすぎるとエラーがでる？
            // 左上を0, 0として、考える
            string iniHorPos = DWEnvIni.GetValue("STAMP", "STAMP_OFFSET_X", "2000");
            string iniVerPos = DWEnvIni.GetValue("STAMP", "STAMP_OFFSET_Y", "1000");
            string upperTitle = DWEnvIni.GetValue("STAMP", "STAMP_UPPER_TITLE", "着手日");
            string userName = DWEnvIni.GetValue("STAMP", "STAMP_USER_NAME", "");
            int horPos = pageInfo.Width - int.Parse(iniHorPos);
            int verPos = int.Parse(iniVerPos);

            // 日付印アノテーション
            Xdwapi.XDW_AA_STAMP_INITIAL_DATA initData = new Xdwapi.XDW_AA_STAMP_INITIAL_DATA();
            // 日付アノテーションのハンドル
            Xdwapi.XDW_ANNOTATION_HANDLE stampHandle = new Xdwapi.XDW_ANNOTATION_HANDLE();

            api_result = Xdwapi.XDW_AddAnnotation(_handle, annType, page, horPos, verPos, null, ref stampHandle);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::AddAnnotation関数が失敗しました");
            }

            // 日付印の属性データを設定
            SetStampAnnAttribute(ref stampHandle, date, upperTitle, userName);

            DWErrorLogService.InfoLog("AddStampAnnotation:" + date.ToString() + upperTitle + userName);

            return stampHandle;
        }

        /// <summary>
        /// SetAnnotationAttributeを使用して日付印に対して、データをセットしていく
        /// </summary>
        /// <param name="stampHandle">日付印アノテーションのハンドル</param>
        /// <param name="date">設定日付</param>
        /// <param name="upperTitle">日付印上部(押印種類)</param>
        /// <param name="userName">日付印下部(押印者名)</param>
        private void SetStampAnnAttribute(ref Xdwapi.XDW_ANNOTATION_HANDLE stampHandle, DateTime date, string upperTitle, string userName)
        {
            string yearStr = date.Year.ToString();
            string monthStr = date.Month.ToString();
            string dayStr = date.Day.ToString();

            // 日付印の色(XDW_COLOR_~ BLACK, MAROON, GREEN)
            int api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_BorderColor,
                Xdwapi.XDW_ATYPE_INT, Xdwapi.XDW_COLOR_RED);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_BorderColor)関数が失敗しました");
            }

            // 日付印の日付設定(0:自動,　1:手動)
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_DateStyle,
                Xdwapi.XDW_ATYPE_INT, 1);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_DateStyle)関数が失敗しました");
            }

            // 日付印の上欄文字列 : （着手日)
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_TopField,
                Xdwapi.XDW_ATYPE_STRING, upperTitle);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_TopField)関数が失敗しました");
            }

            // 日付印の下欄文字列：(名前)
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_BottomField,
                Xdwapi.XDW_ATYPE_STRING, userName);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_BottomField)関数が失敗しました");
            }

            // 年、月、日をばらばらに設定する必要がある・・・
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_YearField,
                Xdwapi.XDW_ATYPE_STRING, yearStr);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_YearField)関数が失敗しました");
            }

            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_MonthField,
                Xdwapi.XDW_ATYPE_STRING, monthStr);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_MonthField)関数が失敗しました");
            }

            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_DayField,
                Xdwapi.XDW_ATYPE_STRING, dayStr);
            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::SetAnnotationAttribute(XDW_ATN_DayField)関数が失敗しました");
            }

            DWErrorLogService.InfoLog("SetStampAnnAttribute:" + date.ToString() + upperTitle + userName);

            return;
        }

    }
}
