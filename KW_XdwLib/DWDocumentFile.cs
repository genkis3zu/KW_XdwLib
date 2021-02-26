﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FujiXerox.DocuWorks.Toolkit;

namespace KW_XdwLib
{
    public class DWDocumentFile : DWDocument
    {
        private const string _extentions = ".xdw";

        /// <summary>
        /// DocuWorks文章ファイルクラス。
        /// バインダーはDWDocumentBinderを使用。
        /// </summary>
        /// <param name="filepath">DocuWorks文章への完全パス</param>
        public DWDocumentFile(string filepath) : base()
        {
            _filepath = filepath;
        }

        /// <summary>
        /// 文章がDocuWorks文章かどうかをチェックする
        /// </summary>
        /// <returns>true:DocuWorks文章,false:XDW_DOCUMENT_INFOのDocTypeがXDW_DT_DOCUMENTでないか、GetDocumentInfomationでエラーが出ている</returns>
        public bool IsDWDocument()
        {
            // まずは、拡張子判断をする。
            if (Path.GetExtension(_filepath) != _extentions )
            {
                return false;
            }

            Xdwapi.XDW_DOCUMENT_INFO info = new Xdwapi.XDW_DOCUMENT_INFO();
            int api_result = Xdwapi.XDW_GetDocumentInformation(_handle, ref info);
            if (api_result < 0)
            {
                return false;
            }

            if (info.DocType == Xdwapi.XDW_DT_DOCUMENT)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 日付印のアノテーションを貼り付ける
        /// </summary>
        /// <returns></returns>
        public int AddStampAnnotation(DateTime date, string upperTitle, string userName)
        {
            // 日付印
            int annType = Xdwapi.XDW_AID_STAMP;
            int page = 1;
            // horPos, verPosが小さすぎるとエラーがでる？
            int horPos = 1000;
            int verPos = 1000;
            // 日付印アノテーション
            Xdwapi.XDW_AA_STAMP_INITIAL_DATA initData = new Xdwapi.XDW_AA_STAMP_INITIAL_DATA();
            // 日付アノテーションのハンドル
            Xdwapi.XDW_ANNOTATION_HANDLE stampHandle = new Xdwapi.XDW_ANNOTATION_HANDLE();

            int api_result = Xdwapi.XDW_AddAnnotation(_handle, annType, page, horPos, verPos, null, ref stampHandle);

            if (api_result < 0)
            {
                return api_result;
            }

            // 日付印の属性データを設定
            api_result = SetStampAnnAttribute(ref stampHandle, date, upperTitle, userName);
            if (api_result < 0)
            {
                return api_result;
            }

            return 1;
        }

        private int SetStampAnnAttribute(ref Xdwapi.XDW_ANNOTATION_HANDLE stampHandle, DateTime date, string upperTitle, string userName)
        {
            string yearStr = date.Year.ToString();
            string monthStr = date.Month.ToString();
            string dayStr = date.Day.ToString();

            // 日付印の色(XDW_COLOR_~ BLACK, MAROON, GREEN)
            int api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_BorderColor,
                Xdwapi.XDW_ATYPE_INT, Xdwapi.XDW_COLOR_RED);
            if (api_result < 0)
            {
                return api_result;
            }

            // 日付印の日付設定(0:自動,　1:手動)
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_DateStyle,
                Xdwapi.XDW_ATYPE_INT, 1);

            if (api_result < 0)
            {
                return api_result;
            }

            // 日付印の上欄文字列 : （着手日)
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_TopField,
                Xdwapi.XDW_ATYPE_STRING, upperTitle);

            if (api_result < 0)
            {
                return api_result;
            }

            // 日付印の下欄文字列：(名前)
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_BottomField,
                Xdwapi.XDW_ATYPE_STRING, userName);
            if (api_result < 0)
            {
                return api_result;
            }

            // 年、月、日をばらばらに設定する必要がある・・・
            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_YearField,
                Xdwapi.XDW_ATYPE_STRING, yearStr);
            if (api_result < 0)
            {
                return api_result;
            }

            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_MonthField,
                Xdwapi.XDW_ATYPE_STRING, monthStr);
            if (api_result < 0)
            {
                return api_result;
            }

            api_result = Xdwapi.XDW_SetAnnotationAttribute(_handle, stampHandle, Xdwapi.XDW_ATN_DayField,
                Xdwapi.XDW_ATYPE_STRING, dayStr);
            if (api_result < 0)
            {
                return api_result;
            }


            return 1;
        }
    }
}
