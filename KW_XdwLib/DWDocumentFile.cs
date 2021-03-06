using System;
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
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFile::GetDocumentInformation関数が失敗しました");
            }

            if (info.DocType == Xdwapi.XDW_DT_DOCUMENT)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 振り分け先指定アノテーションを貼り付ける
        /// 矩形の囲い＋テキストの形式なのでAddAnnotationOnParentAnnotationが必要？
        /// </summary>
        /// <param name="distLabel"></param>
        public void AddDistributionAnnotation(string distLabel)
        {
            int page = 1;

            // 用紙のサイズを取得する。
            Xdwapi.XDW_PAGE_INFO_EX pageInfo = GetPageInfo(page);

            string offsetHorPos = DWEnvIni.GetValue("DISTRIBUTION", "DIST_LABEL_POS_OFFSET_X", "1000");
            string offsetVerPos = DWEnvIni.GetValue("DISTRIBUTION", "DIST_LABEL_POS_OFFSET_Y", "1000");
            int width = int.Parse(DWEnvIni.GetValue("DISTRIBUTION", "DIST_LABEL_DEFAULT_WIDTH", "1000"));
            int height = int.Parse(DWEnvIni.GetValue("DISTRIBUTION", "DIST_LABEL_DEFAULT_HEIGHT", "1000"));
            int horPos = pageInfo.Width - int.Parse(offsetHorPos);
            int verPos = pageInfo.Height - int.Parse(offsetVerPos);

            DWAnnRectAttribute rectAttr = new DWAnnRectAttribute(1, 3, Xdwapi.XDW_COLOR_RED, 0, Xdwapi.XDW_COLOR_NONE, 1);
            DWAnnTextAttribute textAttr = new DWAnnTextAttribute(distLabel);
            DWAnnotation ann = new DWAnnotation(_handle);

            // 四角囲い + テキスト
            ann.CreateTextAnnotation(page, horPos + 100, verPos + 100, textAttr);
            ann.CreateRectAnnotation(width, height, page, horPos, verPos, rectAttr);

        }

        /// <summary>
        /// 着手日のアノテーションを貼り付ける
        /// txtStampがtrueの時は日付印ではなくて、テキストで
        /// </summary>
        public void AddStampAnnotation(DateTime date, bool txtStamp)
        {
            int page = 1;

            Xdwapi.XDW_PAGE_INFO_EX pageInfo = GetPageInfo(page);

            // horPos, verPosが小さすぎるとエラーがでる？
            // 左上を0, 0として、考える
            string iniHorPos = DWEnvIni.GetValue("STAMP", "STAMP_OFFSET_X", "2000");
            string iniVerPos = DWEnvIni.GetValue("STAMP", "STAMP_OFFSET_Y", "1000");
            string upperTitle = DWEnvIni.GetValue("STAMP", "STAMP_UPPER_TITLE", "着手日");
            string userName = DWEnvIni.GetValue("STAMP", "STAMP_USER_NAME", "");
            int horPos = pageInfo.Width - int.Parse(iniHorPos);
            int verPos = int.Parse(iniVerPos);
            DWAnnotation ann = new DWAnnotation(_handle);


            if (txtStamp)
            {
                string dateString = "";

                DWAnnTextAttribute txtAttr = new DWAnnTextAttribute(dateString);

                ann.CreateTextAnnotation(page, horPos, verPos, txtAttr);
            }
            else
            {

                DWAnnStampAttribute stampAttr = new DWAnnStampAttribute(upperTitle, userName, date);

                ann.CreateStampAnnotation(page, horPos, verPos, stampAttr);
            }
            return;
        }

        /// <summary>
        /// 右上や右下に押印するために、ページのサイズ情報を取得するための
        /// GetPageInformation
        /// </summary>
        /// <param name="page"></param>
        /// <returns></returns>
        private Xdwapi.XDW_PAGE_INFO_EX GetPageInfo(int page)
        {
            var info = new Xdwapi.XDW_PAGE_INFO_EX();
            int api_result = Xdwapi.XDW_GetPageInformation(_handle, page, ref info);

            if (api_result < 0)
            {
                DWErrorLogService.APIErrorLog(api_result);
                throw new Exception("DWDocumentFiles::GetPageInfomation関数が失敗しました");
            }

            return info;
        }
    }
}
