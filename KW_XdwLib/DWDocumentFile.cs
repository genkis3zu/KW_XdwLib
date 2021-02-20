using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FujiXerox.DocuWorks.Toolkit;

namespace KW_XdwLib
{
    public class DWDocumentFile : DWDocument
    {
        /// <summary>
        /// DocuWorks文章に対しては
        /// </summary>
        /// <param name="filepath"></param>
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


        public int AddAnnotation()
        {
            return 1;
        }
    }
}
