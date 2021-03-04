using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KW_XdwLib
{
    /// <summary>
    /// 指定のフォルダ内のDocuWorks文章を、指定の構成フォルダ内へと配布する。
    /// 
    /// </summary>
    public class DWFileDistributor
    {
        private string _workspacePath;

        public DWFileDistributor(string workspacePath)
        {
            _workspacePath = workspacePath;
        }


        public void DoDistribute()
        {

        }
    }
}
