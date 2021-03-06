using System;

namespace KW_XdwLib
{
    /// <summary>
    /// 日付印用のデータ設定クラス
    /// </summary>
    public class DWAnnStampAttribute
    {
        private string _upperTitle;
        private string _userName;
        private DateTime _date;

        public DWAnnStampAttribute(string upperTitle, string userName, DateTime date)
        {
            _upperTitle = upperTitle;
            _userName = userName;
            _date = date;
        }

        public string UpperTitle { get => _upperTitle; set => _upperTitle = value; }
        public string UserName { get => _userName; set => _userName = value; }
        public DateTime Date { get => _date; set => _date = value; }
    }
}