using System.Text;
using System.Runtime.InteropServices;

namespace KW_XdwLib
{
    /// <summary>
    /// DocuWorksAPI用環境設定ファイル(INIファイル)のクラス
    /// INIファイルを使いたくはなかったが、前例に従って・・。
    /// </summary>
    public static class DWEnvIni
    {
        private const string _iniFilePath = @"C:\KWSW\env.ini";

        [DllImport("kernel32.dll")]
        private static extern int GetPrivateProfileString(
            string lpApplicationName,
            string lpKeyName,
            string lpDefault,
            StringBuilder lpReturnedString,
            int nSize,
            string lpFileName);

        [DllImport("kernel32.dll")]
        private static extern int WritePrivateProfileString(
            string lpApplicationName,
            string lpKeyName,
            string lpString,
            string lpFileName);

        public static void SetValue(string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, _iniFilePath);
        }
        /// <summary>
        /// sectionとkeyからiniファイル設定値を取得する。
        /// 指定したものがない場合は、default_valueが帰る
        /// </summary>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <param name="defaultvalue">デフォルト値</param>
        /// <returns></returns>
        public static string GetValue(string section, string key, string defaultvalue)
        {
            StringBuilder sb = new StringBuilder(256);
            GetPrivateProfileString(section, key, defaultvalue, sb, sb.Capacity, _iniFilePath);
            return sb.ToString();
        }

    }
}
