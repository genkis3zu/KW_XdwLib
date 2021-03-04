namespace KW_XdwLib
{
    /// <summary>
    /// 矩形アノテーション用アトリビュート設定クラス
    /// </summary>
    public class DWAnnRectAttribute
    {
        private int _borderSytle;
        private int _borderWidth;
        private int _borderColor;
        private int _fillStyle;
        private int _fillColor;
        private int _fillTransparent;

        public DWAnnRectAttribute() { }

        public DWAnnRectAttribute(int borderSytle, int borderWidth, int borderColor, 
                                  int fillStyle, int fillColor, int fillTransparent)
        {
            _borderSytle = borderSytle;
            _borderWidth = borderWidth;
            _borderColor = borderColor;
            _fillStyle = fillStyle;
            _fillColor = fillColor;
            _fillTransparent = fillTransparent;
        }

        /// <summary>
        /// 線の表示(表示=1, 非表示=0)
        /// </summary>
        public int BorderSytle { get => _borderSytle; set => _borderSytle = value; }

        /// <summary>
        /// 線の太さ(ポイント単位)
        /// </summary>
        public int BorderWidth { get => _borderWidth; set => _borderWidth = value; }

        /// <summary>
        /// 線の色(
        /// </summary>
        public int BorderColor { get => _borderColor; set => _borderColor = value; }

        /// <summary>
        /// 矩形内の塗りつぶしをする(=1) / しない(=0)
        /// </summary>
        public int FillStyle { get => _fillStyle; set => _fillStyle = value; }

        /// <summary>
        /// 矩形内の色(
        /// </summary>
        public int FillColor { get => _fillColor; set => _fillColor = value; }

        /// <summary>
        /// 矩形内部の透過をする(=1) / しない(=0)
        /// </summary>
        public int FillTransparent { get => _fillTransparent; set => _fillTransparent = value; }

        
    }
}