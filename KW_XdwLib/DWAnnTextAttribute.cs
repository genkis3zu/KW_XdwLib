using FujiXerox.DocuWorks.Toolkit;

namespace KW_XdwLib
{
    /// <summary>
    /// テキストアノテーション用の設定クラス
    /// 一応すべての項目を出したが、実際には使う項目だけでいいようにする。
    /// 基本的には、文字サイズと色くらいか。
    /// デフォルト値を入れているものは、その用途。
    /// </summary>
    public class DWAnnTextAttribute
    {
        private string _text = "空文字列";
        private string _fontName;
        private int _fontStyle;
        private int _fontSize = 280;
        private int _foreColor = Xdwapi.XDW_COLOR_RED;
        private int _fontPitchAndFamily;
        private int _fontCharSet;
        private int _backColor;
        private int _wordWrap;
        private int _textDirection;
        private int _textOrientation;
        private int _lineSpace;
        private int _textSpacing;
        private int _textTopMargin;
        private int _textLeftMargin;
        private int _textRightMargin;
        private int _textBottomMargin;
        private int _textAutoResizeHeight;

        public DWAnnTextAttribute(string text)
        {
            _text = text;
        }

        /// <summary>
        /// 内容文字列
        /// </summary>
        public string Text { get => _text; set => _text = value; }

        /// <summary>
        /// 文字列のフォントフェイス名
        /// TrueTypeフォントのみが有効
        /// フォントを指定する場合は、文字セット（XDW_ATN_FontCharSet)も指定する必要がある
        /// </summary>
        public string FontName { get => _fontName; set => _fontName = value; }

        /// <summary>
        /// フォントのスタイルと文字飾り
        /// XDW_FS_ITALIC_FLAG      1   (イタリック)
        /// XDW_FS_BOLD_FLAG        2   (ボールド)
        /// XDW_FS_UNDERLINE_FLAG   4   (下線)
        /// XDW_FS_STRIKEOUT_FLAG   8   (取り消し線)
        /// </summary>
        public int FontStyle { get => _fontStyle; set => _fontStyle = value; }

        /// <summary>
        /// フォントサイズ(1/10)ポイント単位
        /// </summary>
        public int FontSize { get => _fontSize; set => _fontSize = value; }

        /// <summary>
        /// 文字の色
        /// </summary>
        public int ForeColor { get => _foreColor; set => _foreColor = value; }

        /// <summary>
        /// フォントのピッチとファミリ
        /// Win32APIのLOGFONT構造体のlfPitchAndFamilyに設定する値
        /// </summary>
        public int FontPitchAndFamilu { get => _fontPitchAndFamily; set => _fontPitchAndFamily = value; }

        /// <summary>
        /// フォントの文字セット
        /// Win32APIのLOGFONT構造体のlfCharSetに設定する値
        /// </summary>
        public int FontCharSet { get => _fontCharSet; set => _fontCharSet = value; }

        /// <summary>
        /// 背景色
        /// </summary>
        public int BackColor { get => _backColor; set => _backColor = value; }

        /// <summary>
        /// 文字列を折り返す(=1) / 折り返さない(=0)
        /// </summary>
        public int WordWrap { get => _wordWrap; set => _wordWrap = value; }

        /// <summary>
        /// 横書き(=0) / 縦書き(=1)
        /// </summary>
        public int TextDirection { get => _textDirection; set => _textDirection = value; }

        /// <summary>
        /// 文字列の回転角度(0～359) 時計回り
        /// </summary>
        public int TextOrientation { get => _textOrientation; set => _textOrientation = value; }

        /// <summary>
        /// 文字列の行間
        /// </summary>
        public int LineSpace { get => _lineSpace; set => _lineSpace = value; }

        /// <summary>
        /// テキストの文字間隔
        /// </summary>
        public int TextSpacing { get => _textSpacing; set => _textSpacing = value; }

        /// <summary>
        /// テキストアノテーションの上枠とテキスト領域の上側との間の距離
        /// </summary>
        public int TextTopMargin { get => _textTopMargin; set => _textTopMargin = value; }

        /// <summary>
        /// テキストアノテーションの左枠とテキスト領域の左側との間の距離
        /// </summary>
        public int TextLeftMargin { get => _textLeftMargin; set => _textLeftMargin = value; }

        /// <summary>
        /// テキストアノテーションの右枠とテキスト領域の右側との間の距離
        /// </summary>
        public int TextRightMargin { get => _textRightMargin; set => _textRightMargin = value; }

        /// <summary>
        /// テキストアノテーションの下枠とテキスト領域の下側との間の距離
        /// </summary>
        public int TextBottomMargin { get => _textBottomMargin; set => _textBottomMargin = value; }

        /// <summary>
        /// 高さの自動調整する(=1) / しない(=0)
        /// WordWrapが1のときのみ有効
        /// </summary>
        public int TextAutoResizeHeight { get => _textAutoResizeHeight; set => _textAutoResizeHeight = value; }
    }
}