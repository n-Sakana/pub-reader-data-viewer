// The native host and WebView use the same fixed Win98 face colour.
// Caption values are Windows COLORREF (0x00BBGGRR), not RGB.
namespace ReaderDataViewer
{
    internal static class Win98
    {
        internal static readonly System.Drawing.Color Background =
            System.Drawing.Color.FromArgb(212, 208, 200);
        internal const int Caption = 0x501b08;
        internal const int CaptionText = 0xffffff;
        internal const int Border = 0x808080;
    }
}
