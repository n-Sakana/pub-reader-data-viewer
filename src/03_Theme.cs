// Presentation metadata only. The business settings schema is not extended.
// C# 5 / .NET Framework. Theme is fixed for the lifetime of the process.
using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace ReaderDataViewer
{
    internal sealed class RdvTheme
    {
        private static RdvTheme current;
        public string Id;
        public string Motion;
        public System.Drawing.Color Background;
        public int Caption;
        public int CaptionText;
        public int Border;
        public bool Modern { get { return Id != "win98"; } }

        public static RdvTheme Current
        {
            get
            {
                if (current == null) { current = Load(App.BaseDirectory); }
                return current;
            }
        }

        private static System.Drawing.Color ParseColour(string text)
        {
            int rgb;
            if (text == null || text.Length != 7 || text[0] != '#' ||
                !int.TryParse(text.Substring(1), NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture, out rgb))
            {
                throw new InvalidDataException("Theme colour must be #RRGGBB.");
            }
            return System.Drawing.Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255);
        }

        private static int ColourRef(System.Drawing.Color colour)
        {
            return colour.R | (colour.G << 8) | (colour.B << 16);
        }

        internal static RdvTheme Load(string directory)
        {
            RdvTheme theme = new RdvTheme();
            theme.Id = "win98";
            theme.Motion = "off";
            theme.Background = ParseColour("#d4d0c8");
            theme.Caption = ColourRef(ParseColour("#081b50"));
            theme.CaptionText = ColourRef(ParseColour("#ffffff"));
            theme.Border = ColourRef(ParseColour("#808080"));
            if (string.IsNullOrEmpty(directory)) { return theme; }
            string path = Path.Combine(directory, "web", "theme.json");
            if (!File.Exists(path)) { return theme; }
            Rdv3Json json = Rdv3Json.Parse(File.ReadAllText(path, Encoding.UTF8));
            json.Only("schema", "id", "motion", "background", "caption", "captionText", "border");
            if (json.Int("schema", 1, 1) != 1) { throw new InvalidDataException("Unsupported theme schema."); }
            theme.Id = json.Word("id", "win98", "win98", "apple", "material", "fluent", "carbon", "spectrum");
            theme.Motion = json.Word("motion", "off", "auto", "off");
            if (!theme.Modern) { theme.Motion = "off"; }
            theme.Background = ParseColour(json.Need("background"));
            theme.Caption = ColourRef(ParseColour(json.Need("caption")));
            theme.CaptionText = ColourRef(ParseColour(json.Need("captionText")));
            theme.Border = ColourRef(ParseColour(json.Need("border")));
            return theme;
        }
    }
}
