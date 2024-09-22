using PdfSharp.Fonts;
using System;
using System.IO;

namespace ECMDocumentHelper.Helpers
{
    public class CustomFontResolver : IFontResolver
    {
        private readonly string _fontDirectory;

        public CustomFontResolver(string fontDirectory)
        {
            _fontDirectory = fontDirectory;
        }

        // Maps font-family names to the specific font files
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (familyName.Equals("LibreBarcode128Text", StringComparison.OrdinalIgnoreCase))
            {
                return new FontResolverInfo("LibreBarcode128Text");
            }
            if (familyName.Equals("LiberationSans", StringComparison.OrdinalIgnoreCase))
            {
                return new FontResolverInfo("LiberationSans");
            }

            return null;
        }

        // Loads the font data as a byte array
        public byte[] GetFont(string faceName)
        {
            if (faceName == "LibreBarcode128Text")
            {
                string fontPath = Path.Combine(_fontDirectory, "LibreBarcode128Text-Regular.ttf"); // Path to your font file
                return File.ReadAllBytes(fontPath);
            }

            if (faceName == "LiberationSans")
            {
                string fontPath = Path.Combine(_fontDirectory, "LiberationSans-Regular.ttf"); // Path to your font file
                return File.ReadAllBytes(fontPath);
            }

            throw new InvalidOperationException($"Font '{faceName}' is not available.");
        }
    }
}
