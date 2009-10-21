using System.ComponentModel.Composition;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace ItalicComments
{
    [Export(typeof(IWpfTextViewCreationListener))]
    [ContentType("code")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    internal sealed class ViewCreationListener : IWpfTextViewCreationListener
    {
        [Import]
        IClassificationFormatMapService formatMapService = null;

        [Import]
        IClassificationTypeRegistryService classificationTypeService = null;

        /// <summary>
        /// When a text view is created, make all comments italicized.
        /// </summary>
        public void TextViewCreated(IWpfTextView textView)
        {
            new FormatMapWatcher(formatMapService.GetClassificationFormatMap(textView),
                                 classificationTypeService.GetClassificationType("text"));
        }
    }

    internal sealed class FormatMapWatcher
    {
        bool inUpdate = false;
        IClassificationFormatMap formatMap;
        IClassificationType text;

        public FormatMapWatcher(IClassificationFormatMap formatMap, IClassificationType text)
        {
            this.formatMap = formatMap;
            this.text = text;
            this.FixComments();

            this.formatMap.ClassificationFormatMappingChanged += FormatMapChanged;
        }

        void FormatMapChanged(object sender, System.EventArgs e)
        {
            if (!inUpdate)
                this.FixComments();
        }

        internal void FixComments()
        {
            try
            {
                inUpdate = true;

                var textFormat = formatMap.GetTextProperties(text);

                // Looks like some of these can be null; be sure to skip those
                foreach (var classification in formatMap.CurrentPriorityOrder.Where(c => c != null))
                {
                    // This won't catch, for example, "XML Doc Tag".  If you want all XML doc comments
                    // to be italicized, include whatever else you want.
                    if (classification.Classification.ToLower().Contains("comment"))
                    {
                        var properties = formatMap.GetTextProperties(classification);
                        var typeface = properties.Typeface;

                        // Add italic and (possibly) remove bold, and change to a font that has good looking
                        // italics (i.e. *not* consolas)
                        var newTypeface = new Typeface(new FontFamily("Lucida Sans"), FontStyles.Italic, FontWeights.Normal, typeface.Stretch);
                        properties = properties.SetTypeface(newTypeface);

                        // Also, make the font size a little bit smaller than the normal text size
                        properties = properties.SetFontRenderingEmSize(textFormat.FontRenderingEmSize - 1);

                        // And put it back in the format map
                        formatMap.SetTextProperties(classification, properties);
                    }
                    // Make doc tags slightly less transparent, so they fade into the background a bit.
                    // This doesn't effect strings inside tags, like the names of params, so it gives the
                    // added effect that those strings look slightly bold.
                    else if (classification.Classification.ToLower().Contains("doc tag"))
                    {
                        var properties = formatMap.GetTextProperties(classification);
                        var brush = properties.ForegroundBrush as SolidColorBrush;

                        // Make the font size a little bit smaller than the normal text size
                        properties = properties.SetFontRenderingEmSize(textFormat.FontRenderingEmSize - 1);

                        // Only do this for SolidColorBrushes, though
                        if (brush != null)
                        {
                            formatMap.SetTextProperties(classification,
                                properties.SetForegroundBrush(new SolidColorBrush(brush.Color) { Opacity = 0.7 }));
                        }
                    }
                }
            }
            finally
            {
                inUpdate = false;
            } 
        }
    }
}