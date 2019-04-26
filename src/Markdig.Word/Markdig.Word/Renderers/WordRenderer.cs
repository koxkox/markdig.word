using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using Markdig.Syntax;
using Markdig.Helpers;
using Markdig.Syntax.Inlines;
using Markdig.Renderers.Word;
using Markdig.Renderers.Word.Inlines;
using Markdig.Renderers.Word.Extensions;
using System.IO;
using System.Xml.Linq;

namespace Markdig.Renderers
{
    public class WordRenderer : RendererBase
    {
        private bool _previousWasLine;
        private WordprocessingDocument _doc;
        private OpenXmlElement _currentElement;
        private Stream _stream;

        public const string STYLE_H = "MDH";
        public const string STYLE_H1 = "MDH1";
        public const string STYLE_H2 = "MDH2";
        public const string STYLE_H3 = "MDH3";
        public const string STYLE_H4 = "MDH4";
        public const string STYLE_H5 = "MDH5";
        public const string STYLE_H6 = "MDH6";
        public const string STYLE_HYPERLINK = "MDHYPERLINK";
        public const string STYLE_QUOTE_BLOCK = "MDQUOTEBLOCK";
        public const string STYLE_CODE_BLOCK = "MDCODEBLOCK";
        public const string STYLE_TAB = "MDTABLE";
        public const string STYLE_TAB_HEADERED = "MDTABLEHD";
        public const string STYLE_THEMATIC_BREAK = "MDTHEMATICBREAK";
        public const string STYLE_BOLD = "MDBOLDSTYLE";
        public const string STYLE_ITALIC = "MDITALICSTYLE";
        public const string STYLE_STRIKE_THROUGH = "MDSTRIKETHROUGHSTYLE";
        public const string STYLE_SUBSCRIPT = "MDSUBSCRIPTSTYLE";
        public const string STYLE_SUPERSCRIPT = "MDSUPERSCRIPTSTYLE";
        public const string STYLE_INSERTED = "MDINSERTEDSTYLE";
        public const string STYLE_MARKED = "MDMARKEDSTYLE";
        public const string STYLE_CODE = "MDCODESTYLE";

        public MainDocumentPart MainPart
        {
            get { return this._doc.MainDocumentPart; }
        }

        public OpenXmlElement CurrentElement
        {
            get { return this._currentElement; }
        }

        public bool ImplicitParagraph { get; set; }

        public WordRenderer(Stream stream, List<OpenXmlElement> styles = null, Numbering numbering = null)
        {
            this._doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
            MainDocumentPart mdp = this._doc.AddMainDocumentPart();
            mdp.Document = new Document(new Body());
            StyleDefinitionsPart sdp = mdp.AddNewPart<StyleDefinitionsPart>();
            Styles st = new Styles(styles ?? WordRenderer.GetDefaultStyles());
            st.Save(sdp);
            NumberingDefinitionsPart ndp = mdp.AddNewPart<NumberingDefinitionsPart>();
            Numbering num = numbering ?? WordRenderer.GetDefaultNumbering();
            num.Save(ndp);

            this._currentElement = mdp.Document.Body;
            
            // Default block renderers
            ObjectRenderers.Add(new CodeBlockRenderer());
            ObjectRenderers.Add(new ListRenderer());
            ObjectRenderers.Add(new HeadingRenderer());
            ObjectRenderers.Add(new HtmlBlockRenderer());
            ObjectRenderers.Add(new ParagraphRenderer());
            ObjectRenderers.Add(new QuoteBlockRenderer());
            ObjectRenderers.Add(new ThematicBreakRenderer());

            // Default inline renderers
            ObjectRenderers.Add(new AutolinkInlineRenderer());
            ObjectRenderers.Add(new CodeInlineRenderer());
            ObjectRenderers.Add(new DelimiterInlineRenderer());
            ObjectRenderers.Add(new EmphasisInlineRenderer());
            ObjectRenderers.Add(new LineBreakInlineRenderer());
            ObjectRenderers.Add(new HtmlInlineRenderer());
            ObjectRenderers.Add(new HtmlEntityInlineRenderer());
            ObjectRenderers.Add(new LinkInlineRenderer());
            ObjectRenderers.Add(new LiteralInlineRenderer());

            // Extension renderers
            ObjectRenderers.Add(new TableRenderer());
        }

        public override object Render(MarkdownObject markdownObject)
        {
            this.Write(markdownObject);
            this._doc.Save();
            this._doc.Close();
            return this._doc;
        }

        public WordRenderer Write(ref StringSlice slice)
        {
            if (slice.Start > slice.End)
            {
                return this;
            }
            return this.Write(slice.Text, slice.Start, slice.Length);
        }

        public WordRenderer Write(StringSlice slice)
        {
            return this.Write(ref slice);
        }

        //public WordRenderer Write(char content)
        //{
        //    this._previousWasLine = content == '\n';
        //    this.Write(new Text(content.ToString()));
        //    return this;
        //}

        public WordRenderer Write(string content, int offset, int length)
        {
            if (content == null)
            {
                return this;
            }

            this._previousWasLine = false;

            if (offset == 0 && content.Length == length)
            {
                this.Write(new Run(new Text(content) { Space = SpaceProcessingModeValues.Preserve }));
            }
            else
            {
                this.Write(new Run(new Text(content.Substring(offset, length)) { Space = SpaceProcessingModeValues.Preserve }));
            }

            return this;
        }

        public WordRenderer WriteLine()
        {
            //this.Writer.WriteElement(new Run(new Break()));
            this._previousWasLine = true;
            return this;
        }

        public WordRenderer WriteLine(string content)
        {
            this._previousWasLine = true;
            this.Write(new Run(new Text(content)));
            return this;
        }

        public WordRenderer WriteLeafInline(LeafBlock leafBlock)
        {
            if (leafBlock == null) throw new ArgumentNullException(nameof(leafBlock));
            var inline = (Inline)leafBlock.Inline;
            if (inline != null)
            {
                while (inline != null)
                {
                    base.Write(inline);
                    inline = inline.NextSibling;
                }
            }

            return this;
        }

        public WordRenderer Write(OpenXmlElement element)
        {
            this._currentElement = this._currentElement.AppendChild(element);
            return this;
        }

        public static Numbering GetDefaultNumbering()
        {
            var bulletLevel0 = new Level(
               new StartNumberingValue() { Val = 1 },
               new NumberingFormat() { Val = NumberFormatValues.Bullet },
               new LevelText() { Val = "" },
               new LevelJustification() { Val = LevelJustificationValues.Left },
               new ParagraphProperties(
                   new Indentation()
                   {
                       Left = "720",
                       Hanging = "360"
                   }),
               new RunProperties(
                   new RunFonts()
                   {
                       Ascii = "Symbol",
                       HighAnsi = "Symbol",
                       Hint = FontTypeHintValues.Default
                   }))
            { LevelIndex = 0 };

            var bulletLevel1 = new Level(
               new StartNumberingValue() { Val = 1 },
               new NumberingFormat() { Val = NumberFormatValues.Bullet },
               new LevelText() { Val = "o" },
               new LevelJustification() { Val = LevelJustificationValues.Left },
               new ParagraphProperties(
                   new Indentation()
                   {
                       Left = "1440",
                       Hanging = "360"
                   }),
               new RunProperties(
                   new RunFonts()
                   {
                       Ascii = "Courier New",
                       HighAnsi = "Courier New",
                       ComplexScript = "Courier New",
                       Hint = FontTypeHintValues.Default
                   }))
            { LevelIndex = 1 };

            var bulletLevel2 = new Level(
               new StartNumberingValue() { Val = 1 },
               new NumberingFormat() { Val = NumberFormatValues.Bullet },
               new LevelText() { Val = "" },
               new LevelJustification() { Val = LevelJustificationValues.Left },
               new ParagraphProperties(
                   new Indentation()
                   {
                       Left = "2160",
                       Hanging = "360"
                   }),
               new RunProperties(
                   new RunFonts()
                   {
                       Ascii = "Wingdings",
                       HighAnsi = "Wingdings",
                       Hint = FontTypeHintValues.Default
                   }))
            { LevelIndex = 2 };



            var level0 = new Level(
                new StartNumberingValue() { Val = 1 },
                new NumberingFormat() { Val = NumberFormatValues.Bullet },
                new LevelText() { Val = "" },
                new LevelJustification() { Val = LevelJustificationValues.Left },
                new ParagraphProperties(
                    new Tabs(
                        new TabStop() { Val = TabStopValues.Number, Position = 720 }),
                    new Indentation() { Left = "720", Hanging = "360" }),
                new RunProperties(
                    new RunFonts() { Ascii = "Symbol", HighAnsi = "Symbol", Hint = FontTypeHintValues.Default })) { LevelIndex = 0 };

            var level1 = new Level(
                new StartNumberingValue() { Val = 1 },
                new NumberingFormat() { Val = NumberFormatValues.Decimal },
                new LevelText() { Val = "%2." },
                new LevelJustification() { Val = LevelJustificationValues.Left },
                new ParagraphProperties(
                    new Tabs(
                        new TabStop() { Val = TabStopValues.Number, Position = 1440 }),
                    new Indentation() { Left = "1440", Hanging = "720" })) { LevelIndex = 1 };

            var level2 = new Level(
                new StartNumberingValue() { Val = 1 },
                new NumberingFormat() { Val = NumberFormatValues.Decimal },
                new LevelText() { Val = "%3." },
                new LevelJustification() { Val = LevelJustificationValues.Left },
                new ParagraphProperties(
                    new Tabs(
                        new TabStop() { Val = TabStopValues.Number, Position = 2160 }),
                    new Indentation() { Left = "2160", Hanging = "720" }))
            { LevelIndex = 1 };

            var absNum = new AbstractNum(
                new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel },
                bulletLevel0,
                bulletLevel1,
                bulletLevel2)
            { AbstractNumberId = 0 };

            var numbering = new Numbering(
                absNum,
                new NumberingInstance(
                    new AbstractNumId() { Val = 0 })
                { NumberID = 1 });

            return numbering;
        }

        public static List<OpenXmlElement> GetDefaultStyles()
        {
            var styles = new List<OpenXmlElement>();

            styles.Add(new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                        new FontSize() { Val = "24" },
                        new Color() { Val = "24292e" }))));

            // heading1 style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_H1 },
                new StyleRunProperties(
                    new FontSize() { Val = "48" }, 
                    new Bold()))
            { Type = StyleValues.Paragraph, StyleId = WordRenderer.STYLE_H1, CustomStyle = true });

            // heading2 style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_H2 },
                new StyleRunProperties(
                    new FontSize() { Val = "36" },
                    new Bold()))
            { Type = StyleValues.Paragraph, StyleId = WordRenderer.STYLE_H2, CustomStyle = true });

            // heading3 style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_H3 },
                new StyleRunProperties(
                    new FontSize() { Val = "30" },
                    new Bold()))
            { Type = StyleValues.Paragraph, StyleId = WordRenderer.STYLE_H3, CustomStyle = true });

            // heading4 style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_H4 },
                new StyleRunProperties(
                    new FontSize() { Val = "27" },
                    new Bold()))
            { Type = StyleValues.Paragraph, StyleId = WordRenderer.STYLE_H4, CustomStyle = true });

            // heading5 style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_H5 },
                new StyleRunProperties(
                    new FontSize() { Val = "21" },
                    new Bold()))
            { Type = StyleValues.Paragraph, StyleId = WordRenderer.STYLE_H5, CustomStyle = true });

            // heading6 style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_H6 },
                new StyleRunProperties(
                    new FontSize() { Val = "18" },
                    new Color() { Val = "6a737d" },
                    new Bold()))
            { Type = StyleValues.Paragraph, StyleId = WordRenderer.STYLE_H6, CustomStyle = true });

            // quote block style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_QUOTE_BLOCK },
                new StyleRunProperties(new Color() { Val = "6a737d" }), 
                new StyleParagraphProperties(
                    new Indentation() { Left = "500" },
                    new Justification() { Val = JustificationValues.Both },
                    new ParagraphBorders(
                        new LeftBorder() { Color = "dfe2e5", Val = BorderValues.Thick, Size = 24, Space = 5 },
                        new BottomBorder() { Color = "ffffff", Val = BorderValues.Single, Size = 2, Space = 0 }), 
                    new SpacingBetweenLines() { After = "500" }))
            { Type = StyleValues.Paragraph, StyleId = WordRenderer.STYLE_QUOTE_BLOCK, CustomStyle = true });

            // code style
            styles.Add(
                new Style(
                    new StyleName() { Val = WordRenderer.STYLE_CODE },
                    new StyleRunProperties(
                        new RunFonts() { Ascii = "Consolas", HighAnsi = "Consolas" },
                        new FontSize() { Val = "18" },
                        new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "f6f8fa" }))
                { Type = StyleValues.Character, StyleId = WordRenderer.STYLE_CODE, CustomStyle = true });

            // code block style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_CODE_BLOCK },
                new StyleRunProperties(
                    new RunStyle() { Val = WordRenderer.STYLE_CODE }/*,
                    new RunFonts() { Ascii = "Consolas", HighAnsi = "Consolas" },
                    new FontSize() { Val = "18" }*/),
                new StyleParagraphProperties(
                    new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "f6f8fa" },
                    new ParagraphBorders(
                        new LeftBorder() { Color = "f6f8fa", Val = BorderValues.Single, Size = 2, Space = 12 },
                        new RightBorder() { Color = "f6f8fa", Val = BorderValues.Single, Size = 2, Space = 12 },
                        new TopBorder() { Color = "f6f8fa", Val = BorderValues.Single, Size = 2, Space = 12 },
                        new BottomBorder() { Color = "f6f8fa", Val = BorderValues.Single, Size = 2, Space = 12 }), 
                    new SpacingBetweenLines() {  After = "500" }))
            { Type = StyleValues.Paragraph, StyleId = WordRenderer.STYLE_CODE_BLOCK, CustomStyle = true });

            // hyperlink style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_HYPERLINK },
                new StyleRunProperties(
                    new Color() { Val = "0366d6" }, 
                    new Underline() { Val = UnderlineValues.Single }))
            { Type = StyleValues.Character, StyleId = WordRenderer.STYLE_HYPERLINK, CustomStyle = true });

            // table style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_TAB },
                new BasedOn() { Val = "TableNormal" },
                new TableStyleRowBandSize() { Val = 1 },
                new TableStyleColumnBandSize() { Val = 1 },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "f6f8fa" }))
                { Type = TableStyleOverrideValues.Band2Horizontal }, 
                new TableProperties(new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Color = "dfe2e5", Size = 8 },
                    new BottomBorder() { Val = BorderValues.Single, Color = "dfe2e5", Size = 8 },
                    new LeftBorder() { Val = BorderValues.Single, Color = "dfe2e5", Size = 8 },
                    new RightBorder() { Val = BorderValues.Single, Color = "dfe2e5", Size = 8 },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "dfe2e5", Size = 8 },
                    new InsideVerticalBorder() { Val = BorderValues.Single, Color = "dfe2e5", Size = 8 })), 
                new TableCellProperties(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new TableCellMargin(
                    new LeftMargin() { Type = TableWidthUnitValues.Dxa, Width = "100" },
                    new RightMargin() { Type = TableWidthUnitValues.Dxa, Width = "100" })))
            { Type = StyleValues.Table, StyleId = WordRenderer.STYLE_TAB, CustomStyle = true });

            // headered table style
            styles.Add(new Style(
                new StyleName() { Val = WordRenderer.STYLE_TAB_HEADERED },
                new BasedOn() { Val = WordRenderer.STYLE_TAB },
                new TableStyleProperties(
                    new RunProperties(new Bold()),
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "ffffff" }))
                {
                    Type = TableStyleOverrideValues.FirstRow
                })
            { Type = StyleValues.Table, StyleId = WordRenderer.STYLE_TAB_HEADERED, CustomStyle = true });

            // thematic break style
            styles.Add(
                new Style(
                    new StyleName() { Val = WordRenderer.STYLE_THEMATIC_BREAK },
                    new StyleParagraphProperties(
                        new ParagraphBorders(
                            new TopBorder() { Val = BorderValues.Single, Color = "6a737d", Size = 6, Space = 1 })))
                { Type = StyleValues.Paragraph, StyleId = WordRenderer.STYLE_THEMATIC_BREAK, CustomStyle = true });

            // bold style
            styles.Add(
                new Style(
                    new StyleName() { Val = WordRenderer.STYLE_BOLD },
                    new StyleRunProperties(
                        new Bold()))
                { Type = StyleValues.Character, StyleId = WordRenderer.STYLE_BOLD, CustomStyle = true });

            // italic style
            styles.Add(
                new Style(
                    new StyleName() { Val = WordRenderer.STYLE_ITALIC },
                    new StyleRunProperties(
                        new Italic()))
                { Type = StyleValues.Character, StyleId = WordRenderer.STYLE_ITALIC, CustomStyle = true });

            // strike through style
            styles.Add(
                new Style(
                    new StyleName() { Val = WordRenderer.STYLE_STRIKE_THROUGH },
                    new StyleRunProperties(
                        new Strike()))
                { Type = StyleValues.Character, StyleId = WordRenderer.STYLE_STRIKE_THROUGH, CustomStyle = true });

            // subscript style
            styles.Add(
                new Style(
                    new StyleName() { Val = WordRenderer.STYLE_SUBSCRIPT },
                    new StyleRunProperties(
                        new FontSizeComplexScript() { Val = "20" },
                        new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript }))
                { Type = StyleValues.Character, StyleId = WordRenderer.STYLE_SUBSCRIPT, CustomStyle = true });

            // superscript style
            styles.Add(
                new Style(
                    new StyleName() { Val = WordRenderer.STYLE_SUPERSCRIPT },
                    new StyleRunProperties(
                        new FontSizeComplexScript() { Val = "20" },
                        new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript }))
                { Type = StyleValues.Character, StyleId = WordRenderer.STYLE_SUPERSCRIPT, CustomStyle = true });

            // inserted style
            styles.Add(
                new Style(
                    new StyleName() { Val = WordRenderer.STYLE_INSERTED },
                    new StyleRunProperties(
                        new Underline() { Val = UnderlineValues.Single }))
                { Type = StyleValues.Character, StyleId = WordRenderer.STYLE_INSERTED, CustomStyle = true });

            // marked style
            styles.Add(
                new Style(
                    new StyleName() { Val = WordRenderer.STYLE_MARKED },
                    new StyleRunProperties(
                        new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "ffff00" }))
                { Type = StyleValues.Character, StyleId = WordRenderer.STYLE_MARKED, CustomStyle = true });

            styles.Add(
                new Style(
                    new StyleName() { Val = "ListParagraph" },
                    new StyleParagraphProperties(
                        new Indentation() { Left = "720" },
                        new ContextualSpacing()))
                { Type = StyleValues.Paragraph, StyleId = "ListParagraph", CustomStyle = true });

            return styles;
        }

        public WordRenderer MoveUp(int levels)
        {
            for (int i = 0; i < levels; i++)
                this._currentElement = this._currentElement.Parent;

            return this;
        }
    }
}
