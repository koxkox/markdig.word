using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Parsers;

namespace Markdig.Renderers.Word
{
    public class CodeBlockRenderer : WordObjectRenderer<CodeBlock>
    {
        public HashSet<string> SupportedLanguages { get; set; }

        public CodeBlockRenderer()
        {
            //var conf = new Highlight.Configuration.DefaultConfiguration();
            
            //this.SupportedLanguages = new HashSet<string>(conf.Definitions.Select(k => k.Key), StringComparer.OrdinalIgnoreCase);
        }

        protected override void Write(WordRenderer renderer, CodeBlock obj)
        {
            var fencedCodeBlock = obj as FencedCodeBlock;
            //Highlight.Highlighter highlighter = null;
            string lang = "";

            if (fencedCodeBlock?.Info != null && this.SupportedLanguages != null && this.SupportedLanguages.Contains(fencedCodeBlock.Info))
            {
                //highlighter = new Highlight.Highlighter(new Highlight.Engines.WordOpenXmlEngine() { IgnoreStyleProperties = Highlight.Engines.IgnoreStyleProperty.FontFamily | Highlight.Engines.IgnoreStyleProperty.FontSize });
                lang = fencedCodeBlock.Info.ToUpper();
            }

            renderer.Write(new Paragraph(new ParagraphProperties(new ParagraphStyleId() { Val = WordRenderer.STYLE_CODE_BLOCK })));

            var leafBlock = obj as LeafBlock;

            if (leafBlock != null && leafBlock.Lines.Lines != null)
            {
                var lines = leafBlock.Lines;
                var slices = lines.Lines;

                for (int i = 0; i < lines.Count; i++)
                {
                    //if (highlighter != null)
                    //{
                    //    string lineStr = slices[i].Slice.Text.Substring(slices[i].Slice.Start, slices[i].Slice.Length);
                    //    int indent = lineStr.TakeWhile(c => char.IsWhiteSpace(c)).Count();
                    //    var codeParXml = highlighter.Highlight(lang, lineStr);
                    //    var codePar = new Paragraph(codeParXml);

                    //    if (indent > 0)
                    //    {
                    //        //string indentStr = string.Join("", Enumerable.Repeat(" ", indent));
                    //        var firstRun = codePar.ChildElements.FirstOrDefault() as Run;

                    //        if (firstRun != null && firstRun.HasChildren)
                    //        {
                    //            Text text = firstRun.ChildElements.Where(ce => ce is Text).FirstOrDefault() as Text;

                    //            if (text != null)
                    //                text.Text = text.Text.PadLeft(indent + text.Text.Length);
                    //        }
                    //    }
                            

                    //    foreach (Run run in codePar.ChildElements)
                    //    {
                    //        renderer.Write(run.Clone() as Run).MoveUp(1);
                    //    }
                    //}
                    //else
                    //{
                        renderer.Write(slices[i].Slice);

                        if (slices[i].Slice.Start <= slices[i].Slice.End)
                            renderer.MoveUp(1);
                    //}

                    if (i < lines.Count - 1)
                        renderer.Write(new Break()).MoveUp(1);
                }
            }

            renderer.MoveUp(1);
        }
    }
}
