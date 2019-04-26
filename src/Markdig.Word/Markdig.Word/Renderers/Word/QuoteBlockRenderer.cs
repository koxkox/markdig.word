using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.Renderers.Word
{
    public class QuoteBlockRenderer : WordObjectRenderer<QuoteBlock>
    {
        protected override void Write(WordRenderer renderer, QuoteBlock obj)
        {
            renderer.Write(new Paragraph(new ParagraphProperties(new ParagraphStyleId() { Val = WordRenderer.STYLE_QUOTE_BLOCK })));
            renderer.WriteChildren(obj);
            //renderer.Write(new Paragraph(new Run(new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.None })));
            //renderer.MoveUp(1);
        }
    }
}
