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
    public class ThematicBreakRenderer : WordObjectRenderer<ThematicBreakBlock>
    {
        protected override void Write(WordRenderer renderer, ThematicBreakBlock obj)
        {
            Paragraph para = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = $"{WordRenderer.STYLE_THEMATIC_BREAK}" }));

            renderer.Write(para);
            renderer.MoveUp(1);
            renderer.Write(new Paragraph());
            renderer.MoveUp(1);
        }
    }
}
