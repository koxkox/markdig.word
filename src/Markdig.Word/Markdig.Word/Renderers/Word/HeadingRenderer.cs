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
    public class HeadingRenderer : WordObjectRenderer<HeadingBlock>
    {
        protected override void Write(WordRenderer renderer, HeadingBlock obj)
        {
            Paragraph para = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = $"{WordRenderer.STYLE_H}{obj.Level}" }));

            renderer.Write(para);
            renderer.WriteLeafInline(obj);
            renderer.MoveUp(1);
        }
    }
}
