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
    public class ParagraphRenderer : WordObjectRenderer<ParagraphBlock>
    {
        protected override void Write(WordRenderer renderer, ParagraphBlock obj)
        {
            if (!(renderer.CurrentElement is Paragraph))
                renderer.Write(new Paragraph());

            renderer.WriteLeafInline(obj);
            renderer.MoveUp(1);
        }
    }
}
