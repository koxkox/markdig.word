using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.Renderers.Word.Inlines
{
    public class LineBreakInlineRenderer : WordObjectRenderer<LineBreakInline>
    {
        protected override void Write(WordRenderer renderer, LineBreakInline obj)
        {
            throw new NotImplementedException();
        }
    }
}
