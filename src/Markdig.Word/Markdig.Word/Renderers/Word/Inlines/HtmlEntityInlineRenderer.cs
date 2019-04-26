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
    public class HtmlEntityInlineRenderer : WordObjectRenderer<HtmlEntityInline>
    {
        protected override void Write(WordRenderer renderer, HtmlEntityInline obj)
        {
            renderer.Write(new Run(new Text(obj.Transcoded.Text)));
            renderer.MoveUp(1);
        }
    }
}
