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
    public class HtmlBlockRenderer : WordObjectRenderer<HtmlBlock>
    {
        protected override void Write(WordRenderer renderer, HtmlBlock obj)
        {
            throw new NotImplementedException();
        }
    }
}
