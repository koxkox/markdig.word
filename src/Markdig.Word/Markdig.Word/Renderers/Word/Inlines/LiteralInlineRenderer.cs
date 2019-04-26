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
    public class LiteralInlineRenderer : WordObjectRenderer<LiteralInline>
    {
        protected override void Write(WordRenderer renderer, LiteralInline obj)
        {
            var run = new Run(new Text(obj.Content.ToString()));

            if(obj.Parent is EmphasisInline)
            {
                run.RunProperties = this.GetRunProperties((EmphasisInline)obj.Parent);
            }

            renderer.Write(run);
            renderer.MoveUp(1);
        }

        public RunProperties GetRunProperties(EmphasisInline obj)
        {
            string style = "";

            switch(obj.DelimiterChar)
            {
                case '*':
                case '_':
                    style = obj.IsDouble ? WordRenderer.STYLE_BOLD : WordRenderer.STYLE_ITALIC;
                    break;
                case '~':
                    style = obj.IsDouble ? WordRenderer.STYLE_STRIKE_THROUGH : WordRenderer.STYLE_SUBSCRIPT;
                    break;
                case '^':
                    style = WordRenderer.STYLE_SUPERSCRIPT;
                    break;
                case '+':
                    style = WordRenderer.STYLE_INSERTED;
                    break;
                case '=':
                    style = WordRenderer.STYLE_MARKED;
                    break;
            }

            return new RunProperties(new RunStyle() { Val = style });
        }
    }
}
