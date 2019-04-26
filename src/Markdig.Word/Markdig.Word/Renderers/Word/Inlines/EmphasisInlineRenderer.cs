using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.Renderers.Word
{
    public class EmphasisInlineRenderer : WordObjectRenderer<EmphasisInline>
    {
        //public GetTagDelegate GetTag { get; set; }
        //public delegate string GetTagDelegate(EmphasisInline obj);

        //public EmphasisInlineRenderer()
        //{
        //    this.GetTag = this.GetDefaultTag;
        //}

        protected override void Write(WordRenderer renderer, EmphasisInline obj)
        {
            //var run = renderer.CurrentElement as Run;

            //if(renderer.CurrentElement is Run)
            //{
            //    ((Run)renderer.CurrentElement).RunProperties = this.GetRunProperties(obj);
            //}

            //var run = new Run() { RunProperties = this.GetRunProperties(obj) };
            //renderer.Write(run);
            renderer.WriteChildren(obj);
            //renderer.MoveUp(1);
        }

        //public string GetDefaultTag(EmphasisInline obj)
        //{
        //    if (obj.DelimiterChar == '*' || obj.DelimiterChar == '_')
        //    {
        //        return obj.IsDouble ? "strong" : "em";
        //    }
        //    return null;
        //}

        //public RunProperties GetRunProperties(EmphasisInline obj)
        //{
        //    if (obj.DelimiterChar == '*' || obj.DelimiterChar == '_')
        //    {
        //        return obj.IsDouble ? new RunProperties() { Bold = new Bold() } : new RunProperties() { Emphasis = new Emphasis() };
        //    }

        //    return null;
        //}
    }
}
