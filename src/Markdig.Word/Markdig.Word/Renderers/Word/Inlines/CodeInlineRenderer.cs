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
    public class CodeInlineRenderer : WordObjectRenderer<CodeInline>
    {
        protected override void Write(WordRenderer renderer, CodeInline obj)
        {
            var run = new Run(
                new RunProperties(
                    new RunStyle() { Val = WordRenderer.STYLE_CODE }), 
                new Text($" {obj.Content} ") { Space = SpaceProcessingModeValues.Preserve });

            renderer.Write(run);
            renderer.MoveUp(1);
        }
    }
}
