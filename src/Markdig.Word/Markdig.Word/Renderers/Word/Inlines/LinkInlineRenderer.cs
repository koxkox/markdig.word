using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace Markdig.Renderers.Word.Inlines
{
    public class LinkInlineRenderer : WordObjectRenderer<LinkInline>
    {
        protected override void Write(WordRenderer renderer, LinkInline obj)
        {
            string url = (obj.GetDynamicUrl != null ? obj.GetDynamicUrl.Invoke() ?? obj.Url : obj.Url);
            HyperlinkRelationship rel = renderer.MainPart.AddHyperlinkRelationship(new Uri(url), true);
            string title = obj.Title;

            if(obj.FirstChild != null && obj.FirstChild is LiteralInline)
            {
                string inline = ((LiteralInline)obj.FirstChild).Content.ToString();

                if (!string.IsNullOrEmpty(inline))
                    title = inline;
            }

            renderer.Write(
                new Run(
                    new Hyperlink(
                        new ProofError() { Type = ProofingErrorValues.GrammarStart },
                        new Run(
                            
                            new RunProperties(
                                new RunStyle() { Val = WordRenderer.STYLE_HYPERLINK }),
                            new Text(string.IsNullOrEmpty(title) ? url : obj.Title)
                            ))
                    {
                        Id = rel.Id
                    }, 
                    new RunProperties()
                    {
                        NoProof = new NoProof()
                    }));

            renderer.MoveUp(1);
        }
    }
}
