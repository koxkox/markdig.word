using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Renderers;
using Markdig.Syntax;
using System.IO;

namespace Markdig.Word
{
    public static partial class Markdown
    {
        public static MarkdownDocument ToWordDocument(string markdown, Stream stream, MarkdownPipeline pipeline = null)
        {
            if (markdown == null) throw new ArgumentNullException(nameof(markdown));
            pipeline = pipeline ?? new MarkdownPipelineBuilder().Build();
            var renderer = new WordRenderer(stream);
            pipeline.Setup(renderer);
            var document = Markdig.Markdown.Parse(markdown, pipeline);
            renderer.Render(document);

            return document;
        } 
    }
}
