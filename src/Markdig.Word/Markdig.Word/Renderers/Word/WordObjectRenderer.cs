using Markdig.Syntax;

namespace Markdig.Renderers.Word
{
    public abstract class WordObjectRenderer<TObject> : MarkdownObjectRenderer<WordRenderer, TObject> 
        where TObject : MarkdownObject
    {
    }
}
