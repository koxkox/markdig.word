using System;
using System.IO;

namespace Markdig.Word.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var stream = File.OpenRead("sample2.md"))
            {
                using (var reader = new StreamReader(stream))
                {
                    var markdown = reader.ReadToEnd();
                    var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();

                    using (var fs = new FileStream("sample.docx", FileMode.Create))
                    {
                        Markdown.ToWordDocument(markdown, fs, pipeline);
                    }
                    
                }
            }
        }
    }
}
