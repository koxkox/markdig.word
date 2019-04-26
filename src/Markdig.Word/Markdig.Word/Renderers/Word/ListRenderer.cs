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
    public class ListRenderer : WordObjectRenderer<ListBlock>
    {
        protected override void Write(WordRenderer renderer, ListBlock obj)
        {
            int level = this.GetItemLevel(obj);

            if (obj.IsOrdered)
            {
                
                
            }
            else
            {

            }
            
            foreach (var item in obj)
            {
                var listItem = (ListItemBlock)item;

                var listItemPara = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId() { Val = "ListParagraph" },
                        new NumberingProperties(
                            new NumberingLevelReference() { Val = level }, 
                            new NumberingId() { Val = 1 })));

                

                renderer.Write(listItemPara);
                renderer.WriteChildren(listItem);
                //renderer.MoveUp(1);
            }


        }

        private int GetItemLevel(ListBlock obj)
        {
            if (obj == null) return -1;
            int level = 0;
            ContainerBlock item = obj;

            while (item.Parent is ListItemBlock || item.Parent is ListBlock)
            {
                if (item.Parent is ListItemBlock)
                    level++;

                item = item.Parent;
            }

            return level;
        }
    }
}
