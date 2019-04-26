using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax;
using DocumentFormat.OpenXml;
using Markdig.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace Markdig.Renderers.Word.Extensions
{
    public class TableRenderer : WordObjectRenderer<Markdig.Extensions.Tables.Table>
    {
        protected override void Write(WordRenderer renderer, Markdig.Extensions.Tables.Table obj)
        {
            bool hasAlreadyHeader = false;


            //renderer.Write(new Table(
            //    new TableGrid(), 
            //    new TableProperties(new TableBorders())));

            bool headered = ((Markdig.Extensions.Tables.TableRow)obj[0]).IsHeader;

            Table tab = new Table(new TableProperties(new TableStyle() { Val = (headered ? WordRenderer.STYLE_TAB_HEADERED : WordRenderer.STYLE_TAB) }));
            
            if (obj.ColumnDefinitions != null)
            {
                TableGrid tg = new TableGrid();

                foreach (var tcd in obj.ColumnDefinitions)
                {
                    var gc = new GridColumn();

                    if(tcd.Width != 0 && tcd.Width != 1)
                        gc.Width = string.Format(CultureInfo.InvariantCulture, "{0:0.##}", tcd.Width);

                    tg.Append(gc);
                }

                tab.Append(tg);
            }

            renderer/*.Write(new Paragraph(new Run(new Break())))
                .MoveUp(1)*/.Write(tab);

            foreach (var rowObj in obj)
            {
                var row = (Markdig.Extensions.Tables.TableRow)rowObj;

                if (row.IsHeader)
                {
                    if (!hasAlreadyHeader)
                    {

                    }

                    hasAlreadyHeader = true;
                }

                renderer.Write(new TableRow());

                for (int i = 0; i < row.Count; i++)
                {
                    var cellObj = row[i];
                    var cell = (Markdig.Extensions.Tables.TableCell)cellObj;

                    if (cell.ColumnSpan != 1)
                    {

                    }

                    if (cell.RowSpan != 1)
                    {

                    }

                    if (obj.ColumnDefinitions != null)
                    {
                        var columnIndex = cell.ColumnIndex < 0 || cell.ColumnIndex >= obj.ColumnDefinitions.Count
                            ? i
                            : cell.ColumnIndex;

                        columnIndex = columnIndex >= obj.ColumnDefinitions.Count ? obj.ColumnDefinitions.Count - 1 : columnIndex;
                        var alignment = obj.ColumnDefinitions[columnIndex].Alignment;
                        var width = obj.ColumnDefinitions[columnIndex].Width;
                        var tcProps = new TableCellProperties();
                        var para = new Paragraph();

                        if (alignment.HasValue)
                        {
                            var props = para.AppendChild(new ParagraphProperties(new Justification()));

                            switch (alignment)
                            {
                                case Markdig.Extensions.Tables.TableColumnAlign.Center:
                                    props.Justification.Val = JustificationValues.Center;
                                    break;
                                case Markdig.Extensions.Tables.TableColumnAlign.Right:
                                    props.Justification.Val = JustificationValues.Right;
                                    break;
                                case Markdig.Extensions.Tables.TableColumnAlign.Left:
                                    props.Justification.Val = JustificationValues.Left;
                                    break;
                            }
                        }

                        //var tc = new TableCell(
                        //    new TableCellProperties(
                        //        new TableCellWidth
                        //        {
                        //            Type = TableWidthUnitValues.Auto
                        //        }));

                        //if(width != 0 && width != 1)
                        //{
                        //    tc.TableCellProperties.TableCellWidth.Type = TableWidthUnitValues.Nil;
                        //    tc.TableCellProperties.TableCellWidth.Width = string.Format(CultureInfo.InvariantCulture, "{0:0.##}", width);
                        //}

                        renderer.Write(new TableCell()).Write(para).Write(cell);
                        renderer.MoveUp(1);
                    }
                }

                renderer.MoveUp(1);
            }
            
            renderer.MoveUp(1).Write(new Paragraph(new Run(new Break()))).MoveUp(1);
        }
    }
}
