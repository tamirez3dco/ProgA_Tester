using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentsLib
{
    public class Worder
    {
        public static void AdjustLines(List<RunLine> mylines, String folderPath)
        {
            for (int i = 0; i < mylines.Count; i++)
            {
                RunLine line = mylines[i];
                if (line.s != Source.ERROR) continue;
                int errorFirstLine = i;
                String totalError = String.Empty;
                int j;
                for (j = i; j < mylines.Count; j++)
                {
                    RunLine errorLine = mylines[j];
                    if (errorLine.s != Source.ERROR) break;
                    if (errorLine.text == null) continue;
                    if (errorLine.text.Trim() == String.Empty) continue;
                    DirectoryInfo dif = new DirectoryInfo(folderPath);
                    DirectoryInfo sourcePath = dif.Parent.Parent.Parent;
                    totalError += (errorLine.text.Replace(sourcePath.FullName + "\\", String.Empty)+"\n");
                }
                mylines.RemoveRange(i, j - i);
                if (totalError.Trim().Length > 0) mylines.Insert(i, new RunLine(Source.ERROR, totalError));
                i = j;
            }
        }
        public static String LinesToTable(List<RunLine> mylines, String folderPath)
        {
            String wordTableFilePath = folderPath + "\\run_table.docx";
            AdjustLines(mylines, folderPath);
            Application oWord = new Application();
            oWord.Visible = true;
            Document wordDoc = oWord.Documents.Add();
            wordDoc.Paragraphs.Format.SpaceAfter = 0;
            Paragraph par1 = wordDoc.Paragraphs.Add();
            //par1.Range.Text = String.Format("הטבלה הבאה מכיל את הרצת התוכנית שלך עד לשלב בו התרחשה שגיאה או שהתוכנית נתקעה:");
            par1.Range.Text = String.Format("Dummmmmmy:");
            //          par1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            //          par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            //          par1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            Range endDoc = wordDoc.Content;
            endDoc.Collapse(WdCollapseDirection.wdCollapseEnd);
            object start = 0;
            object end = 0;
            Range tableRange = wordDoc.Range(ref start, ref end);

            Table table = wordDoc.Tables.Add(endDoc, mylines.Count + 1, 3);
            table.TableDirection = WdTableDirection.wdTableDirectionLtr;
            table.Spacing = 0;
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Cell(1, 1).Range.Text = "INPUT";
            table.Cell(1, 1).Range.Font.Bold = -1;
            table.Cell(1, 2).Range.Text = "OUTPUT";
            table.Cell(1, 2).Range.Font.Bold = -1;
            table.Cell(1, 3).Range.Text = "ERROR";
            table.Cell(1, 3).Range.Font.Bold = -1;

            for (int l = 0; l < mylines.Count; l++)
            {
                RunLine line = mylines[l];
                Cell cell = table.Cell(l + 2, (int)(line.s));
                cell.Range.Text = line.text;
                if (line.s == Source.ERROR) cell.Range.Font.Size = 9;
            }
            Replace_in_doc(wordDoc, "Dummmmmmy:", String.Format("הטבלה הבאה מכילה את הרצת התוכנית שלך עד לשלב בו התרחשה שגיאה או שהתוכנית נתקעה:"));
            wordDoc.Application.Selection.ParagraphFormat.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            wordDoc.Application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            wordDoc.Application.Selection.ParagraphFormat.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            wordDoc.Application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            /*
                        start = 0;
                        end = 0;
                        Paragraph lineParagraph = wordDoc.Paragraphs.Add(wordDoc.Range(ref start, ref end));
                        lineParagraph.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
                        lineParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                        lineParagraph.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
                        lineParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                        lineParagraph.Range.Text = String.Format("הטבלה הבאה מכיל את הרצת התוכנית שלך עד לשלב בו התרחשה שגיאה או שהתוכנית נתקעה:");
            */

            object filePath = wordTableFilePath;
            object missing = Type.Missing;
            wordDoc.SaveAs(ref filePath,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

            wordDoc.Close();
            oWord.Quit();
            return wordTableFilePath;
        }

        public static void Underline_in_doc(Microsoft.Office.Interop.Word.Document doc, String what_to_replace)
        {
            Find findObject = doc.Application.Selection.Find;
            object missing = Type.Missing;

            findObject.Text = what_to_replace;
            object replaceAll = WdReplace.wdReplaceNone;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
            doc.Application.Selection.Font.Underline = WdUnderline.wdUnderlineSingle;
            doc.Application.Selection.Font.UnderlineColor = WdColor.wdColorBlack;
            doc.Application.Selection.Collapse();
        }


        public static void Replace_in_doc(Microsoft.Office.Interop.Word.Document doc, String what_to_replace, String replace_with)
        {
            Find findObject = doc.Application.Selection.Find;
            object missing = Type.Missing;

            findObject.Text = what_to_replace;
            findObject.Replacement.Text = replace_with;
            object replaceAll = WdReplace.wdReplaceAll;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }

        public static InlineShape Replace_to_picture(Microsoft.Office.Interop.Word.Document doc, String what_to_replace, String picture_path)
        {
            Microsoft.Office.Interop.Word.Find findObject = doc.Application.Selection.Find;
            //findObject.form;
            object missing = Type.Missing;

            findObject.Text = what_to_replace;
            object replaceNone = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceNone, ref missing, ref missing, ref missing, ref missing);

            InlineShape shape = doc.Application.Selection.InlineShapes.AddPicture(picture_path);
            return shape;
        }


        public static void English_Format_By_Search(Microsoft.Office.Interop.Word.Document doc, String what_to_replace)
        {
            doc.Application.Selection.Collapse();
            Microsoft.Office.Interop.Word.Find findObject = doc.Application.Selection.Find;
            //findObject.form;
            object missing = Type.Missing;

            findObject.Text = what_to_replace;
            object replaceNone = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceNone, ref missing, ref missing, ref missing, ref missing);
            doc.Application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            doc.Application.Selection.ParagraphFormat.ReadingOrder = WdReadingOrder.wdReadingOrderLtr;
            doc.Application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            doc.Application.Selection.Collapse();
        }

    }
}
