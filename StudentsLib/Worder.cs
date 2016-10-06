using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentsLib
{
    public class Worder
    {
        public static void Replace_in_doc(Microsoft.Office.Interop.Word.Document doc, String what_to_replace, String replace_with)
        {
            Microsoft.Office.Interop.Word.Find findObject = doc.Application.Selection.Find;
            //findObject.form;
            object missing = Type.Missing;

            findObject.Text = what_to_replace;
            //findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replace_with;
            object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
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
        }

    }
}
