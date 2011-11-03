using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Diagnostics;

namespace GeoReports.Words
{
    public class WordDocument
    {
        
        public WordDocument(String templateName, bool visible)
        {

            List<int> startedWords = new List<int>();

            foreach (Process p in System.Diagnostics.Process.GetProcessesByName("winword"))
            {
                startedWords.Add(p.Id);
            }

            Object missingValue = Missing.Value;
            Object template = templateName;
            if (wordapp == null)
                wordapp = new Word.Application();

            foreach (Process p in System.Diagnostics.Process.GetProcessesByName("winword"))
            {
                if (!startedWords.Contains(p.Id))
                {
                    IDs.Add(p.Id);
                }
            }
            wordapp.Visible = visible;
            doc = wordapp.Documents.Add(ref template,
            ref missingValue, ref missingValue, ref missingValue);
            doc.ActiveWindow.Selection.MoveEnd(Word.WdUnits.wdStory);

            
            tempFileName = Path.Combine(Path.GetDirectoryName(templateName), tempFileName);
        }

        public static void KillAll()
        {
            foreach (Process p in System.Diagnostics.Process.GetProcessesByName("winword"))
            {
                try
                {
                    if (IDs.Contains(p.Id))
                    {
                        p.Kill();
                        p.WaitForExit();
                    }
                }
                catch (Exception ex)
                {
                   
                }

            }
            IDs.Clear();
        }
        public WordDocument()
        {
            Object missingValue = Missing.Value;
            Object template = Type.Missing;
            if (wordapp == null)
                wordapp = new Word.Application();
            wordapp.Visible = true;
            doc = wordapp.Documents.Add(ref template,
            ref missingValue, ref missingValue, ref missingValue);
        }

        public void OpenDocument(String fileName)
        {
            wordapp.Documents.Open(fileName);
        }

     public void Close()
       {
           wordapp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
           wordapp = null;

           if (File.Exists(tempFileName) == true) File.Delete(tempFileName);
       } 
       public void CloseDocument()
       {
           doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
       } 


        public void Dispose()
        {
            wordapp = null;
        }

        public  void GetVisible()
        {
            if (wordapp != null)
                wordapp.Visible = true;
        }

        public void ReplaceVariables(String variableName, String value)
        {
            var varible = doc.Variables[variableName];
            if (varible == null) throw new Exception("Ошибка изменения переменной в файле: переменная с именем " + variableName + " не найдена");
            varible.Value = value;
            doc.Fields.Update();
        }
        public static WordDocument MergeDocuments(WordDocument[] documents)
        {
            WordDocument result = new WordDocument();
            for (int i = 0; i < documents.Length; i++)
            {
                result.AddDocument(documents[i]);
                //documents[i].doc.ActiveWindow.Selection.WholeStory();
                //documents[i].doc.ActiveWindow.Selection.Copy();
                //result.doc.ActiveWindow.Selection.Paste();
            }
            return result;
        }

        public void AddDocument(WordDocument document)
        {
            document.doc.ActiveWindow.Selection.WholeStory();
            document.doc.ActiveWindow.Selection.Copy();
            this.doc.ActiveWindow.Selection.Paste();
        }

        private String GetPictureSmaller(string FileName)
        {
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(FileName);
            System.Drawing.Size s = bm.Size;

            s.Width = (int)(s.Width * scale);
            s.Height = (int)(s.Height * scale);

            System.Drawing.Bitmap b = new System.Drawing.Bitmap(bm, s);
            b.Save(tempFileName, System.Drawing.Imaging.ImageFormat.Jpeg);

            return tempFileName;
        }

        public void ReplaceIncludePicture(int shapeNumber, Uri value, float resizeCoeffisient)
        {
            var varible = doc.InlineShapes[shapeNumber];
            if (varible == null) throw new Exception("Ошибка изменения переменной в файле: переменная с именем " + shapeNumber + " не найдена");

            try
            {
                varible.Range.InlineShapes.AddPicture(GetPictureSmaller(value.LocalPath), Type.Missing, Type.Missing, doc.InlineShapes[1].Range);
            }
            catch (Exception ex) 
            { }

            resizeCoeffisient = wordapp.InchesToPoints(0.3937f * 8.0f) / doc.InlineShapes[shapeNumber].Height;
            doc.InlineShapes[shapeNumber].Height = wordapp.InchesToPoints(0.3937f * 8.0f);
            doc.InlineShapes[shapeNumber].Width *= resizeCoeffisient;
            varible.Range.InlineShapes[1].Delete();
            doc.Fields.Update();
        }

        public void AddDataToTable(int tableIndex, Object[] info)
        {
            Word.Table table = doc.Tables[tableIndex];
            if (info.Length != table.Columns.Count) throw new Exception("Ошибка добавления строки в таблицу: размерности не совпадают");
            Object missingValue = Missing.Value;
            table.Rows.Add(ref missingValue);
            int rowCount = table.Rows.Count;
            for (int i = 1; i <= table.Columns.Count; i++)
            {
                if (info[i - 1] is Uri)
                {
                    try
                    {
                        table.Cell(rowCount, i).Range.InlineShapes.AddPicture(GetPictureSmaller(((Uri)info[i - 1]).LocalPath), Missing.Value, Missing.Value, Missing.Value);
                    }
                    catch (Exception) { }
                }
                else table.Cell(rowCount, i).Range.Text = info[i - 1].ToString();
            }
        }
        public void Save(String fileName)
        {
            try
            {
                doc.SaveAs(fileName, Word.WdSaveFormat.wdFormatDocument97);
            }
            catch (Exception ex)
            { wordapp.Visible = true; }
        }

        private  Word.Application wordapp;
        public Word.Document doc;
        float scale = 0.10F;
        String tempFileName = "temp.jpg";
        private static List<int> IDs = new List<int>();
    }
}
