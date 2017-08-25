using Word = Microsoft.Office.Interop.Word;



namespace wordApplication
{
    public class WordManipulator
    {
        public Word.Application wordApp;
        public Word.Document wordDoc;
        object missing = System.Reflection.Missing.Value;
        

        public Word.Application WordApplication
        {
            get
            {
                return wordApp;
            }
        }
        public Word.Document WordDocument
        {
            get
            {
                return wordDoc;
            }
        }
        // The constructor helps opening word application or makes one activate
        public WordManipulator()
        {
            wordApp = new Word.Application();
            
        }
        public WordManipulator(Word.Application wordApplication)
        {
            wordApp = wordApplication;
            
        }
        
        //Open one template and make it active
        public bool OpenAndActive(string FileName, bool IsReadOnly, bool IsVisibleWin)
        {
            if (string.IsNullOrEmpty(FileName))
            {                       
                return false;
            }
            try
            {
                wordDoc = OpenOneDocument(FileName, missing, IsReadOnly, missing, missing, missing, missing, missing, missing, missing, missing, IsVisibleWin, missing, missing, missing, missing);
                wordDoc.Activate();
                return true;
            }
            catch
            {
                return false;
            }
        }
        //Close the Application
        public void Close()
        {
            if (wordDoc != null)
            {
                wordDoc.Close(ref missing, ref missing, ref missing);
                wordApp.Application.Quit(ref missing, ref missing, ref missing);
            }
        }
        //Directly save file without naming
        public void Save()
        {
            if (wordDoc == null)
            {
                wordDoc = wordApp.ActiveDocument;
            }
            wordDoc.Save();
        }
        // Save Document and give it a name
        public void SaveAs(string FileName)
        {
            if (wordDoc == null)
            {
                wordDoc = wordApp.ActiveDocument;
            }
            object objFileName = FileName;
            wordDoc.SaveAs(ref objFileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
        }
        //Creating word document in the application
        public Word.Document CreateOneDocument(object template, object newTemplate, object documentType, object visible)
        {
            return wordDoc = wordApp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
        }

        public Word.Document OpenOneDocument(object FileName, object ConfirmConversions, object ReadOnly,
            object AddToRecentFiles, object PasswordDocument, object PasswordTemplate, object Revert,
            object WritePasswordDocument, object WritePasswordTemplate, object Format, object Encoding,
            object Visible, object OpenAndRepair, object DocumentDirection, object NoEncodingDialog, object XMLTransform)
        {
            try
            {
                return wordApp.Documents.Open(ref FileName, ref ConfirmConversions, ref ReadOnly, ref AddToRecentFiles,
                   ref PasswordDocument, ref PasswordTemplate, ref Revert, ref WritePasswordDocument, ref WritePasswordTemplate,
                   ref Format, ref Encoding, ref Visible, ref OpenAndRepair, ref DocumentDirection, ref NoEncodingDialog, ref XMLTransform);
            }
            catch
            {
                return null;
            }
        }
        //If the name of the bookmark exists in the doc, then navigate to that position
        public bool GoToBookMark(string bookMarkName)
        {
            
            if (wordDoc.Bookmarks.Exists(bookMarkName))
            {
                object what = Word.WdGoToItem.wdGoToBookmark;
                object name = bookMarkName;
                GoTo(what, missing, missing, name);
                return true;
            }
            return false;
        }
        // This goto action is basically a moving cursor action
        public void GoTo(object what, object which, object count, object name)
        {
            wordApp.Selection.GoTo(ref what, ref which, ref count, ref name);
        }
        // Replace the bookmark
        public void ReplaceBookMark(string bookMarkName, string text)
        {
            bool isExist = GoToBookMark(bookMarkName);
            if (isExist)
            {
                InsertText(text);
            }
        }
        // Several mode of replacing 
        public bool Replace(string oldText, string newText, string replaceType, bool isCaseSensitive)
        {
            if (wordDoc == null)
            {
                wordDoc = wordApp.ActiveDocument;
            }
            object findText = oldText;
            object replaceWith = newText;
            object wdReplace;
            object matchCase = isCaseSensitive;
            switch (replaceType)
            {
                case "All":
                    wdReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
                    break;
                case "None":
                    wdReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                    break;
                case "FirstOne":
                    wdReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;
                    break;
                default:
                    wdReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;
                    break;
            }
            wordDoc.Content.Find.ClearFormatting();
            return wordDoc.Content.Find.Execute(ref findText, ref matchCase, ref missing, ref missing,
                  ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                  ref wdReplace, ref missing, ref missing, ref missing, ref missing);
        }
        //Insert text 
        public void InsertText(string text)
        {
            wordApp.Selection.TypeText(text);
        }
        //Insert line below
        public void InsertLineBreak()
        {
            wordApp.Selection.TypeParagraph();
        }
        //Insert a new blank page

        public void InsertPageBreak()
        {
            
            {
                object mymissing = System.Reflection.Missing.Value;
                object myunit = Word.WdUnits.wdStory;
                wordApp.Selection.EndKey(ref myunit, ref mymissing);
                object pBreak = (int)Word.WdBreakType.wdPageBreak;
                wordApp.Selection.InsertBreak(ref pBreak);
            }
        }
        //Insert picture

        public void InsertPic(string fileName)
        {
            object range = wordApp.Selection.Range;
            InsertPic(fileName, missing, missing, range);
        }

        public void InsertPic(string fileName, float width, float height)
        {
            object range = wordApp.Selection.Range;
            InsertPic(fileName, missing, missing, range, width, height);
        }
        
        public void InsertPic(string fileName, float width, float height, string caption)
        {
            object range = wordApp.Selection.Range;
            InsertPic(fileName, missing, missing, range, width, height, caption);
        }
        
        public void InsertPic(string FileName, object LinkToFile, object SaveWithDocument, object Range, float Width, float Height, string caption)
        {
            wordApp.Selection.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Range).Select();
            if (Width > 0)
            {
                wordApp.Selection.InlineShapes[1].Width = Width;
            }
            if (Height > 0)
            {
                wordApp.Selection.InlineShapes[1].Height = Height;
            }
            object Label = Word.WdCaptionLabelID.wdCaptionFigure;
            object Title = caption;
            object TitleAutoText = missing;
            object Position = Word.WdCaptionPosition.wdCaptionPositionBelow;
            object ExcludeLabel = true;
            wordApp.Selection.InsertCaption(ref Label, ref Title, ref TitleAutoText, ref Position, ref ExcludeLabel);
            MoveRight();
        }
        

        public void InsertPic(string FileName, object LinkToFile, object SaveWithDocument, object Range, float Width, float Height)
        {
            wordApp.Selection.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Range).Select();
            wordApp.Selection.InlineShapes[1].Width = Width;
            wordApp.Selection.InlineShapes[1].Height = Height;
            MoveRight();
        }

        public void InsertPic(string FileName, object LinkToFile, object SaveWithDocument, object Range)
        {
            wordApp.Selection.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Range);
        }
        //Insert bookmark

        public void InsertBookMark(string bookMarkName)
        {
            //If the bookmark you want is already existed, delete it, and create a new one
            if (wordDoc.Bookmarks.Exists(bookMarkName))
            {
                DeleteBookMark(bookMarkName);
            }
            object range = wordApp.Selection.Range;
            wordDoc.Bookmarks.Add(bookMarkName, ref range);
        }
        //Delete one bookmark

        public void DeleteBookMark(string bookMarkName)
        {
            if (wordDoc.Bookmarks.Exists(bookMarkName))
            {
                var bookMarks = wordDoc.Bookmarks;
                for (int i = 1; i <= bookMarks.Count; i++)
                {
                    object index = i;
                    var bookMark = bookMarks.get_Item(ref index);
                    if (bookMark.Name == bookMarkName)
                    {
                        bookMark.Delete();
                        break;
                    }
                }
            }
        }
        //Delete all the bookmarks in the doc

        public void DeleteAllBookMark()
        {
            for (; wordDoc.Bookmarks.Count > 0;)
            {
                object index = wordDoc.Bookmarks.Count;
                var bookmark = wordDoc.Bookmarks.get_Item(ref index);
                bookmark.Delete();
            }
        }
        //Insert one table

        public Word.Table AddTable(int NumRows, int NumColumns)
        {
            return AddTable(wordApp.Selection.Range, NumRows, NumColumns, missing, missing);
        }

        public Word.Table AddTable(int NumRows, int NumColumns, Word.WdAutoFitBehavior AutoFitBehavior)
        {
            return AddTable(wordApp.Selection.Range, NumRows, NumColumns, missing, AutoFitBehavior);
        }

        public Word.Table AddTable(Word.Range Range, int NumRows, int NumColumns, object DefaultTableBehavior, object AutoFitBehavior)
        {
            if (wordDoc == null)
            {
                wordDoc = wordApp.ActiveDocument;
            }
            return wordDoc.Tables.Add(Range, NumRows, NumColumns, ref DefaultTableBehavior, ref AutoFitBehavior);
        }

        //Insert a new row in the existed table
        public Word.Row AddRow(Word.Table table)
        {
            return AddRow(table, missing);
        }

        public Word.Row AddRow(Word.Table table, object beforeRow)
        {
            return table.Rows.Add(ref beforeRow);
        }
        public void InsertRows(int numRows)
        {
            object NumRows = numRows;
            object wdCollapseStart = Word.WdCollapseDirection.wdCollapseStart;
            wordApp.Selection.InsertRows(ref NumRows);
            wordApp.Selection.Collapse(ref wdCollapseStart);
        }

        public void MoveLeft(Word.WdUnits unit = Word.WdUnits.wdCharacter, int count = 1, int extend_flag = 0)
        {
            object extend;
            if (extend_flag == 1) extend = Word.WdMovementType.wdExtend;
            else extend = missing;
            wordApp.Selection.MoveLeft(unit, count, extend);
        }

        public void MoveUp(Word.WdUnits unit = Word.WdUnits.wdCharacter, int count = 1, int extend_flag = 0)
        {
            object extend;
            if (extend_flag == 1) extend = Word.WdMovementType.wdExtend;
            else extend = missing;
            wordApp.Selection.MoveUp(unit, count, extend);
        }
        public void MoveRight(Word.WdUnits unit = Word.WdUnits.wdCharacter, int count = 1, int extend_flag = 0)
        {
            object extend;
            if (extend_flag == 1) extend = Word.WdMovementType.wdExtend;
            else extend = missing;
            wordApp.Selection.MoveRight(unit, count, extend);
        }
        public void MoveDown(Word.WdUnits unit = Word.WdUnits.wdCharacter, int count = 1, int extend_flag = 0)
        {
            object extend;
            if (extend_flag == 1) extend = Word.WdMovementType.wdExtend;
            else extend = missing;
            wordApp.Selection.MoveDown(unit, count, extend);
        }
        

        public void SetLinesPage(int size = 40)
        {
            wordApp.ActiveDocument.PageSetup.LinesPage = size;
        }
        public void SetPageHeaderFooter(string context, int HeaderFooter = 0)
        {

           

            // Add page header
            if (wordApp.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdNormalView ||
                wordApp.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdOutlineView)
            {
                wordApp.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }
            if (HeaderFooter == 0)
            {//Set page header
                wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;
            }
            else
            {//Set page footer
                wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            }
            wordApp.Selection.HeaderFooter.LinkToPrevious = false;
            wordApp.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.Selection.HeaderFooter.Range.Text = context;
            //Break out of the setting mode of page header and setter
            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }
        //Set the style of words, adding size, bold, fontcolor and alignment 
        public void InsertFormatText(string context, int fontSize, Word.WdColor fontColor, int fontBold, string familyName, Word.WdParagraphAlignment align)
        {
              
            wordApp.Application.Selection.Font.Size = fontSize;
            wordApp.Application.Selection.Font.Bold = fontBold;
            wordApp.Application.Selection.Font.Color = fontColor;
            wordApp.Selection.Font.Name = familyName;
            wordApp.Application.Selection.ParagraphFormat.Alignment = align;
            wordApp.Application.Selection.TypeText(context);
        }


        //Set the page layout, A4 sized, landscape or portrait alignment

        public void setPageLayout(string paper = "A4", int orient = 0)
        {
            
            if (orient == 0)
            {
                wordDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            }
            else
            {
                wordDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            }
            if (paper == "A4")
            {
                wordDoc.PageSetup.PageWidth = wordApp.CentimetersToPoints(29.7F);
                wordDoc.PageSetup.PageHeight = wordApp.CentimetersToPoints(21F);
            }
        }
        //Find the table in the word doc by its index num
        public Word.Table getTable(int index)
        {
            

            int num = 0;
            foreach (Word.Table table in wordDoc.Tables)
            {
                if (num == index) return table;
                num++;
            }
            return null;
        }
        // Set the width of column
        public void SetColumnWidth(float[] widths, Word.Table tb)
        {
            if (widths.Length > 0)
            {
                int len = widths.Length;
                for (int i = 0; i < len; i++)
                {
                    tb.Columns[i + 1].Width = widths[i];
                }
            }
        }
        // Merge several column into one column
        public void MergeColumn(Word.Table tb, Word.Cell[] cells)
        {
            if (cells.Length > 1)
            {
                Word.Cell c = cells[0];
                int len = cells.Length;
                for (int i = 1; i < len; i++)
                {
                    c.Merge(cells[i]);
                }
            }
            wordApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        }
        // Find the certain bookmark name and fill it with a content
        public void FillBookmarkByName(string[]information,WordManipulator wordhelper)
        {
            // The bookmarkNames array contains the bookmark name in the document, simply replace it the one you used in you document
            string[] bookmarkNames = { "num", "num1" };
            for (int i = 0; i < bookmarkNames.Length; i++)
            {
                wordhelper.GoToBookMark(bookmarkNames[i]);
                wordhelper.InsertText(information[i]);
            }
        }
    }
}

    

