namespace wordApplication
{
    public class MainClass
    {
       static void Main(string[] args)
        {
            object missing = System.Reflection.Missing.Value;
            WordManipulator wordHelper = new WordManipulator();
            string fileName = @"C:\Users\mygeno\Desktop\templateWithBookmarkNames.docx";
            string fileName1 = @"C:\Users\mygeno\Desktop\result.docx";
            wordHelper.OpenAndActive(fileName, false, true);
           
            string[] values = { "potato","tomato" };
            wordHelper.FillBookmarkByName(values, wordHelper);
            wordHelper.SaveAs(fileName1);
            wordHelper.Close();
        }
    }
}
