/* работает открывает ворд
using System.Runtime.InteropServices;

class Program
{

    static void Main()
    {
        var id = Type.GetTypeFromProgID("Word.Application");
        dynamic word = Activator.CreateInstance(id);
        word.Visible = true;
        object file = @"C:\1\Test.docx";
        word.Application.Documents.Open(file);
        //Marshal.ReleaseComObject(word);

    }
}


*/