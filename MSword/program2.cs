/*
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;


    public class program2
    {
        Word.Application word = new Word.Application();
        Word.Document doc = new Word.Document();
        object fileName = @"C:\1\Test.docx";
        // Define an object to pass to the API for missing parameters
        object missing = System.Type.Missing;
        doc = word.Documents.Open(ref fileName,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);
        string ReadValue = string.Empty;
        // Activate the document
        doc.Activate();

         foreach (Word.Range tmpRange in doc.StoryRanges)
         {
            ReadValue += tmpRange.Text;
         }
    }
*/