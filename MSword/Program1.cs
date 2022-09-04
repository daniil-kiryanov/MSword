/*using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSword
{
    internal class Program1
    {
        static void Main(string[] args)
        {
            
            string line = "";
            using (StreamReader sr = new StreamReader(@"C:\1\11.doc"))
            {
                while ((line = sr.ReadLine()) != null)
                {
                    Console.WriteLine(line);
                    Console.ReadLine();
                }
            }
            
            StreamWriter sw;
            string FileName = @"C:\1\12.xls";
            sw = File.CreateText(FileName);
            string FileData = line;
            sw.WriteLine(FileData);
            sw.Close();
            
        }
    }
}
*/