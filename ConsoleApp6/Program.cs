using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeXMLLibrary;

namespace ConsoleApp6
{
    internal class Program
    {
        static void Main(string[] args)
        {
            WordDoc doc = new WordDoc();
            doc.Create("./test.docx");
            
            doc.Open();
            doc.AddParagraph("Hello, World!");
            
            Random rnd = new Random();
            
            for (int i = 0; i < 5000; i++)
            {
                doc.AddParagraph($"{i};{rnd.Next(100)};{rnd.NextDouble()}");
            }
            doc.Close();

            PowerPointDoc ppt = new PowerPointDoc();
            ppt.Create("./test.pptx");
            ppt.InsertNewSlide("./test.pptx",0,"Title");
            ppt.Close();
        }
    }
}
