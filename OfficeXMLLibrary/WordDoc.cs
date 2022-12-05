// ConsoleApp6
// OfficeXMLLibrary
// WordDoc.cs
// ---------------------------------------------
// Alex Samarkin
// Alex
// 
// 20:33 04 12 2022

using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeXMLLibrary
{
    public class WordDoc : IOfficeDocument
    {
        public string FileName { get; set; } = "./tmp.docx";

        WordprocessingDocument wordprocessingDocument;
        Body body;
        Paragraph paragraph;
        Run run;

        public void AddParagraph(string s)
        {
            if (wordprocessingDocument == null) return;
            paragraph = body.AppendChild(new Paragraph());
            run = paragraph.AppendChild(new Run());
            run.AppendChild(new Text(s));

        }

        #region Implementation of IOfficeDocument

        public void Create(string FulllName = "")
        {
            if (FulllName!="") FileName = FulllName;

            using (wordprocessingDocument =
                   WordprocessingDocument.Create(FulllName,
                       WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordprocessingDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                body = mainPart.Document.AppendChild(new Body());
                paragraph = body.AppendChild(new Paragraph());
                run = paragraph.AppendChild(new Run());
                run.AppendChild(new Text($"Текст создан программой в {DateTime.Now}"));
            }
        }
        public void Open(string FulllName ="")
        {
            if (FulllName != "") FileName = FulllName;

            wordprocessingDocument =
                WordprocessingDocument.Open(FileName, true);
            body = wordprocessingDocument.MainDocumentPart.Document.Body;
        }
        public void Close()
        {
            if (wordprocessingDocument!=null)
            {
                wordprocessingDocument.Close();
            }
            wordprocessingDocument = null;
            body = null;
            paragraph = null;
            run = null;
        }
        #endregion
    }
}