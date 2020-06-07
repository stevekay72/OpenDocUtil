using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenDocUtil
{
    class Program
    {
        static void Main(string[] args)
        {
            string testFile = "/Users/stevekay72/Projects/OpenDocUtil/OpenDocUtil/TestDoc001.docx";
            using (var doc = WordprocessingDocument.Open(testFile, true))
            {
                var mainDocPart = doc.MainDocumentPart;
                var links = mainDocPart.HyperlinkRelationships.Where(x => x.Id != null).ToList();
                foreach(var lnk in links)
                {
                    var relationId = lnk.Id;
                    var uri = lnk.Uri.ToString();
                    mainDocPart.DeleteReferenceRelationship(lnk);
                    Uri newUri = new Uri(uri.Replace("google", "microsoft"));
                    mainDocPart.AddHyperlinkRelationship(newUri, true, relationId);
                    Console.WriteLine($"{uri} => {newUri.ToString()}");
                }
                doc.SaveAs(testFile.Replace("001","002"));
            }
        }
    }
}
