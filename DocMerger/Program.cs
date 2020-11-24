using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                var docxfiles = File.ReadLines(args[0]).Where(s => !s.StartsWith("#"));
                Directory.CreateDirectory("out");
                CombineWordDocuments(docxfiles.Select(d => $"in\\{d}"));
                //var filecontents = docxfiles.Select(d => File.ReadAllBytes($"in\\{d}")).ToList();
                //var combined = OpenAndCombine(filecontents);
                //File.WriteAllBytes(@"out\combined.docx", combined);
            }
        }

        public static void CombineWordDocuments(IEnumerable<string> paths)
        {
            var pathlist = paths.ToList();

            string outputpath = @"out\combined.docx";
            File.Delete(outputpath);
            File.Copy(pathlist.First(), outputpath);

            var count = 1;
            
            using (WordprocessingDocument myDoc =
                WordprocessingDocument.Open(outputpath, true))
            {

                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                foreach (var path in pathlist.Skip(1))
                {
                    string altChunkId = $"AltChunkId{count}";
                    AlternativeFormatImportPart chunk =
                        mainPart.AddAlternativeFormatImportPart(
                        AlternativeFormatImportPartType.WordprocessingML, altChunkId);
                    using (FileStream fileStream = File.Open(path, FileMode.Open))
                        chunk.FeedData(fileStream);
                    AltChunk altChunk = new AltChunk();
                    altChunk.Id = altChunkId;
                    mainPart.Document
                        .Body
                        .InsertAfter(altChunk, mainPart.Document.Body
                        .Elements<Paragraph>().Last());
                    count++;
                }

                mainPart.Document.Save();
            }
                
            
            
        }

        public static byte[] OpenAndCombine(IList<byte[]> documents)
        {
            MemoryStream mainStream = new MemoryStream();

            mainStream.Write(documents[0], 0, documents[0].Length);
            mainStream.Position = 0;

            int pointer = 1;
            byte[] ret;
            try
            {
                using (WordprocessingDocument mainDocument = WordprocessingDocument.Open(mainStream, true))
                {

                    XElement newBody = XElement.Parse(mainDocument.MainDocumentPart.Document.Body.OuterXml);

                    for (pointer = 1; pointer < documents.Count; pointer++)
                    {
                        WordprocessingDocument tempDocument = WordprocessingDocument.Open(new MemoryStream(documents[pointer]), true);
                        XElement tempBody = XElement.Parse(tempDocument.MainDocumentPart.Document.Body.OuterXml);

                        newBody.Add(tempBody);
                        mainDocument.MainDocumentPart.Document.Body = new Body(newBody.ToString());
                        mainDocument.MainDocumentPart.Document.Save();
                        mainDocument.Package.Flush();
                    }
                }
            }
            catch (OpenXmlPackageException oxmle)
            {
                throw new Exception($"Error while merging files. Document index {pointer}", oxmle);
            }
            catch (Exception e)
            {
                throw new Exception($"Error while merging files. Document index {pointer}", e);
            }
            finally
            {
                ret = mainStream.ToArray();
                mainStream.Close();
                mainStream.Dispose();
            }
            return (ret);
        }
    }
}
