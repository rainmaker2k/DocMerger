using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                var filelines = File.ReadLines(args[0]);
                var docxfiles = filelines.Where(s => !s.StartsWith("#") && !String.IsNullOrWhiteSpace(s) && !s.StartsWith("out:"));

                var outfilename = filelines.FirstOrDefault(s => s.StartsWith("out:")).Substring("out:".Length).Trim();
                Directory.CreateDirectory("out");
                CombineWordDocuments(docxfiles.Select(d => $"in\\{d}"), outfilename);
            }
        }

        public static void CombineWordDocuments(IEnumerable<string> paths, string outfilename)
        {
            var pathlist = paths.ToList();

            string outputpath = @$"out\{outfilename}";
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
                    if (File.Exists(path))
                    {
                        using (FileStream fileStream = File.Open(path, FileMode.Open))
                            chunk.FeedData(fileStream);
                        AltChunk altChunk = new AltChunk();
                        altChunk.Id = altChunkId;
                        mainPart.Document
                            .Body
                            .InsertAfter(altChunk, mainPart.Document.Body
                            .Elements<Paragraph>().Last());
                    }
                    count++;
                }

                mainPart.Document.Save();
            }
        }
    }
}
