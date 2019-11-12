using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Gluon;

namespace Gluon
{
    static class Gluon
    {
        /// <summary>
        /// Reports an error in the call.
        /// </summary>
        public static void error()
        {
            throw new ArgumentException
            ("Expected folder to be input containing other folders each with files to be concatenated into a .docx file.");
        }

        /// <summary>
        /// The main method of the program, one argument is expected containing a root folder
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {

         
            if (args.Length == 0)
            {
                string error = @"Usage: first argument should be a 'root' folder
containing other folders, where each direct subfolder's contents are recursively added
to a .docx with the name of the subfolder.";
                Console.Error.WriteLine(error);
                Environment.FailFast(error);
                return;
            }

            string path = args[0];

            if (!new DirectoryInfo(path).Exists)
            {
                error();
            }

            GlueFoldersInPath(path);
        }
        /// <summary>
        /// Glues all of the (descendant) files within each subfolder into a .docx with the name of that subfolder, for each subfolder.
        /// </summary>
        /// <param name="root">The root folder containing the other folders.</param>
        public static void GlueFoldersInPath(string root)
        {
            var dirs = Directory.GetDirectories(root);

            if(!dirs.Any())
            {
                error();
            }

            foreach (var dir in dirs)
            {
                //Create a .docx named after the directory, within the root.
                try
                { 
                Glue(dir,$"{dir}.docx");
                }
                catch(Exception ex)
                {
                    Console.Error.WriteLine($"Please check { dir}.error.txt for details regarding an error.");
                    using (var file = File.Open($"{dir}.error.txt", FileMode.OpenOrCreate))
                    using (StreamWriter writer = new StreamWriter(file))
                    {
                        writer.WriteLine(DateTime.Now.ToString());
                        writer.WriteLine(ex.GetType());
                        writer.WriteLine(ex.Message);
                        writer.WriteLine(ex.StackTrace);
                        writer.WriteLine();
                    }
                }
            }
        }



        /// <summary>
        /// Glues all .sql .txt and .docx files in the specified 'root' together into the specified filename,
        /// </summary>
        /// <param name="root"></param>
        /// <param name="originalFileName"></param>
        static void Glue(string root, string originalFileName)
        {

            string resultFileName = originalFileName;
            //var existingFiles = Directory.GetFiles(root);
            var existingFiles  = Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories);

            while (File.Exists(resultFileName))
            {
                resultFileName += "_new.docx";
            }

            using (var stream = File.Open(resultFileName, FileMode.CreateNew))
            {
                // Create Document
                using (WordprocessingDocument document =
                    WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
                {
                    var mainPart = document.AddMainDocumentPart();

                    mainPart.Document = new Document();
                    Body body = new Body();
                    mainPart.Document.Body = body;

                    body.NewParaWithRun($"Folder: \n {root}");

                    foreach (string fullName in existingFiles)
                    {
                        if (fullName == originalFileName)
                        {
                            continue;
                        }
                        string fileName = Path.GetFileName(fullName);

                        //Create a header for the file.
                        body.NewParaWithRun($"bestand {fileName}").Bold();

                        int chunkId = 2;

                        switch (Path.GetExtension(fullName).ToUpperInvariant())
                        {
                            //TODO: maybe merge .doc files? this requires some stuff sing e.g. documentbuilder stuff.
                            case ".DOCX":
                            case ".DOC":
                                body.NewParaWithRun($"bestand {fileName}").Bold().Highlight(HighlightColorValues.Yellow).Size(20);

                                string ChunkId = "A" + Guid.NewGuid().ToString();

                                //Merge documents.
                                AltChunk altChunk = new AltChunk();
                                altChunk.Id = ChunkId;

                                AlternativeFormatImportPart chunk =
                                mainPart.AddAlternativeFormatImportPart(
                                AlternativeFormatImportPartType.WordprocessingML, ChunkId);

                                using (var fs = File.Open(fullName, FileMode.Open))
                                {
                                    chunk.FeedData(fs);
                                }

                                mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>().Last());
                                //mainPart.Document.Save(); ///not sure if required here.
                                break;
                            case ".SQL":
                            case ".TXT":
                                body.NewParaWithRun(File.ReadAllText(fullName)).Size(22).Font("Consolas"); //fixed width is better for code.
                                break;
                            default:
                                body.NewParaWithRun($"Inhoud overgeslagen van: {fileName}").Highlight();
                                body.NewParaWithRun($"{fullName}").Highlight().Size(22);
                                break;
                        }
                    }
                }
            }
        }
    }
}
