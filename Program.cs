/**************************************************************************************************
Original Version: David Palfrey (4/16/2013)                                                      **
Modified: dekanfrus (10/08/19)                                                                   **
                                                                                                 **
Description: This application generates random Word documents for testing.                       **
Inputs:                                                                                          **
    InputFile (Required)    - A file containing a list of filenames without extensions.          **
    OutputDir (Required)    - Directory to save the generated files                              **
    Num (Default: 200)      - The number of files to be generated                                **
    URL (Default: wikipedia)- The URL of an HTML page to copy                                    **
                                                                                                 **
Background:                                                                                      **
While designing a CTF type infrastructure for a Cybersecurity competition, I needed              **
to create several network shares to emulate what one might find in a corporate environment.      **
I wanted to add some actual files to the shares, but didn't want to manually create them, or     **
use real documents.                                                                              **
                                                                                                 **
This project expands David Palfery's "Test Document Generator" console application               **
to accept a file containing potential filenames, a number of documents to generate,              **
a URL to copy, and an output directory.                                                          **
                                                                                                 **
Todo:                                                                                            **
Implement functionality to generate more than just Word documents                                **
Automagically create directory structure and save documents throughout                           **   
                                                                                                 **
**************************************************************************************************/
using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Net;
using System.IO;
using System.Collections.Generic;
using HtmlToOpenXml;
using CommandLine;
using CommandLine.Text;

namespace DocumentGenerator
{
    class Program
    {
        public class Options
        {
            /*************************************************************
            * Parse the command line arguments                           *
            * https://github.com/commandlineparser/commandline           *
            **************************************************************/
            [Option('i', "InputFile", Required = true, HelpText = "Input file containing list of filenames. Do NOT include extensions.")]
            public string InputFile { get; set; }

            [Option('o', "OutputDir", Required = true, HelpText = "Directory to output the documents.")]
            public string OutputDir { get; set; }

            /**********************************************************************************
            * Functionality to specify different document types has not been implemented yet. *
            ***********************************************************************************/
            [Option('e', "Extension", Hidden = true, Required = false, HelpText = "The file extension to use. (Default: docx)")]
            public string Extension { get; set; }

            [Option('n', "Num", Required = false, HelpText = "The number of files to generate. (Default: 200)")]
            public int FileNum { get; set; }

            [Option('u', "URL", Required = false, HelpText = "URL of Website to copy (Default: http://en.wikipedia.org/wiki/special:Random)")]
            public string URL { get; set; }

            [Usage(ApplicationAlias = "TestFileGenerator.exe")]
            public static IEnumerable<Example> Examples
            {
                get
                {
                    yield return new Example("Normal scenario", new Options { InputFile = "filenames.txt", OutputDir = "C:\\temp\\docs", FileNum = 200, URL = "http://en.wikipedia.org/wiki/special:Random" });
                }
            }
        }

        /*************************************************************
        * Function to get a random number to select the filename     *
        **************************************************************/
        //
        private static readonly Random random = new Random();
        //private static readonly object syncLock = new object();
        public static int RandomNumber(int max)
        {
            /*************************************
            * https://stackoverflow.com/a/768001 *
            **************************************/
            return random.Next(0, max);
        }

        static void Main(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(opts => ReadyFight(opts))
                .WithNotParsed<Options>((errs) => HandleParseError(errs));
        }
        private static int ReadyFight(Options options)
        {
            var contents = File.ReadAllLines(options.InputFile);
            var linecount = contents.Count();

            for (int i = 0; i < options.FileNum; i++)
            {
                int random = RandomNumber(linecount);
                var fileName = contents[random];
                string path = "";
                if (options.Extension != null)
                {
                    path = options.OutputDir + "\\" + fileName + options.Extension;
                }
                else
                {
                    path = options.OutputDir + "\\" + fileName + ".docx";
                }
                GetDocs(path, "http://en.wikipedia.org/wiki/special:Random");
                System.Threading.Thread.Sleep(100);
            }
            return 0;
        }

        static void HandleParseError(IEnumerable<Error> errs)
        {
            
        }
        private static string GetPageText(string url)
        {
            WebClient client = new WebClient();
            return client.DownloadString(url);
        }
        private static void GetDocs(string docName, string url)
        {
            bool isDocGenerated = true;
            using (MemoryStream generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }
                    HtmlConverter converter = new HtmlConverter(mainPart);
                    Body body = mainPart.Document.Body;
                    string source = GetPageText(url);
                    converter.ConsiderDivAsParagraph = true;
                    converter.ExcludeLinkAnchor = true;
                    converter.ImageProcessing = ImageProcessing.Ignore;
                    IList<OpenXmlCompositeElement> paragraphs = null;
                    try
                    {
                        paragraphs = converter.Parse(source);
                        for (int i = 0; i < paragraphs.Count; i++)
                        {
                            body.Append(paragraphs[i]);
                        }
                    }
                    catch
                    {
                        isDocGenerated = false;
                    }
                    mainPart.Document.Save();
                }
                if (isDocGenerated)
                {
                    if (File.Exists(docName)) File.Delete(docName);
                    File.WriteAllBytes(docName, generatedDocument.ToArray());
                }
            }
        }
    }
}