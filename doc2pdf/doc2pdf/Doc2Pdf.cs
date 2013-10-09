using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;

using Microsoft.Office.Interop.Word;

namespace pt.gov.dgarq.roda.common.doc2pdf
{
    class Doc2Pdf
    {
        static private bool verbose = false;
        static private bool showVersion = false;
        static private bool showUsage = false;

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                showUsage = true;
            }

            foreach (string arg in args) 
            {
                if ("--help".Equals(arg))
                {
                    showUsage = true;
                }

                if ("--version".Equals(arg))
                {
                    showVersion = true;
                }
            }

            if (showUsage)
            {
                Console.WriteLine(typeof(Doc2Pdf).Name + " input_doc output_pdf");
                Console.WriteLine("\t--help Show usage.");
                //Console.WriteLine("\t-v verbose mode.");
                Console.WriteLine("\t--version Show version.");
                System.Environment.Exit(0);
            }
            if (showVersion)
            {
                System.Version doc2pdfVersion = Assembly.GetExecutingAssembly().GetName().Version;
                System.Version wordVersion = Assembly.GetAssembly(typeof(ApplicationClass)).GetName().Version;
                Console.WriteLine(typeof(Doc2Pdf).Name + " " + doc2pdfVersion + " - Microsoft Word " + wordVersion);

                System.Environment.Exit(0);
            }
            else
            {
                try
                {

                    object originalDocFile = Path.GetFullPath(args[0]);
                    object convertedPdfFile = Path.GetFullPath(args[1]);

                    if (verbose)
                    {
                        Console.WriteLine("Creating Word application");
                    }

                    ApplicationClass wordApplication = new ApplicationClass();

                    object missing = System.Reflection.Missing.Value;

                    object confirmConversions = missing;
                    object readOnly = true;
                    object addToRecentFiles = missing;
                    object passwordDocument = missing;
                    object passwordTemplate = missing;
                    object revert = missing;
                    object writePasswordDocument = missing;
                    object writePasswordTemplate = missing;
                    //object format = WdOpenFormat.wdOpenFormatAuto;
                    object format = missing;
                    object encoding = missing;
                    object visible = missing;
                    object openAndRepair = missing;
                    object documentDirection = missing;
                    object noEncodingDialog = missing;
                    object xmlTranform = missing;

                    if (verbose)
                    {
                        Console.WriteLine("Opening DOC file " + originalDocFile);
                    }

                    Document document = wordApplication.Documents.Open(
                        ref originalDocFile, ref confirmConversions, ref readOnly,
                        ref addToRecentFiles, ref passwordDocument, ref passwordTemplate,
                        ref revert, ref writePasswordDocument, ref writePasswordTemplate,
                        ref format, ref encoding, ref visible,
                        ref openAndRepair, ref documentDirection, ref noEncodingDialog, ref xmlTranform);

                    //ref fileName, ref missing, ref readOnly, ref falsevalue, ref missing, ref missing, ref missing, ref missing, ref missing, ref docType, ref missing, ref isVisible

                    object fileFormat = WdSaveFormat.wdFormatPDF;
                    object embedTrueTypeFonts = true;

                    if (verbose)
                    {
                        Console.WriteLine("Saving PDF file " + convertedPdfFile);
                    }

                    document.SaveAs(
                        ref convertedPdfFile, ref fileFormat, ref missing,
                        ref missing, ref missing, ref missing,
                        ref missing, ref embedTrueTypeFonts, ref missing,
                        ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing);

                    object saveChanges = false;

                    if (verbose)
                    {
                        Console.WriteLine("Closing document");
                    }
                    document.Close(ref saveChanges, ref missing, ref missing);
                    
                    if (verbose)
                    {
                        Console.WriteLine("Closing Word application");
                    }
                    wordApplication.Quit(ref saveChanges, ref missing, ref missing);
                
                } catch (Exception e) {
                    if (verbose) 
                    {
                        Console.Error.WriteLine("ERROR - " + e.Message);
                    }

                    System.Environment.Exit(1);
                }
            }
        }
    }
}
