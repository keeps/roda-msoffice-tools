using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.PowerPoint;

using Microsoft.Office.Interop;
using Microsoft.Office.Core;

namespace pt.gov.dgarq.roda.common.ppt2pdf
{
    class Ppt2Pdf
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
                Console.WriteLine(typeof(Ppt2Pdf).Name + " input_ppt output_pdf");
                Console.WriteLine("\t--help Show usage.");
                //Console.WriteLine("\t-v verbose mode.");
                Console.WriteLine("\t--version Show version.");
                System.Environment.Exit(0);
            }
            if (showVersion)
            {
                System.Version ppt2pdfVersion = Assembly.GetExecutingAssembly().GetName().Version;
                System.Version excelVersion = Assembly.GetAssembly(typeof(Microsoft.Office.Interop.PowerPoint.Application)).GetName().Version;
                Console.WriteLine(typeof(Ppt2Pdf).Name + " " + ppt2pdfVersion + " - Microsoft PowerPoint " + excelVersion);

                System.Environment.Exit(0);
            }
            else
            {
                try
                {

                    string originalPptFile = Path.GetFullPath(args[0]);
                    string convertedPdfFile = Path.GetFullPath(args[1]);

                    if (verbose)
                    {
                        Console.WriteLine("Creating PowerPoint application");
                    }

                    var pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();

                    if (verbose)
                    {
                        Console.WriteLine(string.Format("PowerPoint application, version: {0} created",
                            pptApplication.Version
                            )
                        );
                    }

                    

                    object missing = System.Reflection.Missing.Value;

                    object confirmConversions = missing;
                    object readOnly = true;
                    object addToRecentFiles = missing;
                    object passwordDocument = missing;
                    object passwordTemplate = missing;
                    object revert = missing;
                    object writePasswordDocument = missing;
                    object writePasswordTemplate = missing;
                    object format = missing;
                    object encoding = missing;
                    object visible = missing;
                    object openAndRepair = missing;
                    object documentDirection = missing;
                    object noEncodingDialog = missing;
                    object xmlTranform = missing;

                    if (verbose)
                    {
                        Console.WriteLine("Opening PPT file " + originalPptFile);
                    }


                    var presentation = pptApplication.Presentations.Open(
                        originalPptFile,
                        MsoTriState.msoTrue,
                        MsoTriState.msoFalse,
                        MsoTriState.msoFalse
                    );

                    if (verbose)
                    {
                        Console.WriteLine("Saving PDF file " + convertedPdfFile);
                    }

                    presentation.ExportAsFixedFormat(
                        convertedPdfFile,
                        PpFixedFormatType.ppFixedFormatTypePDF,
                        PpFixedFormatIntent.ppFixedFormatIntentScreen,
                        MsoTriState.msoFalse,
                        PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                        PpPrintOutputType.ppPrintOutputSlides,
                        MsoTriState.msoFalse,
                        null,
                        PpPrintRangeType.ppPrintAll,
                        "",
                        false,
                        true,
                        true,
                        true,
                        false,
                        Type.Missing
                        );

                    object saveChanges = false;

                    if (verbose)
                    {
                        Console.WriteLine("Closing presentation");
                    }
                    presentation.Close();

                    if (verbose)
                    {
                        Console.WriteLine("Closing PowerPoint application");
                    }
                    pptApplication.Quit();

                    pptApplication = null;
                    presentation = null;
                }
                catch (Exception ex)
                {
                    if (verbose)
                    {
                        Console.Error.WriteLine("ERROR - " + ex.Message);
                    }

                    System.Environment.Exit(1);
                }
                finally
                {
                    System.GC.WaitForPendingFinalizers();
                    System.GC.Collect();
                }
            }
        }
    }
}
