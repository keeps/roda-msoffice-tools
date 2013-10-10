using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

using Microsoft.Office.Interop;

namespace pt.gov.dgarq.roda.common.xls2pdf
{
    class Xls2Pdf
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
                Console.WriteLine(typeof(Xls2Pdf).Name + " input_xls output_pdf");
                Console.WriteLine("\t--help Show usage.");
                //Console.WriteLine("\t-v verbose mode.");
                Console.WriteLine("\t--version Show version.");
                System.Environment.Exit(0);
            }
            if (showVersion)
            {
                System.Version xls2pdfVersion = Assembly.GetExecutingAssembly().GetName().Version;
                System.Version excelVersion = Assembly.GetAssembly(typeof(Microsoft.Office.Interop.Excel.Application)).GetName().Version;
                Console.WriteLine(typeof(Xls2Pdf).Name + " " + xls2pdfVersion + " - Microsoft Excel " + excelVersion);

                System.Environment.Exit(0);
            }
            else
            {
                try
                {

                    string originalDocFile = Path.GetFullPath(args[0]);
                    object convertedPdfFile = Path.GetFullPath(args[1]);

                    if (verbose)
                    {
                        Console.WriteLine("Creating Excel application");
                    }

                    var excelApplication = new Microsoft.Office.Interop.Excel.Application();

                    if (verbose)
                    {
                        Console.WriteLine(string.Format("Excel application, version: {0} created",
                            excelApplication.Version
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
                        Console.WriteLine("Opening XLS file " + originalDocFile);
                    }


                    Workbook workBook = excelApplication.Workbooks.Open(
                        originalDocFile,
                        true,
                        true,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing
                    );

                    if (verbose)
                    {
                        Console.WriteLine("Saving PDF file " + convertedPdfFile);
                    }

                    workBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
                        convertedPdfFile,
                        XlFixedFormatQuality.xlQualityStandard,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing,
                        missing
                        );

                    object saveChanges = false;

                    if (verbose)
                    {
                        Console.WriteLine("Closing worksheet");
                    }
                    workBook.Close(false, missing, missing);

                    if (verbose)
                    {
                        Console.WriteLine("Closing Excel application");
                    }
                    excelApplication.Quit();

                    excelApplication = null;
                    workBook = null;
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
