using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace EwsRelentless
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.i
        /// </summary>
        [STAThread]
 

        public static void Main(string[] args)
        {
 
            bool bSuppressLogging = false;
            bool bLogAsCSV = false;
            string LogFileName = string.Empty;
            bool bError = false;
            bool bTrustAllCerts = false;
            string sFile = string.Empty;

            if (args.Length == 0)
            {
                Console.WriteLine("Configuration file must be specified.");
                bError = true;
            }

            // Check config file
            if (bError == false)
            {
                int iFound = 0;

                if (args[0].StartsWith("/"))
                {
                    Console.WriteLine("Configuration file must the first parameter.");
                    bError = true;
                }
                else
                {
                    if (File.Exists(args[0]) == false)
                    {
                        Console.WriteLine("The configuration file is missing: {0}", args[0]);
                        bError = true;
                    }
                    else
                    {
                        sFile = args[0];
                    }
 
                }

                // Checks around logging to csv format
                if (bError == false)
                {
                    iFound = FindParam(args, "/LOGASCSV");
                    if (iFound != 0)
                    {
                        bSuppressLogging = true;
                        bLogAsCSV = true;
                        int iExpectedLogFile = iFound + 1;

                        if (args.Length >= iFound + 1)
                        {
                            if (args[iExpectedLogFile].StartsWith("/"))
                            {
                                Console.WriteLine("Error - The log file name needs to be specified since /LOGASCSV is in the command line.");
                            }
                            else 
                            {
                                if (File.Exists(args[iExpectedLogFile]) == true)
                                {
                                    Console.WriteLine("Error - The file specified log file already exists: {0}", args[iExpectedLogFile]);
                                    bError = true;
                                }
                                else
                                {
                                    LogFileName = args[iExpectedLogFile];
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error - The log file name needs to be specified since /LOGASCSV is in the command line.");
                        }
                    }

                }

                // Check for singlton switches
                if (bError == false)
                {
                    iFound = FindParam(args, "/SILENT");
                    if (iFound != 0)
                        bSuppressLogging = true;

                    iFound = FindParam(args, "/TRUSTALLCERTS");
                    if (iFound != 0)
                        bTrustAllCerts = true;
                }
 

                //if (args.Length >= 1 || args.Length <= 4)
                //{
                //    string sFile = args[0];

                //    if (args.Length >= 2)
                //    {
                //        if (args[1].ToUpper() == "/SILENT") // SILENT
                //        {
                //            bSuppressLogging = true;
                //        }
                //        else
                //        {
                //            if (args[1].ToUpper() == "/LOGASCSV") // LogAsCSV
                //            {
                //                bSuppressLogging = true;
                //                bLogAsCSV = true;

                //                if (args.Length == 3 && (args[2].StartsWith("/") == false))  // Have a filename?
                //                {
                //                    if (args[2].StartsWith("/") == false)
                //                    {
 
                //                        if (File.Exists(args[2]))
                //                        {
                //                            Console.WriteLine("Error The file specified log file already exists: {0}", args[2]);
                //                            bError = true;
                //                        }
                //                        else
                //                        {

                //                            LogFileName = args[2];

                //                            if (args[3].ToUpper() == "/TRUSTALLCERTS") // TRUSTALLCERTS
                //                            {
                //                                bTrustAllCerts = true;
                //                            }
                //                            else
                //                            {
                //                                Console.WriteLine("Unknown Parameter: " + args[3]);
                //                                bError = true;
                //                            }

                //                        }
                //                    }
                //                }
                //            }
                //            else
                //            {
                //                if (args[1].ToUpper() == "/TRUSTALLCERTS") // TrustAllCerts
                //                {
                //                    bTrustAllCerts = true;
                //                }
                //                else
                //                {
                //                    Console.WriteLine("Unknown Parameter: " + args[1]);
                //                    bError = true;
                //                }
                //            }
                //        }
                //    }

                    if (bError == false)
                    {

                        if (File.Exists(sFile) == true && bError == false) // Verify that the config file exists 
                        {
                            Console.WriteLine("Using configuration file: " + sFile);

                            List<EWSConfigItems> listEWSConfigItems = new List<EWSConfigItems>();
                            listEWSConfigItems = FileIoHelper.LoadConfigFile(sFile);

                            TextWriterTraceListener oLogListener = null;

                            if (listEWSConfigItems != null)
                            {

                                if (LogFileName != string.Empty)
                                {
                                    oLogListener = new TextWriterTraceListener(LogFileName, "LogListener");

                                }
 
                                ProcessConfigFile oPCF = new ProcessConfigFile();
                                oPCF.ProcessFile(sFile, bSuppressLogging, bLogAsCSV, oLogListener, listEWSConfigItems, bTrustAllCerts);
 
                                if (LogFileName != string.Empty)
                                    oLogListener.Flush();

        
                            }
                            else
                            {
                                Console.WriteLine("");
                                Console.WriteLine("Error occurred in loading configuration file. Exiting application.");
                                System.Threading.Thread.Sleep(3000);
                            }
                        }
                        else
                        {
                            Console.WriteLine("The configuration file does not exist.");
                        }
                    }
                     
                //}
 
            }
 
        }


        private static int FindParam(string[] args, string sParam)
        {
            int iFound = 0;
            for (int iCount = 0; iCount == args.Length -1; iCount++)
            {
                if (args[iCount] == sParam)
                {
                    iFound = iCount;
                }
            }

            return iFound;
        }
    }
}
