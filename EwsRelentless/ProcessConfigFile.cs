using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Security;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

namespace EwsRelentless
{
    class ProcessConfigFile
    {
        bool _SuppressLogging = false;
        bool _LogAsCSV = false;
        bool _TrustAllCerts = false;
        
        public int _ThreadCount = 0;
        public bool _BeginShutdown = false;

        TextWriterTraceListener _LogListener;

        public void ProcessFile(
                string ConfigFile, 
                bool bSuppressLogging, 
                bool bLogAsCSV, 
                TextWriterTraceListener oLogListener, 
                List<EWSConfigItems> listEWSConfigItems,
                bool bTrustAllCerts)
        {
            _SuppressLogging = bSuppressLogging;  // Suppress all output?
            _LogAsCSV = bLogAsCSV;
            _TrustAllCerts = bTrustAllCerts;
            Console.WriteLine("");
            Console.WriteLine(""); 
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("                    ---------------------------------------------");
            Console.WriteLine("                           Press any key to exit testing.");
            Console.WriteLine("                    ---------------------------------------------");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Running tests in 3...");
            Thread.Sleep(1000);
            Console.WriteLine("                 2...");
            Thread.Sleep(1000);
            Console.WriteLine("                 1...");
            Thread.Sleep(1000);
            Console.WriteLine("Started running tests...");

            if (_TrustAllCerts)
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            else
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            //ServicePointManager.UseNagleAlgorithm = true;
            //ServicePointManager.Expect100Continue = true;
            //ServicePointManager.CheckCertificateRevocationList = true;
            ServicePointManager.DefaultConnectionLimit = 65000;  //Yeah... really, really big
 
            _LogListener = oLogListener; // This will turn on logging to CSV file
            Console.WriteLine("");

            if (_LogAsCSV)
            {
                WriteCsvLine("Start Time, End Time, MS Runtime, Test Name, Test Number, Thread, Error Message");
            }

            Console.WriteLine("");

            int iCount = 0;
            foreach (EWSConfigItems oEWSConfigItems in listEWSConfigItems)
            {
                iCount++;
                WriteLog("Started processing config test " + iCount.ToString() + ".");
                 
                DoTest(oEWSConfigItems);
                 
                WriteLog("Finished processing config test " + iCount.ToString() + ".");
            }
 

            WriteLog("All requests executing.");

            WriteLog("");

            int iCharUpper = Convert.ToInt16('X');
            bool KeepWaitingForExit = true;
           
            ConsoleKeyInfo oConsoleKeyInfo;
            while (KeepWaitingForExit == true)
            {

               oConsoleKeyInfo = Console.ReadKey();
 
               if (oConsoleKeyInfo.KeyChar.ToString().ToUpper() == "X")
                    KeepWaitingForExit = false;
            }

            WriteLog("Telling threads to end after finishing current EWS calls.");
 
            _BeginShutdown = true;

            while (_ThreadCount !=0)
            {

                //Console.WriteLine("Threads: {0}", _ThreadCount);
                //Console.ReadKey();
                Thread.Sleep(500);
            }

            WriteLog("Threads shut down - Exiting.");
 
            
            _LogListener = null; // This will turn off logging to CSV file

        }


        private void WriteCsvLine(string sLine)
        {
            if (_LogAsCSV)
            {
                if (_LogListener == null)
                {
                    Console.WriteLine(sLine);
                }
                else
                {
                    _LogListener.WriteLine(sLine);
                    _LogListener.Flush();
                }
            }
        }

        private void WriteCsvLine(DateTime FromDateTime, DateTime ToDateTime, string Test, int LineNumber, string ThisThread, string ErrorMessage)
        {
 
            TimeSpan oTimeSpan = ToDateTime - FromDateTime;

            if (_LogAsCSV)
            {
                if (_LogListener == null)
                {
                    Console.WriteLine(string.Format("{0},{1},{2},{3},{4},{5},{6}",
                        GetTimeString(FromDateTime),
                        GetTimeString(ToDateTime),
                        oTimeSpan.TotalMilliseconds,
                        Test,
                        LineNumber,
                        ThisThread,
                        ErrorMessage.Replace(",", "_")
                        ));
                }
                else
                {
                    _LogListener.WriteLine(string.Format("{0},{1},{2},{3},{4},{5},{6}",
                        GetTimeString(FromDateTime),
                        GetTimeString(ToDateTime),
                        oTimeSpan.TotalMilliseconds,
                        Test,
                        LineNumber,
                        ThisThread,
                        ErrorMessage.Replace(",", "_")
                        ));
                    _LogListener.Flush();
                }
            }
 
        }


        // Everything here is done once per config file line
        public void DoTest(EWSConfigItems oEWSConfigItems)
        {
            ExchangeService service = null;

            if ((oEWSConfigItems.TestOpertation == "AutodiscoverOnly") )
            {
                System.Threading.Tasks.Task t = null;
                for (int iCount = 1; iCount <= oEWSConfigItems.NumberOfSimultaneousCalls; iCount++)
                {
                    WriteLog("Starting Thread #" + iCount.ToString() + " for test " + oEWSConfigItems.TestOpertation + " ");
                    switch (oEWSConfigItems.TestOpertation)
                    {
                        case "AutodiscoverOnly":
                            service = GetServiceObject(oEWSConfigItems);
                            t = System.Threading.Tasks.Task.Factory.StartNew(() =>
                            {
                                DoAutodiscoverOnlyForever(oEWSConfigItems);
                            });
                            _ThreadCount++;
                            break;
 
  
                    }
                }
            }
            else
            {
                // Don't share ExchangeSerice object instances accross threads
                // NO:     service = GetServiceObject(oEWSConfigItems);

                
                //if (service != null)
                //{ 

                    System.Threading.Tasks.Task t = null;
                    for (int iCount = 1; iCount <= oEWSConfigItems.NumberOfSimultaneousCalls; iCount++)
                    {
                        WriteLog("Starting Thread #" + iCount.ToString() + " for test " + oEWSConfigItems.TestOpertation + " ");
                        switch (oEWSConfigItems.TestOpertation)
                        {
                            case "ItemSearch":
                                t = System.Threading.Tasks.Task.Factory.StartNew(() =>
                                {
                                    DoFindItemsSearchForever(GetServiceObject(oEWSConfigItems), oEWSConfigItems, WellKnownFolderName.Inbox);
                                });
                                _ThreadCount++;
                                break;
                            case "CalendarSearch":
                                t = System.Threading.Tasks.Task.Factory.StartNew(() =>
                                {
                                    DoSearchCalendarForever(GetServiceObject(oEWSConfigItems), oEWSConfigItems, WellKnownFolderName.Calendar);
                                });
                                _ThreadCount++;
                                break;

                            case "ReadAllContacts":
                                t = System.Threading.Tasks.Task.Factory.StartNew(() =>
                                {
                                    DoReadAllContactsForever(GetServiceObject(oEWSConfigItems), oEWSConfigItems, WellKnownFolderName.Contacts);
                                });
                                _ThreadCount++;
                                break;

                            case "ReadMimeForInboxAllItems":
                                t = System.Threading.Tasks.Task.Factory.StartNew(() =>
                                {
                                    DoReadMimeForAllItemsForever(GetServiceObject(oEWSConfigItems), oEWSConfigItems, WellKnownFolderName.Inbox);
                                });
                                _ThreadCount++;
                                break;

                            case "ResolveRecipient":
                                t = System.Threading.Tasks.Task.Factory.StartNew(() =>
                                {
                                    DoResolveRecipientForever(GetServiceObject(oEWSConfigItems), oEWSConfigItems, WellKnownFolderName.Inbox);
                                });
                                _ThreadCount++;
                                break;
                            case "SendEmail":
                                t = System.Threading.Tasks.Task.Factory.StartNew(() =>
                                {
                                    DoSendEmailForever(GetServiceObject(oEWSConfigItems), oEWSConfigItems);
                                });
                                _ThreadCount++;
                                break;
                        }
                   // }
                }
            }
      
 
        }

        private void WriteLog(string LogLine)
        {
            if (_SuppressLogging == false)
            {
                DateTime oDateTime = DateTime.Now;
                Console.WriteLine(string.Format("{0} - {1}", GetTimeString(oDateTime).PadRight(23), LogLine));
            }
        }

        private void WriteLogError(string TestingOperation, int LineNumber, string ThisThread, double TotalMilliseconds, string ErrorMessage)
        {
 
            if (_SuppressLogging == false)
            {
                //WriteLog(string.Format("Error - DoAutodiscoverOnly (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));

                DateTime oDateTime = DateTime.Now;
                string LogLine = string.Format("Error - {0} - {1} - (Test:{2}) (Thread:{3}) (MS TimeSpan: {4}) :\r\n{5}", 
                                TestingOperation,
                                GetTimeString(oDateTime).PadRight(23),
                                LineNumber, 
                                ThisThread, 
                                TotalMilliseconds, 
                                ErrorMessage
                                );
                 
                Console.WriteLine(string.Format("{0} - {1}", GetTimeString(oDateTime).PadRight(23), LogLine));
            }
        }

        private string GetTimeString(DateTime oDateTime)
        {
            return string.Format("{0}/{1}/{2} {3}:{4}:{5}.{6}",
                    oDateTime.Month,
                    oDateTime.Day,
                    oDateTime.Year,
                    oDateTime.Hour.ToString().PadLeft(2, '0'),
                    oDateTime.Minute.ToString().PadLeft(2, '0'),
                    oDateTime.Second.ToString().PadLeft(2, '0'),
                    oDateTime.Millisecond);
        }

        public void DoAutodiscoverOnlyForever(EWSConfigItems oEWSConfigItems)
        {
            bool bTest = true;
            while (bTest == true)
            {
                DoAutodiscoverOnly(oEWSConfigItems);
                Thread.Sleep(oEWSConfigItems.MillisecondsBeforeNextCall);

                if (_BeginShutdown == true)
                    bTest = false;
            }

            _ThreadCount--;

        }

        public void DoAutodiscoverOnly(EWSConfigItems oEWSConfigItems)
        {
            string ErrorMessage = string.Empty;
            string ThisThread = Thread.CurrentThread.ManagedThreadId.ToString();
            DateTime FromDateTime = DateTime.Now;
            DateTime ToDateTime = DateTime.Now;
            TimeSpan oTimeSpan = ToDateTime - FromDateTime;
 

            ExchangeService service = null;

            //WriteLog(string.Format("Started  - SendEmail (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));
            WriteLog(string.Format("Started  - DoAutodiscoverOnly (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));

            try
            {
                service = GetServiceObject(oEWSConfigItems);

            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceResponseException exEws)
            {
                ErrorMessage = "(" + exEws.ErrorCode.ToString() + ") " + exEws.Message;
                //WriteLog(string.Format("Error - DoAutodiscoverOnly (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
                WriteLogError("DoAutodiscoverOnly", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);

            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                //ResponseCodeType  rct = (ResponseCodeType) Enum.Parse(typeof(ResponseCodeType), 2134, true);
                WriteLogError("DoAutodiscoverOnly", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);

                //WriteLog(string.Format("Error - DoAutodiscoverOnly (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ex.Message));
            }


            ToDateTime = DateTime.Now;
            oTimeSpan = ToDateTime - FromDateTime;
            WriteLog(string.Format("Finished - DoAutodiscoverOnly (Test:{0}) (Thread:{1}) (MS TimeSpan: {2})", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds));
            WriteCsvLine(FromDateTime, ToDateTime, "DoAutodiscoverOnly", oEWSConfigItems.LineNumber, ThisThread, ErrorMessage);
        }

        public void DoSendEmailForever(ExchangeService service, EWSConfigItems oEWSConfigItems)
        {
            bool bTest = true;
            while (bTest == true)
            {
                DoSendEmail(service, oEWSConfigItems);
                Thread.Sleep(oEWSConfigItems.MillisecondsBeforeNextCall);

                if (_BeginShutdown == true)
                    bTest = false;
            }

            _ThreadCount--;
        }

        public void DoSendEmail(ExchangeService service, EWSConfigItems oEWSConfigItems)
        {
            //Param1: To SMTP Address
            //Param2: Subject
            //Param3: Body
            //Param4: SendOnly - default is SaveCopyOnSend if left blank

            string ErrorMessage = string.Empty;
            string ThisThread = Thread.CurrentThread.ManagedThreadId.ToString();
            DateTime FromDateTime = DateTime.Now;
            DateTime ToDateTime = DateTime.Now;
            TimeSpan oTimeSpan = ToDateTime - FromDateTime;

            WriteLog(string.Format("Started  - SendEmail (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));

            try
            {

                EmailMessage message = new EmailMessage(service);
                message.ToRecipients.Add(oEWSConfigItems.Param1);
                message.Subject = oEWSConfigItems.Param2;
                message.Body = oEWSConfigItems.Param3;
                if (oEWSConfigItems.Param4 == "SendOnly")
                {
                    FromDateTime = DateTime.Now;
                    message.Send();
                    ToDateTime = DateTime.Now;
                }
                else
                {
                    FromDateTime = DateTime.Now;
                    message.SendAndSaveCopy();
                    ToDateTime = DateTime.Now;
                }
                message = null;
            }
            catch (Microsoft.Exchange.WebServices.Data.ServerBusyException exServerBusyException)
            {
                ErrorMessage = "ServerBusyException - " +
                        "(Error: " + exServerBusyException.ErrorCode.ToString() + ")" +
                        "(Backoff Miliseconds:  " + exServerBusyException.BackOffMilliseconds.ToString() + ") " +
                        exServerBusyException.Message;

                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("SendEmail", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            } 
            catch (Microsoft.Exchange.WebServices.Data.ServiceResponseException exEws)
            {
                ErrorMessage = "(" + exEws.ErrorCode.ToString() + ") " + exEws.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                //WriteLog(string.Format("Error - SendEmail (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
                WriteLogError("SendEmail", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);


            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                //WriteLog(string.Format("Error - SendEmail (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
                WriteLogError("SendEmail", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }


            ToDateTime = DateTime.Now;
            oTimeSpan = ToDateTime - FromDateTime;
            WriteLog(string.Format("Finished - SendEmail (Test:{0}) (Thread:{1}) (MS TimeSpan: {2})", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds));
            WriteCsvLine(FromDateTime, ToDateTime, "SendEmail", oEWSConfigItems.LineNumber, ThisThread, ErrorMessage);
        }

        public void DoResolveRecipientForever(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
            bool bTest = true;
            while (bTest == true)
            {
                DoResolveRecipient(service, oEWSConfigItems, oWellKnownFolderName);
                Thread.Sleep(oEWSConfigItems.MillisecondsBeforeNextCall);

                if (_BeginShutdown == true)
                    bTest = false;
            }

            _ThreadCount--;
        }

        public void DoResolveRecipient(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
           string ErrorMessage = string.Empty;
           string ThisThread =  Thread.CurrentThread.ManagedThreadId.ToString();
           DateTime FromDateTime = DateTime.Now;
           DateTime ToDateTime = DateTime.Now;
           TimeSpan oTimeSpan = ToDateTime - FromDateTime;

           WriteLog(string.Format("Started  - ResolveRecipient (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));

           try
           {
               NameResolutionCollection oNameResolutionCollection = null;
               ResolveNameSearchLocation oResolveNameSearchLocation = ResolveNameSearchLocation.ContactsThenDirectory;

               switch (oEWSConfigItems.Param2)
               {
                   case "ContactsThenDirectory":
                       oResolveNameSearchLocation = ResolveNameSearchLocation.ContactsThenDirectory;
                       break;
                   case "ContactsOnly":
                       oResolveNameSearchLocation = ResolveNameSearchLocation.ContactsOnly;
                       break;
                   case "DirectoryOnly":
                       oResolveNameSearchLocation = ResolveNameSearchLocation.DirectoryOnly;
                       break;
                   case "DirectoryThenContacts":
                       oResolveNameSearchLocation = ResolveNameSearchLocation.DirectoryThenContacts;
                       break;
               }

               string sUseThis = oEWSConfigItems.Param1;
               if (oEWSConfigItems.Param1 == "RANDOM")
                   sUseThis = GetRandomString();
               else
                   sUseThis = oEWSConfigItems.Param1;

               FromDateTime = DateTime.Now;
               oNameResolutionCollection = service.ResolveName(sUseThis, oResolveNameSearchLocation, true);
               ToDateTime = DateTime.Now;
           }
            catch (Microsoft.Exchange.WebServices.Data.ServerBusyException exServerBusyException)
            {
                ErrorMessage = "ServerBusyException - " +
                        "(Error: " + exServerBusyException.ErrorCode.ToString() + ")" +
                        "(Backoff Miliseconds:  " + exServerBusyException.BackOffMilliseconds.ToString() + ") " +
                        exServerBusyException.Message;

                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ResolveRecipient", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            } 
           catch (Microsoft.Exchange.WebServices.Data.ServiceResponseException exEws)
           {
               ErrorMessage = "(" + exEws.ErrorCode.ToString() + ") " + exEws.Message;
               ToDateTime = DateTime.Now;
               oTimeSpan = ToDateTime - FromDateTime;
               //WriteLog(string.Format("Error - ResolveRecipient (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
               WriteLogError("ResolveRecipient", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
           }
           catch (Exception ex)
           {
               ErrorMessage = ex.Message;
               ToDateTime = DateTime.Now;
               oTimeSpan = ToDateTime - FromDateTime;
               //WriteLog(string.Format("Error - ResolveRecipient (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
               WriteLogError("ResolveRecipient", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
           }


            ToDateTime = DateTime.Now;
            oTimeSpan = ToDateTime - FromDateTime;
            WriteLog(string.Format("Finished - ResolveRecipient (Test:{0}) (Thread:{1}) (MS TimeSpan: {2})", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds));
            WriteCsvLine(FromDateTime, ToDateTime, "ResolveRecipient", oEWSConfigItems.LineNumber, ThisThread, ErrorMessage);

        }

        private string GetRandomString()
        {
            return System.IO.Path.GetRandomFileName().Replace(".","");
        }
        

        public void DoReadMimeForAllItemsForever(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
            bool bTest = true;
            while (bTest == true)
            {
                DoReadMimeForAllItems(service, oEWSConfigItems, oWellKnownFolderName);
                Thread.Sleep(oEWSConfigItems.MillisecondsBeforeNextCall);

                if (_BeginShutdown == true)
                    bTest = false;
            }

            _ThreadCount--;
        }

        public void DoReadMimeForAllItems(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
            string ErrorMessage = string.Empty;
            string ThisThread =  Thread.CurrentThread.ManagedThreadId.ToString();
            DateTime FromDateTime = DateTime.Now;
            DateTime ToDateTime = DateTime.Now;
            TimeSpan oTimeSpan = ToDateTime - FromDateTime;
            StringBuilder xdebug = new StringBuilder();

            WriteLog(string.Format("Started  - ReadMimeForAllItems (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));
            try
            {
                xdebug.Append("1a,");
                ItemView oView = new ItemView(99999);
                xdebug.Append("1b,");
                Folder oFolder = Folder.Bind(service, oWellKnownFolderName);
                xdebug.Append("2,");
                WriteLog(string.Format("Started  - ReadMimeForAllItems - Getting Item List (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));
                FromDateTime = DateTime.Now;
               // xdebug.Append("3,");
                FindItemsResults<Microsoft.Exchange.WebServices.Data.Item> oResults = oFolder.FindItems(oView);
                //xdebug.Append("4,");
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                //xdebug.Append("5,");
                WriteLog(string.Format("End - ReadMimeForAllItems - Getting Item List (Test:{0}) (Thread:{1}) (MS TimeSpan: {2})", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds));
                //xdebug.Append("6,");
                string SomeMime = string.Empty;
                PropertySet oMimePropertySet = new PropertySet(ItemSchema.MimeContent);
                //xdebug.Append("7,");

                DateTime oInnerFromTime = DateTime.Now;
                DateTime oInnerToTime = DateTime.Now;
                TimeSpan oInnerTimeSpan = oInnerToTime - oInnerFromTime;

                //xdebug.Append("8,");

                foreach (Item o in oResults)
                {

                    WriteLog(string.Format("Started  - ReadMimeForAllItems - Getting Item MIME (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));

                    try
                    {
                        oInnerFromTime = DateTime.Now;
                        Item oItem = Item.Bind(service, o.Id, oMimePropertySet);
                        oInnerToTime = DateTime.Now;
                        //oTimeSpan = ToDateTime - FromDateTime;
                    }
                    catch (Microsoft.Exchange.WebServices.Data.ServerBusyException exServerBusyException)
                    {
                        ErrorMessage = "ServerBusyException - " +
                                "(Error: " + exServerBusyException.ErrorCode.ToString() + ")" +
                                "(Backoff Miliseconds:  " + exServerBusyException.BackOffMilliseconds.ToString() + ") " +
                                exServerBusyException.Message;

                        ToDateTime = DateTime.Now;
                        oTimeSpan = ToDateTime - FromDateTime;
                        WriteLogError("ReadMimeForAllItems", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
                    } 
                    catch (Microsoft.Exchange.WebServices.Data.ServiceResponseException exEws)
                    {
                        ErrorMessage = "(" + exEws.ErrorCode.ToString() + ") " + exEws.Message;
                        ToDateTime = DateTime.Now;
                        oInnerTimeSpan = oInnerToTime - oInnerFromTime;
                        //WriteLog(string.Format("Error - ReadMimeForAllItems - Getting Item MIME (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
                        WriteLogError("ReadMimeForAllItems", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
                    }
                    catch (Exception ex)
                    {
                        ErrorMessage = ex.Message;
                        ToDateTime = DateTime.Now;
                        oInnerTimeSpan = oInnerToTime - oInnerFromTime;
                        //WriteLog(string.Format("Error - ReadMimeForAllItems - Getting Item MIME (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
                        WriteLogError("ReadMimeForAllItems", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
                    }

                    oInnerTimeSpan = oInnerToTime - oInnerFromTime;
                    WriteLog(string.Format("End - ReadMimeForAllItems - Getting Item MIME (Test:{0}) (Thread:{1}) (MS TimeSpan: {2})", oEWSConfigItems.LineNumber, ThisThread, oInnerTimeSpan.TotalMilliseconds));
                    WriteCsvLine(oInnerFromTime, oInnerToTime, "ReadMimeForAllItems - Getting Item MIME", oEWSConfigItems.LineNumber, ThisThread, ErrorMessage);

                }

                oView = null;
                oFolder = null;
                oResults = null;
                oMimePropertySet = null;
            }
            catch (Microsoft.Exchange.WebServices.Data.ServerBusyException exServerBusyException)
            {
                ErrorMessage = "ServerBusyException - " +
                        "(Error: " + exServerBusyException.ErrorCode.ToString() + ")" +
                        "(Backoff Miliseconds:  " + exServerBusyException.BackOffMilliseconds.ToString() + ") " +
                        exServerBusyException.Message;

                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ReadMimeForAllItems", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceXmlDeserializationException exServiceXmlDeserializationException)
            {
                ErrorMessage = "ServiceXmlDeserializationException - " + exServiceXmlDeserializationException.Message;

                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ServiceXmlDeserializationException", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceXmlSerializationException exServiceXmlSerializationException)
            {
                ErrorMessage = "ServiceXmlSerializationException - " + exServiceXmlSerializationException.Message;

                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ServiceXmlSerializationException", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            } 
            catch (Microsoft.Exchange.WebServices.Data.ServiceResponseException exEws)
            {

                ErrorMessage = "(" + exEws.ErrorCode.ToString() + ") " + exEws.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                //WriteLog(string.Format("Error - ReadMimeForAllItems (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
                WriteLogError("ReadMimeForAllItems", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (Exception ex)
            {
                string xdebuginfo = xdebug.ToString();
                ErrorMessage = ex.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
               // WriteLog(string.Format("Error - ReadMimeForAllItems (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
                WriteLogError("ReadMimeForAllItems", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }

            ToDateTime = DateTime.Now;
            oTimeSpan = ToDateTime - FromDateTime;
            WriteLog(string.Format("Finished - ReadMimeForAllItems (Test:{0}) (Thread:{1}) (MS TimeSpan: {2})", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds));
            WriteCsvLine(FromDateTime, ToDateTime, "ReadMimeForAllItems", oEWSConfigItems.LineNumber, ThisThread, ErrorMessage);

        }

        public void DoReadAllContactsForever(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
            bool bTest = true;
            while (bTest == true)
            {
                DoReadAllContacts(service, oEWSConfigItems, oWellKnownFolderName);
                Thread.Sleep(oEWSConfigItems.MillisecondsBeforeNextCall);

                if (_BeginShutdown == true)
                    bTest = false;
            }

            _ThreadCount--;
        }

        public void DoReadAllContacts(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
           string ErrorMessage = string.Empty;
           DateTime FromDateTime = DateTime.Now;
           DateTime ToDateTime = DateTime.Now;
           TimeSpan oTimeSpan = ToDateTime - FromDateTime;

           string ThisThread =  Thread.CurrentThread.ManagedThreadId.ToString();

           WriteLog(string.Format("Started  - ReadAllContacts (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));
           try
           {

               PropertySet oPropSet = new PropertySet(PropertySet.FirstClassProperties);
               oPropSet.Add(ContactSchema.ItemClass);
               oPropSet.Add(ContactSchema.DisplayName);
               oPropSet.Add(ContactSchema.Department);
               oPropSet.Add(ContactSchema.Manager);

               oPropSet.Add(ContactSchema.BusinessAddressStreet);
               oPropSet.Add(ContactSchema.BusinessAddressCity);
               oPropSet.Add(ContactSchema.BusinessAddressState);
               oPropSet.Add(ContactSchema.BusinessAddressPostalCode);
               oPropSet.Add(ContactSchema.BusinessPhone);

               oPropSet.Add(ContactSchema.HomeAddressStreet);
               oPropSet.Add(ContactSchema.HomeAddressCity);
               oPropSet.Add(ContactSchema.HomeAddressState);
               oPropSet.Add(ContactSchema.HomeAddressPostalCode);
               oPropSet.Add(ContactSchema.HomePhone);


               ItemView oView = new ItemView(9999);
               Folder folder = Folder.Bind(service, oWellKnownFolderName, oPropSet);

               FromDateTime = DateTime.Now;
               FindItemsResults<Microsoft.Exchange.WebServices.Data.Item> oResults = folder.FindItems(oView);
               ToDateTime = DateTime.Now;
               oTimeSpan = ToDateTime - FromDateTime;

               oView = null;
               folder = null;
               oResults = null;
               oPropSet = null;


           }
           catch (Microsoft.Exchange.WebServices.Data.ServerBusyException exServerBusyException)
           {
               ErrorMessage = "ServerBusyException - " +
                       "(Error: " + exServerBusyException.ErrorCode.ToString() + ")" +
                       "(Backoff Miliseconds:  " + exServerBusyException.BackOffMilliseconds.ToString() + ") " +
                       exServerBusyException.Message;

               ToDateTime = DateTime.Now;
               oTimeSpan = ToDateTime - FromDateTime;
               WriteLogError("ReadAllContacts", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
           } 
           catch (Microsoft.Exchange.WebServices.Data.ServiceResponseException exEws)
           {
               ErrorMessage = "(" + exEws.ErrorCode.ToString() + ") " + exEws.Message;
               ToDateTime = DateTime.Now;
               oTimeSpan = ToDateTime - FromDateTime;
               //WriteLog(string.Format("Error - ReadAllContacts (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
               WriteLogError("ReadAllContacts", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
           }
           catch (Exception ex)
           {
               ErrorMessage = ex.Message;
               ToDateTime = DateTime.Now;
               oTimeSpan = ToDateTime - FromDateTime;
               //WriteLog(string.Format("Error - ReadAllContacts (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
               WriteLogError("ReadAllContacts", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
           }

           ToDateTime = DateTime.Now;
           oTimeSpan = ToDateTime - FromDateTime;
           WriteLog(string.Format("Finished - ReadAllContacts (Test:{0}) (Thread:{1}) (MS TimeSpan: {2})", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds));
           WriteCsvLine(FromDateTime, ToDateTime, "ReadAllContacts", oEWSConfigItems.LineNumber, ThisThread, ErrorMessage);

        }


        public void DoSearchCalendarForever(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
            bool bTest = true;
            while (bTest == true)
            {
                DoSearchCalendar(service, oEWSConfigItems, oWellKnownFolderName);
                Thread.Sleep(oEWSConfigItems.MillisecondsBeforeNextCall);

                if (_BeginShutdown == true)
                    bTest = false;
            }

            _ThreadCount--;
        }

        public void DoSearchCalendar(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
            string ErrorMessage = string.Empty;
            DateTime FromDateTime = DateTime.Now;
            DateTime ToDateTime = DateTime.Now;
            TimeSpan oTimeSpan = ToDateTime - FromDateTime;

            DateTime oStartingDateTime = DateTime.Now;
            DateTime oToFutureDateTime = DateTime.Now;
            oToFutureDateTime = oToFutureDateTime.AddMonths(3);

            string ThisThread =  Thread.CurrentThread.ManagedThreadId.ToString();

            WriteLog(string.Format("Started  - CalendarSearch (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));

            try
            {


                CalendarView oCalendarView = new CalendarView(oStartingDateTime, oToFutureDateTime);


                oCalendarView.PropertySet = new PropertySet(BasePropertySet.IdOnly,
                    AppointmentSchema.Subject,
                    AppointmentSchema.Location,
                    AppointmentSchema.Start,
                    AppointmentSchema.End,
                    AppointmentSchema.IsRecurring,
                    AppointmentSchema.AppointmentType,
                    AppointmentSchema.ItemClass
                    );

                FromDateTime = DateTime.Now;
                FindItemsResults<Appointment> findResults = service.FindAppointments(oWellKnownFolderName, oCalendarView);
                ToDateTime = DateTime.Now;

            }
            catch (Microsoft.Exchange.WebServices.Data.ServerBusyException exServerBusyException)
            {
                ErrorMessage = "ServerBusyException - " +
                        "(Error: " + exServerBusyException.ErrorCode.ToString() + ")" +
                        "(Backoff Miliseconds:  " + exServerBusyException.BackOffMilliseconds.ToString() + ") " +
                        exServerBusyException.Message;

                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("CalendarSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            } 
            catch (Microsoft.Exchange.WebServices.Data.ServiceResponseException exEws)
            {
                ErrorMessage = "(" + exEws.ErrorCode.ToString() + ") " + exEws.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                //WriteLog(string.Format("Error - CalendarSearch (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
                WriteLogError("CalendarSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                //WriteLog(string.Format("Error - CalendarSearch (Test:{0}) (Thread:{1}) (MS TimeSpan: {2}) :\r\n{3}", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage));
                WriteLogError("CalendarSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }

            ToDateTime = DateTime.Now;
            oTimeSpan = ToDateTime - FromDateTime;
            WriteLog(string.Format("Finished - CalendarSearch (Test:{0}) (Thread:{1}) (MS TimeSpan: {2})", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds));
            WriteCsvLine(FromDateTime, ToDateTime, "CalendarSearch", oEWSConfigItems.LineNumber, ThisThread, ErrorMessage);

        }

        public void DoFindItemsSearchForever(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
            bool bTest = true;
            while (bTest == true)
            {
                DoFindItemsSearch(service,oEWSConfigItems,oWellKnownFolderName);
                Thread.Sleep(oEWSConfigItems.MillisecondsBeforeNextCall);

                if (_BeginShutdown == true)
                {
                    bTest = false;
                }
            }

            _ThreadCount--;
        }
         
        public void DoFindItemsSearch(ExchangeService service, EWSConfigItems oEWSConfigItems, WellKnownFolderName oWellKnownFolderName)
        {
            string ErrorMessage = string.Empty;
            DateTime FromDateTime = DateTime.Now;
            DateTime ToDateTime = DateTime.Now;
            TimeSpan oTimeSpan = ToDateTime - FromDateTime;

            int offset = 0;
            const int pageSize = 100;
            bool MoreItems = true;

            string ThisThread =  Thread.CurrentThread.ManagedThreadId.ToString();
            WriteLog(string.Format("Started  - ItemSearch (Test:{0}) (Thread:{1}))", oEWSConfigItems.LineNumber, ThisThread));

            try
            {

                FindItemsResults<Item> oFindItemsResults = null;

                while (MoreItems)
                {
                    List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
                    ItemView oItemView = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);

                    if (oEWSConfigItems.Param3.Length == 0)
                    {
                        oItemView.PropertySet = new PropertySet(BasePropertySet.IdOnly,
                                            ItemSchema.Subject,
                                            ItemSchema.DisplayTo,
                                            ItemSchema.DisplayCc,
                                            ItemSchema.DateTimeReceived,
                                            ItemSchema.HasAttachments,
                                            ItemSchema.ItemClass
                                            );

                        //oItemView.PropertySet.Add(oExtendedPropertyDefinition);
                        oItemView.PropertySet.Add(new ExtendedPropertyDefinition(0x1000, MapiPropertyType.String)); // PR_BODY
                        oItemView.PropertySet.Add(new ExtendedPropertyDefinition(0x1035, MapiPropertyType.String)); // CdoPR_INTERNET_MESSAGE_ID 
                        oItemView.PropertySet.Add(new ExtendedPropertyDefinition(0x0C1A, MapiPropertyType.String)); // CdoPR_SENDER_NAME
                    }
                    else
                    {
                        if (oEWSConfigItems.Param3 == "IdOnly")
                        {
                            oItemView.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                        }
                    }

                    //oItemView.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
                    oItemView.Traversal = ItemTraversal.Shallow; // shallow, associated, soft deleted

                    string sUseThis = oEWSConfigItems.Param1;
                    if (oEWSConfigItems.Param1 == "RANDOM")
                        sUseThis = GetRandomString();
                    else
                        sUseThis = oEWSConfigItems.Param1;

                    PropertyDefinition oItemSchema = ItemSchema.Subject;

                    if (oEWSConfigItems.Param2.Length == 0)
                        oItemSchema = ItemSchema.Subject;
                    else
                    {
                        switch (oEWSConfigItems.Param2)
                        {
                            case "Subject":
                                oItemSchema = ItemSchema.Subject;
                                break;
                            case "Body":
                                oItemSchema = ItemSchema.Body;
                                break;
                            case "ItemClass":
                                oItemSchema = ItemSchema.ItemClass;
                                break;

                            // new ExtendedPropertyDefinition(0x0C1A, MapiPropertyType.String)
                        }


                    }

                    searchFilterCollection.Add(new SearchFilter.ContainsSubstring(oItemSchema, sUseThis));

                    SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchFilterCollection.ToArray());

                    FromDateTime = DateTime.Now;
                    oFindItemsResults = service.FindItems(oWellKnownFolderName, searchFilter, oItemView);
                    ToDateTime = DateTime.Now;

                    // Set the flag to discontinue paging.
                    if (!oFindItemsResults.MoreAvailable)
                    {
                        MoreItems = false;
                    }

                    // Update the offset if there are more items to page.
                    if (MoreItems)
                    {
                        offset += pageSize;
                    }

                }
            }

            catch (Microsoft.Exchange.WebServices.Data.ServerBusyException exServerBusyException)
            {
                ErrorMessage = "ServerBusyException - " + 
                        "(Error: " + exServerBusyException.ErrorCode.ToString() + ")" +  
                        "(Backoff Miliseconds:  " + exServerBusyException.BackOffMilliseconds.ToString() + ") " +
                        exServerBusyException.Message;

                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ItemSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            } 
            catch (Microsoft.Exchange.WebServices.Data.ServiceValidationException exServiceValidationException)
            {
                ErrorMessage = "ServiceValidationException - " + exServiceValidationException.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ItemSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            } 
            catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException exServiceRequestException)
            {
                ErrorMessage = "ServiceRequestException - " + exServiceRequestException.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ItemSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceResponseException exEws)
            {
                ErrorMessage = "ServiceResponseException - " + "(" + exEws.ErrorCode.ToString() + ") " + exEws.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                 WriteLogError("ItemSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (Microsoft.Exchange.WebServices.Data.AccountIsLockedException exAccountIsLockedException)
            {
                ErrorMessage = "AccountIsLockedException - " + "(" + exAccountIsLockedException.AccountUnlockUrl.ToString() + ") " + exAccountIsLockedException.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ItemSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceLocalException exServiceLocalException)
            {
                ErrorMessage = "ServiceLocalException - " + exServiceLocalException.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ItemSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceRemoteException exServiceRemoteException)
            {
                ErrorMessage = "ServiceRemoteException - " + exServiceRemoteException.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ItemSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (System.Net.WebException exWeb)
            {
                ErrorMessage = "WebException - " + "(" + exWeb.Status.ToString() + ") " + exWeb.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ItemSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                ToDateTime = DateTime.Now;
                oTimeSpan = ToDateTime - FromDateTime;
                WriteLogError("ItemSearch", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds, ErrorMessage);

            }
            ToDateTime = DateTime.Now;
            oTimeSpan = ToDateTime - FromDateTime;
            WriteLog(string.Format("Finished - ItemSearch (Test:{0}) (Thread:{1}) (MS TimeSpan: {2})", oEWSConfigItems.LineNumber, ThisThread, oTimeSpan.TotalMilliseconds));
            WriteCsvLine(FromDateTime, ToDateTime, "ItemSearch", oEWSConfigItems.LineNumber, ThisThread, ErrorMessage);

        }


        // Returns a new service object
        public ExchangeService GetServiceObject(EWSConfigItems oEWSConfigItems)
        {
            string ErrorMessage = string.Empty;
            ExchangeService service = null;
            ExchangeVersion oVerion = ExchangeVersion.Exchange2007_SP1;
            if (oEWSConfigItems.EWSVersion == "Exchange2007_SP1") oVerion = ExchangeVersion.Exchange2007_SP1;
            if (oEWSConfigItems.EWSVersion == "Exchange2010") oVerion = ExchangeVersion.Exchange2010;
            if (oEWSConfigItems.EWSVersion == "Exchange2010_SP1") oVerion = ExchangeVersion.Exchange2010_SP1;
            if (oEWSConfigItems.EWSVersion == "Exchange2010_SP2") oVerion = ExchangeVersion.Exchange2010_SP2;
            if (oEWSConfigItems.EWSVersion == "Exchange2013") oVerion = ExchangeVersion.Exchange2013;

            try
            {

                service = new ExchangeService(oVerion);

                service.PreAuthenticate = oEWSConfigItems.PreAuthenticate;

                if (oEWSConfigItems.Domain.Length == 0)
                {
                    service.Credentials = new NetworkCredential(oEWSConfigItems.User, oEWSConfigItems.Password);
                }
                else
                {
                    service.Credentials = new NetworkCredential(oEWSConfigItems.User, oEWSConfigItems.Password, oEWSConfigItems.Domain);
                }

                // Impersontate ?
                if (oEWSConfigItems.ImpersonationType.Length != 0)
                {
                    ConnectingIdType oConnectingIdType = Microsoft.Exchange.WebServices.Data.ConnectingIdType.SmtpAddress;
                    switch (oEWSConfigItems.ImpersonationType)
                    {
                        case "SmtpAddress":
                            oConnectingIdType = ConnectingIdType.SmtpAddress;
                            break;
                        case "PrincipalName":
                            oConnectingIdType = ConnectingIdType.PrincipalName;
                            break;
                        case "SID":
                            oConnectingIdType = ConnectingIdType.SID;
                            break;
                    }
                    ImpersonatedUserId oImpersonatedUserId = new ImpersonatedUserId(oConnectingIdType, oEWSConfigItems.ImpersonationId);
                }

                if (oEWSConfigItems.Discovery == "DirectUrl")  // Need to autodiscover?
                {
                    WriteLog("Url Specified (" + oEWSConfigItems.LineNumber + ") - Path: " + oEWSConfigItems.CasURL);
                    service.Url = new Uri(oEWSConfigItems.CasURL);
                }
                else
                {
                    if (oEWSConfigItems.Discovery == "AudodiscoverNoSCP")
                        service.EnableScpLookup = false;  // AudodiscoverNoSCP
                    else
                        service.EnableScpLookup = true;   // Audodiscover

                    WriteLog("Starting Autodiscover (" + oEWSConfigItems.LineNumber + ").");
                    service.AutodiscoverUrl(oEWSConfigItems.UserSmtp, RedirectionUrlValidationCallback);
                    WriteLog("Finished Autodiscover (" + oEWSConfigItems.LineNumber + ") - Path: " + service.Url.AbsolutePath);

                }
            }
            catch (Microsoft.Exchange.WebServices.Data.AutodiscoverLocalException exAutoDisc)
            {
                ErrorMessage = exAutoDisc.Message;
                WriteLog("Error - Autodiscover failed - Error establishing service connection object (" + oEWSConfigItems.LineNumber + ")." + ErrorMessage + "  \r\n" + DateTime.Now.ToString());
            }
            catch (System.Net.WebException exWeb)
            {
                ErrorMessage = "(" + exWeb.Status.ToString() + ") " + exWeb.Message;
                WriteLog("Error - Autodiscover failed - Error establishing service connection object (" + oEWSConfigItems.LineNumber + ")." + ErrorMessage + "  \r\n" + DateTime.Now.ToString());
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceResponseException exEws)
            {
                ErrorMessage = "(" + exEws.ErrorCode.ToString() + ") " + exEws.Message;
                WriteLog("Error - Autodiscover failed - Error establishing service connection object (" + oEWSConfigItems.LineNumber + ")." + ErrorMessage + "  \r\n" + DateTime.Now.ToString());

            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                WriteLog("Error - Autodiscover failed - Error establishing service connection object (" + oEWSConfigItems.LineNumber + ")." + ErrorMessage + "  \r\n" + DateTime.Now.ToString());
            }

            //string sHeader = string.Format("Test{0}", oEWSConfigItems.LineNumber);
            //service.HttpHeaders.Add("EwsRelentless", "");
            service.UserAgent = string.Format("EwsRelentless-Test{0}", oEWSConfigItems.LineNumber);
           
            return service;
        }

        // http://msdn.microsoft.com/en-us/library/gg194011(v=exchg.140).aspx
        static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }

            return result;
        }


        private  bool CertificateValidationCallBack(
         object sender,
         System.Security.Cryptography.X509Certificates.X509Certificate certificate,
         System.Security.Cryptography.X509Certificates.X509Chain chain,
         System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // WriteLog("Validating the certificate.");

            if (_TrustAllCerts == true)
            {
                return true;
            }

            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {

                //WriteLog("No certificate errors found.");
                return true;
            }

            // If thre are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                //WriteLog("No certificate errors found in chain.");
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certifcates. These certificates are valid
                // for default Exchange server installations, so return true.
               // WriteLog("No certificate errors found in chain.");
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }
    }

    public class EWSConfigItems
    {
        public int LineNumber = 0;
        public string User = "";
        public string Password = "";
        public string Domain = "";  // Dont enter domain for 365

        public string ImpersonationType = string.Empty; // SMTPAddress, PricipalName, SID (New for 1.1)
        public string ImpersonationId = string.Empty; // (New for 1.1)
 
        public string Discovery = "";  // Autodiscover, AudodiscoverNoSCP, DirectUrl (New for 1.1)
        public string CasURL = "";  // if no cas URL then it will autodiscover
        public string EWSVersion = ""; //Exchange2007_SP1, Exchange2010, Exchange2010_SP1, Exchange2010_SP2, Exchange2013
        public string UserSmtp = "";   // User's mailbox smtp address
        public bool PreAuthenticate = false;  // True or False

        public int NumberOfSimultaneousCalls = 0; // number of calls running at one time
        public int MillisecondsBeforeNextCall = 0;
        public string TestOpertation = "";  // SearchInbox, SearchCalendar
        public string Param1 = "";
        public string Param2 = "";
        public string Param3 = "";
        public string Param4 = "";
 
    }

 
}
