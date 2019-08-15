using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;

namespace EwsRelentless
{
    public class FileIoHelper
    {

        public static List<EWSConfigItems> LoadConfigFile(string ConfigFile)
        {
            List<EWSConfigItems> listEWSConfigItems = new List<EWSConfigItems>(); 

            EWSConfigItems oEWSConfigItems = null;
            bool bError = false;

            StreamReader reader = new StreamReader(File.OpenRead(ConfigFile));

            string ConfigLine = string.Empty;
            int iLine = 0;

            while (!reader.EndOfStream)
            {
                 
                ConfigLine = reader.ReadLine();
               

                if (ConfigLine.Trim().Length != 0)
                {
                     

                    if (iLine != 0) // Skip header.
                    {

                        oEWSConfigItems = new EWSConfigItems();

                        string[] values = ConfigLine.Split(',');
                        int NumberOfParams = values.Length;

                        if (NumberOfParams == 17)
                        {
                            oEWSConfigItems.LineNumber = iLine;
                            oEWSConfigItems.User = values[0].Trim();
                            oEWSConfigItems.Password = values[1].Trim();
                            oEWSConfigItems.Domain = values[2].Trim();
                            oEWSConfigItems.ImpersonationType = values[3].Trim();
                            oEWSConfigItems.ImpersonationId = values[4].Trim();

                            oEWSConfigItems.Discovery = values[5].Trim();

                            oEWSConfigItems.CasURL = values[6].Trim();
                            oEWSConfigItems.EWSVersion = values[7].Trim();
                            oEWSConfigItems.UserSmtp = values[8].Trim();

                            oEWSConfigItems.PreAuthenticate = false;
                            if (values[9].Trim().ToLower() == "true")
                                oEWSConfigItems.PreAuthenticate = true;
                            else
                                oEWSConfigItems.PreAuthenticate = false;

                            oEWSConfigItems.NumberOfSimultaneousCalls = Int32.Parse(values[10].Trim());
                            oEWSConfigItems.MillisecondsBeforeNextCall = Int32.Parse(values[11].Trim());

                            oEWSConfigItems.TestOpertation = values[12].Trim();
                            oEWSConfigItems.Param1 = values[13].Trim();
                            oEWSConfigItems.Param2 = values[14].Trim();
                            oEWSConfigItems.Param3 = values[15].Trim();
                            oEWSConfigItems.Param4 = values[16].Trim();
                            listEWSConfigItems.Add(oEWSConfigItems);


                        }
                        else
                        {
                            Console.WriteLine("****");
                            Console.WriteLine("Config file error - Line: " + (iLine -1).ToString() + " Error:  Incorrect number parameters.");
                            Console.WriteLine("****");
                            //Thread.Sleep(2000);
                            bError = true;
                        }
                    }
                    iLine++;
                }
            }

            if (bError == false)
                return listEWSConfigItems;
            else
                return null;  // null will signal calling method to exit the app.
        }


    }

 
}
