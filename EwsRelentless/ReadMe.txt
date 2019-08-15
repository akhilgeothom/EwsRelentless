EwsRelentless 1.0   4/24/2013

This is a sample application which demonstrates you might place a heavy load of EWS calls against an Exchange server in order 
to test performance. It is intended to be used in a lab and for educational purposes only. Used incorrectly it can generate 
enough traffic to effectively take down an Exchange server by generating DOS-Attack levels of calls. It is not intended to be
used in a production environment.  By using the code or binary of this application, you take responsibility for its usage.
 
This is a console application which works off of a CSV format configuration file being passed to it as a command line parameter.  
The config file is central in establishing what type and amount of load generated.  The config file controls what tests are done and
how many instances of that test are performed at the same time. You can have multiple instances of the same test running
at the same time by adjusting the NumberOfSimultaneousCalls setting, which controls the number of threads the test will be performed on. 
When a test on a thread completes, the thread will pause for the amount of time specified by the MillisecondsBeforeNextCall setting and 
then the test will repete on the same thread. 

Output is directed to the screen in a columnar fashion by default. The application can also write its output as a CSV file. The generated 
CSV file can be loaded into Excel for review.

The first line in the config file is a header file and is required.  All blank character spaces will be stripped from the log file when 
its read. Be sure to not include commas in the data used in the file since commas are field delimiters. Blank lines will be skipped.

Both Autodiscover and direct URL reference are supported. Impersonation is currently not supported in this version.

Usage:  
	EwsRelentless <ConfigFilename>
	EwsRelentless <ConfigFilename> /Silent
	EwsRelentless <ConfigFilename> /LogAsCSV  
	EwsRelentless <ConfigFilename> /LogAsCSV <CsvLogFile>
	EwsRelentless <ConfigFilename> /LogAsCSV <CsvLogFile> /TrustAllCerts


Usage examples: 
	EwsRelentless TestEwsConfig.txt
	EwsRelentless "c:\mytest\TestEWSConfig2.txt" /LogAsCSV 
	EwsRelentless "c:\mytest\TestEWSConfig2.txt" /LogAsCSV "c:\mytest\text.csv"

Parameters:
	<ConfigFilename> is a configuration file you must create.  Please refer to the information below.  There is also a sample config
	file included.  The first line in the file is a header file and will be skipped upon processing.

	The /Silent parameter will supress console output with the exception of errors in the initial launching of the application.

	The /LogAsCSV will log summeryinfo on the calls in CSV format. Commas which might appear in error messages will appear as
	underscore (_) characters. Normal logging for each request will be surpressed.

	CsvLogFile is the file that all CSV content will be written.

Below is a break-down of each field in the config file.  Each field must be seperated by a comma. All fields positions are 
required, All fields with the exception of the CasURL and parameter fields are required for each config file line; however, some tests may 
require certain parameters. The header file is required and will be skipped while processing.
 
        User              - User alias or UPN/Smpt.  Use latter for 365.
        Password          - User password
        Domain            - The domain of the user account. Dont enter the domain for 365/BPOS.
        ImpersonationType - Impersonation is done if this is field has a value.  It should be set to one of the following:
		                        SMTPAddress, PricipalName, SID  
		ImpersonationId   - ID to user for Impersonation.  The type of ID is designated by ImpersonationType		
		Discovery         - The way the CAS URL for the mailbox is found:
		                        Autodiscover
								AudodiscoverNoSCP
								DirectUrl
							With the exception of the Autodicover specific tests, autodiscover is done only one time
		                    per test and the resulting connection object is reused on each testing thread created for that test.
		CasURL            - Use this to specify the CAS URL of the mailbox directly. 
							Discovery should be set to "DirectUrl" when the this field is used.
        EWSVersion        - Version of Exchange your testing against: Exchange2007_SP1, Exchange2010, Exchange2010_SP1, Exchange2010_SP2, Exchange2013
        UserSmtp          - The SMTP address of the mailbox.
		PreAuthenticate   - Controls pre-authentication. The initial anonymouse posts which are normally done if this 
				          value is set to "True".  Set to "False" to include the initial anonymous posts.
		NumberOfSimultaneousCalls -  Number of calls running at one time.  A seperate thread will be launched for
		                  the number of calls specified here and the specified TestOperation will repeat on the same thread 
				          after the prior one completes.  This value needs to be at least one.
        MillisecondsBeforeNextCall  - Number of milliseconds to sleep before the test is repeated on the testing thread.
		                  1 second is 1000 milliseconds.  So, 3 seconds is 3000 milliseconds.

        TestOpertation -  Test operation to perform:  SearchInbox, SearchCalendar
        Param1 - First parameter to be used with the testing operation specified by TestOpertation. This value
				 varries depending upon the testing operation.
        Param2 - Second parameter to be used with the testing operation specified by TestOpertation. This value
				 varries depending upon the testing operation.
		Param3 - Third parameter to be used with the testing operation specified by TestOpertation. This value
				 varries depending upon the testing operation.
		Param4 - Fourth parameter to be used with the testing operation specified by TestOpertation. This value
				 varries depending upon the testing operation.

Example test file:
SMTPAddress, PricipalName, SID

User,             Password,    Domain,         ImpersonationType, ImpersonationId,     Discovery,         CasURL,                                EWSVersion,       UserSmtp,                    PreAuthenticate,   NumberOfSimultaneousCalls, MillisecondsBeforeNextCall, TestOpertation,             Param1,                          Param2,                       Param3,                      Param4
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    Autodiscover,      ,                                      Exchange2010_SP1, administrator@microsoft.com,                ,   1,                         1,                          ReadMimeForInboxAllItems,   ,                                ,                             ,
administrator,    MyPassword!, danzx2010.com,  SMTPAddress,       bob@danzx2010.com,   AudodiscoverNoSCP, ,                                      Exchange2010_SP2, administrator@microsoft.com,                ,   1,                         1,                          ReadMimeForInboxAllItems,   ,                                ,                             ,
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP1, administrator@microsoft.com,                ,   10,                        1,                          ItemSearch,                 Other Things,                    ,                             IdOnly,
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP2, administrator@microsoft.com,                ,    1,                         1,                          ItemSearch,                 Mr. Jones,                       Body,                         ,
          bob,    MyPassword!, contoso,        ,                  ,                    ,                  http://127.0.0.1/ews/exchange.asmx,    Exchange2010_SP1, bob@contoso.com,                            ,    12,                         2000,                       ItemSearch,                 Something,                       ,                             ,
          bob,    MyPassword!, contoso,        ,                  ,                    ,                  http://127.0.0.1/ews/exchange.asmx,    Exchange2010_SP1, bob@contoso.com,                            ,    12,                         1,                          CalendarSearch,             ,                                ,                             ,
          bob,    MyPassword!, contoso,        ,                  ,                    ,                  http://127.0.0.1/ews/exchange.asmx,    Exchange2010_SP1, bob@contoso.com,                            ,    17,                         1,                          ItemSearch,                 Other Things,                    ,                             ,
          bob,    MyPassword!, contoso,        ,                  ,                    ,                  http://127.0.0.1/ews/exchange.asmx,    Exchange2010_SP1, bob@contoso.com,                            ,    12,                         2000,                       ReadAllContacts,            ,                                ,                             ,
          bob,    MyPassword!, contoso,        ,                  ,                    ,                  http://127.0.0.1/ews/exchange.asmx,    Exchange2010_SP2, bob@contoso.com,                            ,    12,                         2000,                       ReadMimeForInboxAllItems,   ,                                ,                             ,
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP1, administrator@microsoft.com,                ,    11,                         2000,                       ItemSearch,                 RANDOM,                          Body,                         ,
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP1, administrator@microsoft.com,                ,    11,                         2000,                       ItemSearch,                 test,                            Body,                         ,
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP1, administrator@microsoft.com,                ,    11,                         5000,                       ItemSearch,                 IPM.Note,						  ItemClass,                   ,
          joe,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP1, joe@microsoft.com,                          ,    13,                         2000,                       ResolveRecipient,           Smith,                           ContactsThenDirectory,        ,
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP1, administrator@microsoft.com,                ,    11,                         2000,                       ResolveRecipient,           admin,                           DirectoryOnly,                ,
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP2, administrator@microsoft.com,                ,    11,                         1000,                       ResolveRecipient,           RANDOM,                          ,                             ,
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP1, administrator@microsoft.com,                ,    11,                         2000,                       SendEmail,                  administrator@danzx2010.com,     Test Subject,                 This is my body text., 
administrator,    MyPassword!, danzx2010.com,  ,                  ,                    ,                  http://65.53.2.65/ews/exchange.asmx,   Exchange2010_SP1, administrator@microsoft.com,                ,    11,                         2000,                       SendEmail,                  administrator@danzx2010.com,     Test Subject - Send Only,     This is my otherbody text.,   SendOnly

 
Tests:
   AutodiscoverOnly - Perform Autodiscover (SCP Autodiscover will also be done).
   AutodiscoverOnlyNoSCP - Perform Autodiscover for an external server only.
            This should be used for Outlook 365 and other out of network servers.
	SearchInbox - Search the wellknown inbox for the string specified in Param1.
		SearchInbox (search string)
		SearchInbox (search string, poperty to serach)

	    Param1:
		    If Param1 is set to "RANDOM" a random search string will be generated for every search.
		Param2:
		    Param2 specifies the property to serach against.
		    Param2 will default to the subject being searhed if its not specified.  
		        The following properties can be searched against:
			        Subject
			        Body
			        ItemClass
			Param3 is used to change what is returned by the seach.
				If this is not set then the default is to return the following:

				        BasePropertySet.IdOnly 
                        ItemSchema.Subject 
                        ItemSchema.DisplayTo 
                        ItemSchema.DisplayCc 
                        ItemSchema.DateTimeReceived 
                        ItemSchema.HasAttachments 
                        ItemSchema.ItemClass
                        ExtendedPropertyDefinition(0x1000, MapiPropertyType.String)   - PR_BODY
                        ExtendedPropertyDefinition(0x1035, MapiPropertyType.String)   - CdoPR_INTERNET_MESSAGE_ID 
                        ExtendedPropertyDefinition(0x0C1A, MapiPropertyType.String)   - CdoPR_SENDER_NAME

				If set to "IdOnly" then only the item identifiers (ItemId and ChangeKey) are returned.

	SearchCalendar - Search for all wellknown folder calendar events in the next 3 months.
		SearchCalendar()

	ReadAllContacts - Reads properties on all contacts in the wellknown contacts folder.
		ReadAllContacts()

	ReadMimeForInboxAllItems - Reads MIME from all items in the wellknown inbox folder.
		ReadMimeForInboxAllItems()

	ResolveRecipient - Resolves a recipient.  
		ResolveRecipient (search string)
		ResolveRecipient (search string, search scope)

		Param1:
		    Param1 holds the string to resolve against.
		    If Param1 is set to "RANDOM" a random search string will be generated for every search. Of course this will result
		    in a lot of failed resolutions.
		Param2:
		    Param2 is used to designate the scope of resolution. If not specified then the default is ContactsThenDirectory.
 
				ContactsThenDirectory 
			    ContactsOnly 
			    DirectoryOnly               
			    DirectoryThenContacts 

			 If not specified then the default is ContactsThenDirectory.

	SendEmail - Sends an email to a designated user's smtp address.
			SendEmail(smtp address, subject, body, SendOnly or Save and send)
			 
			Param1: To SMTP Address
            Param2: Subject
            Param3: Body
			Param4: SendOnly - Default is SaveCopyOnSend if left blank. 
					SendOnly - Just send the message and don't save it in the sender's Sent Items folder.
					SaveCopyOnSend - Send the message and save it in the sender's Sent Items folder.
 

Searches can generate a high load and run long - this may be a plus for certain types of testing.  You should be able
to run multiple instances of this application and also run it on multiple boxes in order to increase the volume of calls.
Note that since this application can create a lot of threads which each is doing separate calls against Exchange, you should
shut down anything not needed on the box in order to have the highest amount of resources available.  This application
does not do any checking on resource and could go past process/thread/system limitations if testing criteria are extremely high.
 
Requirements:

	This application uses .NET 4.0 (must be installed)
	The 2.0 version of the Exchange Managed API (inlcuded)