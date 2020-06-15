# Download from https://github.com/adrecon/ADRecon

# ADRecon: Active Directory Recon [![Follow ADRecon on Twitter](https://img.shields.io/twitter/follow/ad_recon.svg?style=social&label=Follow%20%40ad_recon)](https://twitter.com/intent/user?screen_name=ad_recon "Follow ADRecon on Twitter")

ADRecon is a tool which extracts and combines various artefacts (as highlighted below) out of an AD environment. The information can be presented in a specially formatted Microsoft Excel report that includes summary views with metrics to facilitate analysis and provide a holistic picture of the current state of the target AD environment.

The tool is useful to various classes of security professionals like auditors, DFIR, students, administrators, etc. It can also be an invaluable post-exploitation tool for a penetration tester.

It can be run from any workstation that is connected to the environment, even hosts that are not domain members. Furthermore, the tool can be executed in the context of a non-privileged (i.e. standard domain user) account. Fine Grained Password Policy, LAPS and BitLocker may require Privileged user accounts. The tool will use Microsoft Remote Server Administration Tools (RSAT) if available, otherwise it will communicate with the Domain Controller using LDAP.

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
- GroupPolicy objects and gPLink details;
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

ADRecon was presented at: [![Black Hat Arsenal Asia 2018](https://github.com/toolswatch/badges/blob/master/arsenal/asia/2018.svg)](https://www.blackhat.com/asia-18/arsenal.html#adrecon-active-directory-recon) - [Slidedeck](https://speakerdeck.com/prashant3535/adrecon-bh-asia-2018-arsenal-presentation)

[![Black Hat Arsenal USA 2018](https://github.com/toolswatch/badges/blob/master/arsenal/usa/2018.svg)](https://www.blackhat.com/us-18/arsenal/schedule/index.html#adrecon-active-directory-recon-11912) | [![DEFCON 26 Demolabs](https://hackwith.github.io/badges/defcon/26/demolabs.svg)](https://www.defcon.org/html/defcon-26/dc-26-demolabs.html) - [Slidedeck](https://speakerdeck.com/prashant3535/adrecon-bh-usa-2018-arsenal-and-def-con-26-demo-labs-presentation)

[Bay Area OWASP](https://www.meetup.com/en-AU/Bay-Area-OWASP/events/253585385/) - [Slidedeck](https://speakerdeck.com/prashant3535/active-directory-recon-101-owasp-bay-area-presentation)

[CHCON](https://2018.chcon.nz/mainevent.html) - [Slidedeck](https://speakerdeck.com/prashant3535/adrecon-detection-chcon-2018)

## Getting Started

These instructions will get you a copy of the tool up and running on your local machine.

### Prerequisites

- .NET Framework 3.0 or later (Windows 7 includes 3.0)
- PowerShell 2.0 or later (Windows 7 includes 2.0)

### Optional

- Microsoft Excel (to generate the report)
- Remote Server Administration Tools (RSAT):
  - Windows 10 (https://www.microsoft.com/en-au/download/details.aspx?id=45520)
  - Windows 7 (https://www.microsoft.com/en-au/download/details.aspx?id=7887)

### Installing

If you have git installed, you can start by cloning the [repository](https://github.com/sense-of-security/ADRecon/):

```
git clone https://github.com/sense-of-security/ADRecon.git
```

Otherwise, you can [download a zip archive of the latest release](https://github.com/sense-of-security/ADRecon/archive/master.zip). The intent is to always keep the master branch in a working state.

## Usage

### Examples

To run ADRecon on a domain member host.

```
PS C:\> .\ADRecon.ps1
```

To run ADRecon on a domain member host as a different user.

```
PS C:\>.\ADRecon.ps1 -DomainController <IP or FQDN> -Credential <domain\username>
```

To run ADRecon on a non-member host using LDAP.

```
PS C:\>.\ADRecon.ps1 -Protocol LDAP -DomainController <IP or FQDN> -Credential <domain\username>
```

To run ADRecon with specific modules on a non-member host with RSAT. (Default OutputType is STDOUT with -Collect parameter)

```
PS C:\>.\ADRecon.ps1 -Protocol ADWS -DomainController <IP or FQDN> -Credential <domain\username> -Collect Domain, DomainControllers
```

To generate the ADRecon-Report.xlsx based on ADRecon output (CSV Files).

```
PS C:\>.\ADRecon.ps1 -GenExcel C:\ADRecon-Report-<timestamp>
```

When you run ADRecon, a `ADRecon-Report-<timestamp>` folder will be created which will contain ADRecon-Report.xlsx and CSV-Folder with the raw files.

### Parameters

```
-Protocol <String>
    Which protocol to use; ADWS (default) or LDAP

-DomainController <String>
    Domain Controller IP Address or Domain FQDN.

-Credential <PSCredential>
    Domain Credentials.

-GenExcel <String>
    Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx. Use it to generate the ADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run ADRecon.

-OutputDir <String>
    Path for ADRecon output folder to save the CSV/XML/JSON/HTML files and the ADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist) (Default pwd)

-Collect <String>
    Which modules to run (Comma separated; e.g Forest,Domain. Default all except Kerberoast)
    Valid values include: Forest, Domain, Trusts, Sites, Subnets, PasswordPolicy, FineGrainedPasswordPolicy, DomainControllers, Users, UserSPNs, PasswordAttributes, Groups, GroupMembers, OUs, ACLs, GPOs, gPLinks, GPOReport, DNSZones, Printers, Computers, ComputerSPNs, LAPS, BitLocker, Kerberoast DomainAccountsusedforServiceLogon.

-OutputType <String>
    Output Type; Comma seperated; e.g CSV,STDOUT,Excel (Default STDOUT with -Collect parameter, else CSV and Excel).
    Valid values include: STDOUT, CSV, XML, JSON, HTML, Excel, All (excludes STDOUT).

-DormantTimeSpan <Int>
    Timespan for Dormant accounts. (Default 90 days)

-PassMaxAge <Int>
    Maximum machine account password age. (Default 30 days)

-PageSize <Int>
    The PageSize to set for the LDAP searcher object. (Default 200)

-Threads <Int>
    The number of threads to use during processing objects (Default 10)

-Log <Switch>
    Create ADRecon Log using Start-Transcript
```

### Future Plans

- Replace System.DirectoryServices.DirectorySearch with System.DirectoryServices.Protocols and add support for LDAP STARTTLS and LDAPS (TCP port 636).
- ~~Add Domain Trust Enumeration.~~
- Add option to filter default ACLs.
- ~~Gather ACLs for other objects such as Users, Group, etc.~~
- Additional export and storage option: export to ~~STDOUT~~, SQLite, ~~xml~~, ~~json~~, ~~html~~, pdf.
- Use the EPPlus library for Excel Report generation and remove the dependency on MS Excel.
- List issues identified and provide recommended remediation advice based on analysis of the data.
- Add PowerShell Core support.

### Bugs, Issues and Feature Requests

Please report all bugs, issues and feature requests in the [issue tracker](https://github.com/sense-of-security/ADRecon/issues). Or let me (@prashant3535) know directly.

### Contributing

Pull request are always welcome.

### Mad props

Thanks for the awesome work by @_wald0, @CptJesus, @harmj0y, @mattifestation, @PyroTek3, @darkoperator, the Sense of Security Team and others.

### License

ADRecon is a tool which gathers information about the Active Directory and generates a report which can provide a holistic picture of the current state of the target AD environment.

Copyright (C) Sense of Security

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Affero General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License along with this program. If not, see http://www.gnu.org/licenses/.

This program borrows and uses code from many sources. All attempts are made to credit the original author. If you find that your code is used without proper credit, please shoot an insult to @prashant3535, Thanks.
