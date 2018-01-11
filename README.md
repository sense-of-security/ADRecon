# ActiveDirectoryRecon

ADRecon is a tool which extracts various artifacts (as highlighted below) out of an AD environment in a specially formatted Microsoft Excel report that includes summary views with metrics to facilitate analysis.
The report can provide a holistic picture of the current state of the target AD environment.
The tool is useful to various classes of security professionals like auditors, DIFR, students, administrators, etc. It can also be an invaluable post-exploitation tool for a penetration tester.
It can be run from any workstation that is connected to the environment, even hosts that are not domain members. Furthermore, the tool can be executed in the context of a non-privileged (i.e. standard domain user) accounts. Fine Grained Password Policy, LAPS and BitLocker may require Privileged user accounts.
The tool will use Microsoft Remote Server Administration Tools (RSAT) if available, otherwise it will communicate with the Domain Controller using LDAP.
The following information is gathered by the tool:

- Forest;
- Domains in the Forest and other attributes such as Sites;
- Domain Password Policy;
- Domain Controllers and their roles;
- Users and their attributes;
- Service Principal Names;
- Groups and and their members;
- Organizational Units and their ACLs;
- Group Policy Object details;
- DNS Zones and Records;
- Printers;
- Computers and their attributes;
- LAPS passwords (if implemented); and
- BitLocker Recovery Keys (if implemented).

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
PS C:\>.\ADRecon.ps1 -Protocol ADWS -DomainController <IP or FQDN> -Credential <domain\username> -Collect Domian, DCs
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
    Path for ADRecon output folder to save the CSV files and the ADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist) (Default pwd)

-Collect <String>
    What attributes to collect (Comma separated; e.g Forest,Domain)
    Valid values include: Forest, Domain, PasswordPolicy, DCs, Users, UserSPNs, Groups, GroupMembers, OUs, OUPermissions, GPOs, GPOReport, DNSZones, Printers, Computers, ComputerSPNs, LAPS, BitLocker.

-OutputType <String>
    Output Type; Comma seperated; e.g CSV,STDOUT,Excel (Default STDOUT with -Collect parameter, else CSV and Excel).
    Valid values include: STDOUT, CSV, Excel.

-DormantTimeSpan <Int>
    Timespan for Dormant accounts. (Default 90 days)

-PageSize <Int>
    The PageSize to set for the LDAP searcher object. (Default 200)

-Threads <Int>
    The number of threads to use during processing objects (Default 10)

-FlushCount <Int>
    The number of processed objects which will be flushed to disk. (Default -1; Flush after all objects are processed).

```

### Future Plans

- Replace System.DirectoryServices.DirectorySearch with System.DirectoryServices.Protocols and add support for LDAP STARTTLS and LDAPS (TCP port 636).
- Add Domain Trust Enumeration.
- Gather ACLs for the useraccountcontrol attribute and the ms-mcs-admpwd LAPS attribute to determine which users can read the values.
- Gather DS_CONTROL_ACCESS and Extended Rights, such as User-Force-Change-Password, DS-Replication-Get-Changes, DS-Replication-Get-Changes-All, etc. which can be used as alternative attack vectors.
- Additional export and storage option: export to ~STDOUT~, SQLite, xml, html.
- List issues identified and provide recommended remediation advice based on analysis of the data.

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
