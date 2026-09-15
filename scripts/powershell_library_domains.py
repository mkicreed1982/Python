#!/usr/bin/env python3
"""
Which library is which: the domain and the offline cost of every name that
extract_powershell_libraries.py pulls out of PowerShell, VBScript and VBA code.

domain | what the dependency is there for
------ | -----------------------------------------------------------------
excel  | reading or writing workbooks
vba    | macros and the VBA object model
xml    | XML, XSLT, XSD, and the Open XML parts inside an .xlsx
office | the rest of Office automation: Word, Outlook, Access
data   | databases, ADO.NET, ODBC and OLEDB, CSV parsing

availability | what it takes to run the code on a machine with no internet
------------ | ----------------------------------------------------------
builtin      | ships with Windows, .NET or PowerShell, so nothing to do
office       | needs Microsoft Office, or the Access Database Engine, installed
external     | comes from the PowerShell Gallery or NuGet: stage it beforehand
local        | a file that travels with the code
unknown      | could not be placed, check it by hand

The tables below are meant to be edited: add the module or assembly names your
own finance and accounting scripts use and the reports follow.
"""

from __future__ import annotations

import re

DOMAINS = ("excel", "vba", "xml", "office", "data")
AVAILABILITIES = ("builtin", "office", "external", "local", "unknown")

# Matched in order against the case folded library name; first hit wins.
CATALOG: tuple[tuple[str, str, str], ...] = (
    # Excel through an installed copy of Office: COM programmatic identifiers
    # and the interop assemblies that wrap them.
    (r"excel\.[\w.]+", "excel", "office"),
    (r"microsoft\.office\.interop\.excel(\.dll)?", "excel", "office"),
    (r"microsoft\.office\.(core|tools)(\..*)?(\.dll)?", "office", "office"),
    (r"(word|outlook|access|powerpoint|publisher)\.[\w.]+", "office", "office"),
    (r"microsoft\.office\.interop\.(word|outlook|access)(\.dll)?", "office", "office"),
    # The VBA object model itself.
    (r"microsoft\.vbe\.interop(\..*)?(\.dll)?", "vba", "office"),
    (r"vbide\..*", "vba", "office"),
    (r"msforms\..*", "vba", "office"),
    # Workbook libraries that do not need Excel on the machine.  Ship the DLL
    # next to the script, or the module in a local repository.
    (r"importexcel", "excel", "external"),
    (r"psexcel|pswriteexcel", "excel", "external"),
    (r"epplus(\..*)?(\.dll)?|officeopenxml(\.dll)?", "excel", "external"),
    (r"closedxml(\..*)?(\.dll)?", "excel", "external"),
    (r"documentformat\.openxml(\..*)?(\.dll)?", "excel", "external"),
    (r"npoi(\..*)?(\.dll)?|exceldatareader(\..*)?(\.dll)?", "excel", "external"),
    (r"spreadsheetlight(\.dll)?|spreadsheetgear(\..*)?(\.dll)?", "excel", "external"),
    (
        r"aspose\.cells(\.dll)?|syncfusion\.xlsio(\.dll)?|gembox\.spreadsheet(\.dll)?",
        "excel",
        "external",
    ),
    # XML: all of it is in the box, including the plumbing an .xlsx is made of.
    (r"system\.xml(\..*)?(\.dll)?", "xml", "builtin"),
    (r"msxml\d*(\..*)?", "xml", "builtin"),
    (r"system\.io\.packaging(\.dll)?|windowsbase(\.dll)?", "xml", "builtin"),
    (r"system\.io\.compression(\..*)?(\.dll)?", "xml", "builtin"),
    (r"system\.private\.xml(\.dll)?", "xml", "builtin"),
    # Where the numbers come from.
    (r"system\.data(\..*)?(\.dll)?", "data", "builtin"),
    (r"adodb\..*|adox\..*", "data", "builtin"),
    (r"microsoft\.(ace|jet)\.oledb.*", "data", "office"),
    (r"microsoft (excel|text) driver.*", "excel", "office"),
    (r"microsoft access( text)? driver.*", "data", "office"),
    (r"odbc driver \d+ for sql server|sql server native client.*", "data", "external"),
    (r"sqloledb|msoledbsql\d*|sqlncli\d*", "data", "external"),
    (r"sql server", "data", "builtin"),
    (r"scripting\.(filesystemobject|dictionary|encoder)", "data", "builtin"),
    # COM servers that are simply part of Windows
    (r"shell\.application|wscript\.(shell|network)", "", "builtin"),
    (r"schedule\.service|winhttp\.[\w.]+|wbemscripting\.[\w.]+", "", "builtin"),
    (r"microsoft\.update\.[\w.]+|internetexplorer\.application", "", "builtin"),
    (r"microsoft\.visualbasic(\..*)?(\.dll)?", "data", "builtin"),
    (r"system\.globalization(\.dll)?|system\.numerics(\.dll)?", "data", "builtin"),
    (
        r"microsoft\.data\.sqlite(\.dll)?|system\.data\.sqlite(\.dll)?",
        "data",
        "external",
    ),
    (r"pssqlite|simplysql", "data", "external"),
    (r"sqlserver|sqlps|dbatools", "data", "external"),
)

# Cmdlets that give away a module dependency even when nothing imports it.
CMDLET_MODULES = {
    "import-excel": "ImportExcel",
    "export-excel": "ImportExcel",
    "open-excelpackage": "ImportExcel",
    "close-excelpackage": "ImportExcel",
    "add-worksheet": "ImportExcel",
    "add-pivottable": "ImportExcel",
    "add-excelchart": "ImportExcel",
    "add-conditionalformatting": "ImportExcel",
    "set-excelrange": "ImportExcel",
    "set-excelcolumn": "ImportExcel",
    "set-excelrow": "ImportExcel",
    "join-worksheet": "ImportExcel",
    "get-excelsheetinfo": "ImportExcel",
    "get-excelworkbookinfo": "ImportExcel",
    "send-sqldatatoexcel": "ImportExcel",
    "convertfrom-excelsheet": "ImportExcel",
    "convertto-excelxlsx": "ImportExcel",
    "new-conditionaltext": "ImportExcel",
    "new-excelchartdefinition": "ImportExcel",
    "import-xlsx": "PSExcel",
    "export-xlsx": "PSExcel",
    "select-xml": "Microsoft.PowerShell.Utility",
    "convertto-xml": "Microsoft.PowerShell.Utility",
    "export-clixml": "Microsoft.PowerShell.Utility",
    "import-clixml": "Microsoft.PowerShell.Utility",
    "import-csv": "Microsoft.PowerShell.Utility",
    "export-csv": "Microsoft.PowerShell.Utility",
    "convertfrom-csv": "Microsoft.PowerShell.Utility",
    "convertto-csv": "Microsoft.PowerShell.Utility",
    "invoke-sqlcmd": "SqlServer",
    "read-sqltabledata": "SqlServer",
    "write-sqltabledata": "SqlServer",
    "invoke-sqlitequery": "PSSQLite",
    "new-sqliteconnection": "PSSQLite",
    "invoke-dbaquery": "dbatools",
    "import-dbacsv": "dbatools",
}

# Evidence that a script drives macros rather than just reading a workbook.
MACRO_INDICATORS = (
    (r"\.RunAutoMacros\b", "RunAutoMacros"),
    (r"\bApplication\s*\.\s*Run\b", "Application.Run"),
    (r"\bExecuteExcel4Macro\b", "ExecuteExcel4Macro"),
    (r"\.VBProject\b", "VBProject"),
    (r"\.VBComponents\b", "VBComponents"),
    (r"\bAutomationSecurity\b", "AutomationSecurity"),
    (r"\bmsoAutomationSecurity\w*\b", "msoAutomationSecurity"),
    (r"[\w.$/\\-]*[\w$-]\.(?:xlsm|xlam|xlsb|xla|xltm)\b", ""),
)

# Names that are simply present on a Windows box with PowerShell on it.
BUILTIN_PREFIXES = (
    "system.",
    "microsoft.powershell.",
    "microsoft.win32.",
    "microsoft.management.infrastructure",
    "microsoft.wsman.",
    "windows.",
    "netstandard",
    "mscorlib",
)
BUILTIN_MODULES = frozenset(
    {
        "cimcmdlets",
        "iscsi",
        "microsoft.wsman.management",
        "packagemanagement",
        "powershellget",
        "psdiagnostics",
        "psreadline",
        "psworkflow",
        "scheduledtasks",
        "storage",
        "threadjob",
    }
)
WINDOWS_NATIVE_RE = re.compile(
    r"(kernel32|user32|advapi32|gdi32|shell32|shlwapi|ole32|oleaut32|ntdll|msvcrt"
    r"|crypt32|wintrust|wininet|ws2_32|psapi|pdh|slc|powrprof|version|setupapi"
    r"|iphlpapi|netapi32|secur32|userenv|winspool\.drv|api-ms-win-[\w.-]+)(\.dll)?",
    re.IGNORECASE,
)
LOCAL_SUFFIXES = (".psm1", ".psd1", ".ps1", ".cdxml", ".bas", ".cls", ".frm", ".vbs")

_CATALOG_RE = tuple(
    (re.compile(pattern, re.IGNORECASE), domain, availability)
    for pattern, domain, availability in CATALOG
)


def classify(name: str, kind: str = "") -> tuple[str, str]:
    """
    Place a library name in a domain and say what it costs to run offline.

    >>> classify("ImportExcel", "module")
    ('excel', 'external')
    >>> classify("Excel.Application", "com")
    ('excel', 'office')
    >>> classify("System.Xml.Linq", "namespace")
    ('xml', 'builtin')
    >>> classify("Budget.xlsm", "macro")
    ('vba', 'office')
    >>> classify("kernel32.dll", "native")
    ('', 'builtin')
    >>> classify("build.psm1", "module")
    ('', 'local')
    >>> classify("PSFramework", "module")
    ('', 'external')
    """
    key = name.casefold()
    for pattern, domain, availability in _CATALOG_RE:
        if pattern.fullmatch(key):
            return domain, availability
    if kind == "macro":
        return "vba", "office"
    if key.endswith(LOCAL_SUFFIXES):
        return "", "local"
    if key.startswith(BUILTIN_PREFIXES) or key in BUILTIN_MODULES:
        return "", "builtin"
    if kind == "native":
        return ("", "builtin") if WINDOWS_NATIVE_RE.fullmatch(key) else ("", "unknown")
    if kind == "com":
        return "", "unknown"
    return "", "external"


STAGING_ADVICE = {
    "external": "stage it first: Save-Module / nuget install on a connected box",
    "office": "needs Microsoft Office or the Access Database Engine installed",
    "unknown": "not recognised, check by hand",
    "builtin": "ships with Windows, .NET or PowerShell",
    "local": "travels with the code",
}
