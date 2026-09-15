#!/usr/bin/env python3
"""
Extract the libraries a PowerShell, VBScript or VBA code base depends on.

The script walks ``*.ps1``, ``*.psm1``, ``*.psd1`` and the Basic sources
``*.vbs``, ``*.bas``, ``*.cls``, ``*.frm``, and reports every dependency it can
find:

kind      | recovered from
--------- | ----------------------------------------------------------------
module    | ``#requires -Modules``, ``using module``, ``Import-Module``,
          | ``Install-Module``, ``Save-Module``, manifest ``RequiredModules``,
          | ``NestedModules``, ``RootModule``, and cmdlets that give a module
          | away, such as ``Export-Excel`` meaning ``ImportExcel``
namespace | ``using namespace``, ``New-Object System.Xml.XmlDocument``,
          | ``[System.Data.DataTable]``, the ``[xml]`` accelerator
assembly  | ``using assembly``, ``Add-Type -AssemblyName|-Path``,
          | ``[Reflection.Assembly]::Load*``, manifest ``RequiredAssemblies``
com       | ``New-Object -ComObject Excel.Application``, VBA and VBScript
          | ``CreateObject``/``GetObject``, early binding (``As Excel.Range``),
          | and the provider named in a connection string
snapin    | ``#requires -PSSnapin``, ``Add-PSSnapin``
native    | ``[DllImport("kernel32.dll")]``, VBA ``Declare ... Lib "kernel32"``
macro     | ``Application.Run``, ``RunAutoMacros``, ``VBProject``, and the
          | macro-enabled workbooks (``.xlsm``, ``.xlam``) a script opens

Every name is then placed in a domain (excel, vba, xml, office, data) and given
an offline cost (builtin, office, external, local), so a report says what has to
be staged before the code runs on a machine with no internet.  The tables that
drive that live in powershell_library_domains.py and are meant to be edited.

Comments and string literals are tracked, so a command name inside a comment or
a message string is not mistaken for a real dependency.  References whose value
is a runtime expression (``Import-Module $name``) cannot be resolved
statically: they are counted, and only listed with ``--include-dynamic``.

Usage:
    python extract_powershell_libraries.py PATH [PATH ...]
    python extract_powershell_libraries.py repo --details
    python extract_powershell_libraries.py repo --domains excel,vba,xml
    python extract_powershell_libraries.py repo --format json > libraries.json
"""

from __future__ import annotations

import argparse
import csv
import dataclasses
import io
import json
import os
import re
import sys
from collections import Counter
from collections.abc import Iterator
from pathlib import Path

try:
    from .powershell_library_domains import (
        AVAILABILITIES,
        CMDLET_MODULES,
        DOMAINS,
        MACRO_INDICATORS,
        STAGING_ADVICE,
        classify,
    )
except ImportError:  # running the file directly instead of as a package
    from powershell_library_domains import (  # type: ignore[no-redef]
        AVAILABILITIES,
        CMDLET_MODULES,
        DOMAINS,
        MACRO_INDICATORS,
        STAGING_ADVICE,
        classify,
    )

CODE = "c"
STRING = "s"
COMMENT = "#"

POWERSHELL_SUFFIXES = (".ps1", ".psm1", ".psd1")
BASIC_SUFFIXES = (".vbs", ".vba", ".bas", ".cls", ".frm")  # VBScript and VBA
DEFAULT_SUFFIXES = POWERSHELL_SUFFIXES + BASIC_SUFFIXES
KIND_ORDER = ("module", "namespace", "assembly", "com", "snapin", "native", "macro")

# Parameters of a command that carry a library name; "" is the first positional.
MODULE_PARAMETERS = {
    "name": "module",
    "fullyqualifiedname": "module",
    "assembly": "assembly",
    "": "module",
}
SNAPIN_PARAMETERS = {"name": "snapin", "": "snapin"}
ADD_TYPE_PARAMETERS = {
    "assemblyname": "assembly",
    "referencedassemblies": "assembly",
    "path": "assembly",
    "literalpath": "assembly",
}
COMMANDS = {
    "import-module": MODULE_PARAMETERS,
    "ipmo": MODULE_PARAMETERS,
    "install-module": MODULE_PARAMETERS,
    "save-module": MODULE_PARAMETERS,
    "add-pssnapin": SNAPIN_PARAMETERS,
    "asnp": SNAPIN_PARAMETERS,
    "add-type": ADD_TYPE_PARAMETERS,
}
REQUIRES_PARAMETERS = {
    "modules": "module",
    "pssnapin": "snapin",
    "assembly": "assembly",
}
# Parameters whose value is never a library name; listing them keeps a
# positional name such as "Pester" in `Import-Module -Force Pester` reachable
# without mistaking `Stop` in `-ErrorAction Stop` for a library.
VALUE_PARAMETERS = frozenset(
    {
        "alias",
        "argumentlist",
        "cimsession",
        "cmdlet",
        "compileroptions",
        "credential",
        "erroraction",
        "errorvariable",
        "function",
        "informationaction",
        "informationvariable",
        "language",
        "maximumversion",
        "memberdefinition",
        "minimumversion",
        "moduleinfo",
        "namespace",
        "outbuffer",
        "outputassembly",
        "outputtype",
        "outvariable",
        "pipelinevariable",
        "prefix",
        "pssession",
        "repository",
        "requiredversion",
        "scope",
        "typedefinition",
        "variable",
        "version",
        "warningaction",
        "warningvariable",
    }
)
MANIFEST_KEYS = {
    "requiredmodules": "module",
    "nestedmodules": "module",
    "rootmodule": "module",
    "moduletoprocess": "module",
    "requiredassemblies": "assembly",
}

COMMAND_RE = re.compile(
    r"(?<![\w.`-])(" + "|".join(sorted(COMMANDS)) + r")(?![\w-])", re.IGNORECASE
)
REQUIRES_RE = re.compile(r"^[ \t]*#requires\b(?P<rest>[^\n]*)", re.IGNORECASE | re.M)
USING_RE = re.compile(
    r"(?:^|;)[ \t]*(?P<kw>using)[ \t]+(?P<what>module|namespace|assembly)[ \t]+",
    re.IGNORECASE | re.M,
)
ASSEMBLY_LOAD_RE = re.compile(
    r"\[[ \t]*(?:System\.)?Reflection\.Assembly[ \t]*\][ \t]*::[ \t]*"
    r"(?:Unsafe)?Load(?:WithPartialName|From|File)?[ \t]*\([ \t]*",
    re.IGNORECASE,
)
DLL_IMPORT_RE = re.compile(
    r"DllImport[ \t]*\([ \t]*(?P<quote>[\"'])(?P<lib>[^\"']+)(?P=quote)"
)
MANIFEST_KEY_RE = re.compile(
    r"^[ \t]*(?P<key>" + "|".join(sorted(MANIFEST_KEYS)) + r")[ \t]*=[ \t]*",
    re.IGNORECASE | re.M,
)
HASHTABLE_NAME_RE = re.compile(
    r"ModuleName[ \t]*=[ \t]*(?P<quote>[\"'])(?P<name>[^\"']+)(?P=quote)", re.IGNORECASE
)
NEW_OBJECT_RE = re.compile(r"(?<![\w.`-])New-Object(?![\w-])", re.IGNORECASE)
NEW_OBJECT_PARAMETERS = {
    "comobject": "com",
    "typename": "type",
    "": "type",
}
TYPE_LITERAL_RE = re.compile(r"\[\s*(?P<type>(?:[A-Za-z_]\w*\.){1,6}[A-Za-z_]\w*)\s*\]")
XML_ACCELERATOR_RE = re.compile(r"\[\s*xml\s*\]", re.IGNORECASE)
CMDLET_RE = re.compile(
    r"(?<![\w.`-])(" + "|".join(sorted(CMDLET_MODULES)) + r")(?![\w-])",
    re.IGNORECASE,
)
MACRO_RES = tuple(
    (re.compile(pattern, re.IGNORECASE), label) for pattern, label in MACRO_INDICATORS
)
# VBScript and VBA, where a library is named by CreateObject, by a Declare
# statement or by an early bound type such as `Dim ws As Excel.Worksheet`.
CREATE_OBJECT_RE = re.compile(
    r"\b(?:Server\.)?CreateObject\s*\(\s*\"(?P<progid>[^\"]+)\"", re.IGNORECASE
)
GET_OBJECT_RE = re.compile(
    r"\bGetObject\s*\(\s*(?:[^,()]*,)?\s*\"(?P<progid>[^\"]+)\"", re.IGNORECASE
)
DECLARE_LIB_RE = re.compile(
    r"\bDeclare\b[^\n]*?\bLib\s+\"(?P<lib>[^\"]+)\"", re.IGNORECASE
)
BASIC_TYPE_RE = re.compile(
    r"\b(?:As|New)\s+(?P<type>[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)+)", re.IGNORECASE
)
# A connection string names the data provider a finance script leans on.
OLEDB_PROVIDER_RE = re.compile(r"\bProvider\s*=\s*(?P<provider>[\w.]+)", re.IGNORECASE)
ODBC_DRIVER_RE = re.compile(r"\bDriver\s*=\s*\{(?P<provider>[^}]+)\}", re.IGNORECASE)
FILE_LITERAL_RE = re.compile(
    r"[\"'][^\"']*?([^\"'/\\]+\.(?:psm1|psd1|ps1|dll))[\"']", re.IGNORECASE
)


@dataclasses.dataclass(frozen=True)
class Reference:
    """A single library reference found in PowerShell source."""

    kind: str
    name: str
    origin: str
    line: int
    dynamic: bool
    domain: str = ""
    availability: str = ""
    path: str = ""


def classify_characters(source: str) -> str:
    """
    Tag every character of ``source`` as code, string or comment.  The mask has
    one character per source character, so ``mask[i]`` says what ``source[i]``
    belongs to.

    >>> classify_characters("a 'b' # c")
    'ccsssc###'
    >>> classify_characters("<# x #>y")
    '#######c'
    >>> classify_characters('"a`"b" 1')
    'sssssscc'
    """
    mask: list[str] = []
    index, length = 0, len(source)
    while index < length:
        start, char = index, source[index]
        if source.startswith("<#", index):
            index = _skip_block_comment(source, index)
            mask.append(COMMENT * (index - start))
        elif char == "#" and _starts_token(source, index):
            end = source.find("\n", index)
            index = length if end < 0 else end
            mask.append(COMMENT * (index - start))
        elif source.startswith(('@"', "@'"), index):
            index = _skip_here_string(source, index)
            mask.append(STRING * (index - start))
        elif char in "\"'":
            index = _skip_string(source, index)
            mask.append(STRING * (index - start))
        else:
            mask.append(CODE)
            index += 1
    return "".join(mask)


def _starts_token(source: str, index: int) -> bool:
    """
    A ``#`` only opens a comment when it starts a token.

    >>> _starts_token("# x", 0), _starts_token("a#b", 1)
    (True, False)
    """
    return index == 0 or source[index - 1] in " \t\r\n;(){}|,&="


def _skip_block_comment(source: str, index: int) -> int:
    """Return the index just past a (nestable) ``<# ... #>`` comment."""
    depth, length = 0, len(source)
    while index < length:
        if source.startswith("<#", index):
            depth += 1
            index += 2
        elif source.startswith("#>", index):
            depth -= 1
            index += 2
            if depth == 0:
                return index
        else:
            index += 1
    return length


def _skip_here_string(source: str, index: int) -> int:
    """Return the index just past an ``@" ... "@`` or ``@' ... '@`` here-string."""
    terminator = "\n" + source[index + 1] + "@"
    end = source.find(terminator, index + 2)
    return len(source) if end < 0 else end + len(terminator)


def _skip_string(source: str, index: int) -> int:
    """
    Return the index just past a single or double quoted string.

    >>> _skip_string("'a''b' rest", 0)
    6
    >>> _skip_string('"a$(f("b"))c" rest', 0)  # nested quotes in a subexpression
    13
    """
    quote, length = source[index], len(source)
    index += 1
    while index < length:
        char = source[index]
        if quote == '"' and char == "`":
            index += 2
        elif quote == '"' and source.startswith("$(", index):
            index = _skip_subexpression(source, index + 1)
        elif char != quote:
            index += 1
        elif source[index + 1 : index + 2] == quote:  # "" and '' escape a quote
            index += 2
        else:
            return index + 1
    return length


def _skip_subexpression(source: str, index: int) -> int:
    """Return the index just past a ``$( ... )`` block, quotes and all."""
    depth, length = 0, len(source)
    while index < length:
        char = source[index]
        if char in "\"'":
            index = _skip_string(source, index)
            continue
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth == 0:
                return index + 1
        index += 1
    return length


def masked_source(text: str) -> tuple[str, str]:
    """
    Pair ``text`` with its mask, the input every scanning helper expects.

    >>> masked_source("-Force")
    ('-Force', 'cccccc')
    """
    return text, classify_characters(text)


def argument_segment(source: str, mask: str, start: int) -> tuple[str, str]:
    """
    Slice the argument text of a command: everything up to an unquoted pipe,
    semicolon, closing bracket or unescaped end of line.

    >>> argument_segment(*masked_source("Import-Module Foo -Force | Out-Null"), 13)
    (' Foo -Force ', 'cccccccccccc')
    >>> argument_segment(*masked_source("Import-Module A `\\n -Force\\nnext"), 13)[0]
    ' A `\\n -Force'
    """
    depth, index, length = 0, start, len(source)
    while index < length:
        char = source[index]
        if mask[index] == CODE:
            if char in "([{":
                depth += 1
            elif char in ")]}":
                if depth == 0:
                    break
                depth -= 1
            elif (depth == 0 and char in ";|") or (
                depth == 0
                and char == "\n"
                and not source[:index].rstrip(" \t").endswith(("`", ","))
            ):
                break
        index += 1
    return source[start:index], mask[start:index]


def split_arguments(segment: str, mask: str) -> list[str]:
    """
    Split command arguments into tokens.  Quoted strings and bracketed
    sub-expressions stay in one piece, a comma becomes its own token.

    >>> split_arguments(*masked_source("-Name Foo, Bar -Force"))
    ['-Name', 'Foo', ',', 'Bar', '-Force']
    >>> split_arguments(*masked_source("(Join-Path $a 'b.psm1') -Force"))
    ["(Join-Path $a 'b.psm1')", '-Force']
    >>> split_arguments(*masked_source("@{ModuleName='A'; Version='2'}"))
    ["@{ModuleName='A'; Version='2'}"]
    """
    tokens: list[str] = []
    current: list[str] = []
    depth = 0
    for index, char in enumerate(segment):
        is_code = mask[index] == CODE
        if is_code and char in "([{":
            depth += 1
        elif is_code and char in ")]}":
            depth -= 1
        if depth <= 0 and is_code and (char.isspace() or char == ","):
            if current:
                tokens.append("".join(current))
                current = []
            if char == ",":
                tokens.append(",")
            continue
        current.append(char)
    if current:
        tokens.append("".join(current))
    return tokens


def strip_quotes(token: str) -> str:
    """
    >>> strip_quotes('"Pester"'), strip_quotes("'Pester'"), strip_quotes("Pester")
    ('Pester', 'Pester', 'Pester')
    """
    if len(token) >= 2 and token[0] == token[-1] and token[0] in "\"'":
        return token[1:-1]
    return token


def normalize_name(token: str) -> tuple[str, bool]:
    """
    Reduce an argument to a library name, and report whether it is dynamic, i.e.
    only known at run time.  A file name embedded in an expression is recovered,
    because that is the part naming the library.

    >>> normalize_name("'Pester'")
    ('Pester', False)
    >>> normalize_name('"$PSScriptRoot/../build.psm1"')
    ('build.psm1', False)
    >>> normalize_name("(Join-Path $PSScriptRoot 'certificateCommon.psm1')")
    ('certificateCommon.psm1', False)
    >>> normalize_name('"$PSScriptRoot\\..\\Xml"')
    ('Xml', False)
    >>> normalize_name("${modulePath}")
    ('${modulePath}', True)
    >>> normalize_name("(Join-Path $repoRoot 'tools/packaging')")
    ("(Join-Path $repoRoot 'tools/packaging')", True)
    """
    embedded = FILE_LITERAL_RE.search(token)
    if embedded:  # a module or assembly file name inside an expression
        return embedded.group(1), False
    value = strip_quotes(token)
    if "(" in value:  # a sub-expression with nothing recoverable in it
        return token, True
    if "/" in value or "\\" in value:
        value = value.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1]
    dynamic = (
        not value
        or value in {".", ".."}
        or value.startswith("-")
        or any(char in value for char in "$({@`")
    )
    return (token if dynamic else value), dynamic


def is_parameter(token: str) -> bool:
    """
    >>> is_parameter("-Force"), is_parameter("-1"), is_parameter("Pester")
    (True, False, False)
    """
    return len(token) > 1 and token[0] == "-" and not token[1].isdigit()


def parameter_kind(token: str, parameters: dict[str, str]) -> str | None:
    """
    Map a written parameter to a reference kind, honouring PowerShell's
    unambiguous prefix rule (``-Assembly`` means ``-AssemblyName``).

    >>> parameter_kind("-Assembly", ADD_TYPE_PARAMETERS)
    'assembly'
    >>> parameter_kind("-PassThru", ADD_TYPE_PARAMETERS) is None
    True
    """
    written = token.lstrip("-").casefold()
    for full, kind in parameters.items():
        if full and full.startswith(written):
            return kind
    return None


def takes_value(token: str) -> bool:
    """
    Whether a parameter the scanner does not care about still swallows the
    token behind it, as opposed to being a switch such as ``-Force``.

    >>> takes_value("-ErrorAction"), takes_value("-Force")
    (True, False)
    """
    written = token.lstrip("-").casefold()
    return any(full.startswith(written) for full in VALUE_PARAMETERS)


def collect_values(
    tokens: list[str], parameters: dict[str, str]
) -> list[tuple[str, str]]:
    """
    Pull ``(kind, token)`` pairs out of an argument list.

    >>> arguments = split_arguments(*masked_source("A, B -Force"))
    >>> collect_values(arguments, MODULE_PARAMETERS)
    [('module', 'A'), ('module', 'B')]
    >>> arguments = split_arguments(*masked_source("-Path a.dll -PassThru"))
    >>> collect_values(arguments, ADD_TYPE_PARAMETERS)
    [('assembly', 'a.dll')]
    >>> arguments = split_arguments(*masked_source("-ErrorAction Stop -Force Pester"))
    >>> collect_values(arguments, MODULE_PARAMETERS)
    [('module', 'Pester')]
    """
    values: list[tuple[str, str]] = []
    positional = parameters.get("")
    kind: str | None = None
    expecting = skip_next = positional_used = False
    for position, argument in enumerate(tokens):
        if argument == ",":
            continue
        if is_parameter(argument):
            kind = parameter_kind(argument, parameters)
            expecting = kind is not None
            skip_next = not expecting and takes_value(argument)
            continue
        followed_by_comma = tokens[position + 1 : position + 2] == [","]
        if skip_next:
            skip_next = followed_by_comma
        elif expecting and kind is not None:
            values.append((kind, argument))
            expecting = followed_by_comma
        elif positional is not None and not positional_used:
            values.append((positional, argument))
            positional_used = not followed_by_comma
    return values


def line_of(source: str, index: int) -> int:
    """
    >>> line_of("a\\nb", 2)
    2
    """
    return source.count("\n", 0, index) + 1


def _record(
    found: list[Reference],
    kind: str,
    token: str,
    origin: str,
    source: str,
    index: int,
) -> None:
    """Normalize one token, classify it and append it to the result list."""
    hashtable = HASHTABLE_NAME_RE.search(token)
    name, dynamic = normalize_name(hashtable.group("name") if hashtable else token)
    domain, availability = classify(name, kind)
    found.append(
        Reference(
            kind=kind,
            name=name,
            origin=origin,
            line=line_of(source, index),
            dynamic=dynamic,
            domain=domain,
            availability=availability,
        )
    )


def scan_requires(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``#requires -Modules Pester``, which lives inside a comment."""
    for match in REQUIRES_RE.finditer(source):
        if mask[match.start()] != COMMENT:
            continue
        tokens = split_arguments(*masked_source(match.group("rest")))
        for kind, token in collect_values(tokens, REQUIRES_PARAMETERS):
            _record(found, kind, token, "#requires", source, match.start())


def scan_using(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``using module``, ``using namespace`` and ``using assembly``."""
    for match in USING_RE.finditer(source):
        if mask[match.start("kw")] != CODE:
            continue
        tokens = split_arguments(*argument_segment(source, mask, match.end()))
        if not tokens:
            continue
        what = match.group("what").casefold()
        _record(found, what, tokens[0], f"using {what}", source, match.start("kw"))


def scan_commands(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``Import-Module``, ``Add-Type``, ``Add-PSSnapin`` and friends."""
    for match in COMMAND_RE.finditer(source):
        if mask[match.start()] != CODE:
            continue
        parameters = COMMANDS[match.group(1).casefold()]
        tokens = split_arguments(*argument_segment(source, mask, match.end()))
        for kind, token in collect_values(tokens, parameters):
            _record(found, kind, token, match.group(1), source, match.start())


def scan_assembly_loads(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``[Reflection.Assembly]::LoadWithPartialName('System.Web')``."""
    for match in ASSEMBLY_LOAD_RE.finditer(source):
        if mask[match.start()] != CODE:
            continue
        tokens = split_arguments(*argument_segment(source, mask, match.end()))
        if tokens:
            origin = "Reflection.Assembly"
            _record(found, "assembly", tokens[0], origin, source, match.start())


def scan_dll_imports(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``[DllImport("kernel32.dll")]`` in inline C#, i.e. inside strings."""
    for match in DLL_IMPORT_RE.finditer(source):
        if mask[match.start()] == COMMENT:
            continue
        _record(found, "native", match.group("lib"), "DllImport", source, match.start())


def scan_manifest_keys(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``RequiredModules = @('Pester')`` in ``*.psd1`` manifests."""
    for match in MANIFEST_KEY_RE.finditer(source):
        if mask[match.start("key")] != CODE:
            continue
        kind = MANIFEST_KEYS[match.group("key").casefold()]
        pending = split_arguments(*argument_segment(source, mask, match.end()))
        while pending:
            entry = pending.pop(0)
            if entry == ",":
                continue
            if entry.startswith("@(") and entry.endswith(")"):
                pending.extend(split_arguments(*masked_source(entry[2:-1])))
                continue
            origin = match.group("key")
            _record(found, kind, entry, origin, source, match.start("key"))


def _record_type(
    found: list[Reference], type_name: str, origin: str, source: str, index: int
) -> None:
    """Record the namespace of a .NET type, but only for a tracked domain."""
    namespace = strip_quotes(type_name).rsplit(".", 1)[0]
    if "." not in namespace and not classify(namespace, "namespace")[0]:
        return
    if classify(namespace, "namespace")[0]:
        _record(found, "namespace", namespace, origin, source, index)


def scan_com_objects(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``New-Object -ComObject Excel.Application`` and typed New-Object."""
    for match in NEW_OBJECT_RE.finditer(source):
        if mask[match.start()] != CODE:
            continue
        tokens = split_arguments(*argument_segment(source, mask, match.end()))
        for kind, token in collect_values(tokens, NEW_OBJECT_PARAMETERS):
            if kind == "com":
                origin = "New-Object -ComObject"
                _record(found, "com", token, origin, source, match.start())
            else:
                _record_type(found, token, "New-Object", source, match.start())


def scan_domain_types(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``[System.Xml.XmlDocument]`` and the ``[xml]`` accelerator."""
    for match in TYPE_LITERAL_RE.finditer(source):
        if mask[match.start()] == CODE:
            _record_type(found, match.group("type"), "[type]", source, match.start())
    for match in XML_ACCELERATOR_RE.finditer(source):
        if mask[match.start()] == CODE:
            _record(found, "namespace", "System.Xml", "[xml]", source, match.start())


def scan_cmdlets(source: str, mask: str, found: list[Reference]) -> None:
    """Infer the module behind a cmdlet: ``Export-Excel`` means ``ImportExcel``."""
    for match in CMDLET_RE.finditer(source):
        if mask[match.start()] != CODE:
            continue
        module = CMDLET_MODULES[match.group(1).casefold()]
        _record(found, "module", module, match.group(1), source, match.start())


def scan_macro_indicators(source: str, mask: str, found: list[Reference]) -> None:
    """Record the signs that code drives macros, not just reads a workbook."""
    for pattern, label in MACRO_RES:
        for match in pattern.finditer(source):
            if mask[match.start()] == COMMENT:
                continue
            name = label or match.group(0)
            origin = "VBA API" if label else "macro-enabled file"
            _record(found, "macro", name, origin, source, match.start())


def scan_connection_strings(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``Provider=Microsoft.ACE.OLEDB.12.0`` and ODBC driver names."""
    for pattern in (OLEDB_PROVIDER_RE, ODBC_DRIVER_RE):
        for match in pattern.finditer(source):
            if mask[match.start()] == COMMENT:
                continue
            provider = match.group("provider")
            _record(found, "com", provider, "connection string", source, match.start())


def scan_basic_objects(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``CreateObject("Excel.Application")`` in VBScript and VBA."""
    for pattern, origin in (
        (CREATE_OBJECT_RE, "CreateObject"),
        (GET_OBJECT_RE, "GetObject"),
    ):
        for match in pattern.finditer(source):
            if mask[match.start()] == CODE:
                _record(
                    found, "com", match.group("progid"), origin, source, match.start()
                )


def scan_basic_declares(source: str, mask: str, found: list[Reference]) -> None:
    """Handle ``Declare PtrSafe Function GetTickCount Lib "kernel32" ()``."""
    for match in DECLARE_LIB_RE.finditer(source):
        if mask[match.start()] == CODE:
            _record(
                found,
                "native",
                match.group("lib"),
                "Declare Lib",
                source,
                match.start(),
            )


def scan_basic_types(source: str, mask: str, found: list[Reference]) -> None:
    """Handle early binding: ``Dim ws As Excel.Worksheet``."""
    for match in BASIC_TYPE_RE.finditer(source):
        if mask[match.start()] == CODE:
            _record(
                found,
                "com",
                match.group("type"),
                "early binding",
                source,
                match.start(),
            )


def classify_characters_basic(source: str) -> str:
    """
    The same mask for VBScript and VBA, where a comment starts at ``'`` or
    ``REM`` and a doubled quote escapes a quote inside a string.

    >>> classify_characters_basic("a = 1 ' note")
    'cccccc######'
    >>> classify_characters_basic('Set a = "x"" y" REM note')
    'ccccccccsssssssc########'
    """
    mask: list[str] = []
    index, length = 0, len(source)
    while index < length:
        start, char = index, source[index]
        is_rem = source[index : index + 3].casefold() == "rem" and _starts_token(
            source, index
        )
        if char == "'" or (is_rem and source[index + 3 : index + 4] in (" ", "\t", "")):
            end = source.find("\n", index)
            index = length if end < 0 else end
            mask.append(COMMENT * (index - start))
        elif char == '"':
            index = _skip_basic_string(source, index)
            mask.append(STRING * (index - start))
        else:
            mask.append(CODE)
            index += 1
    return "".join(mask)


def _skip_basic_string(source: str, index: int) -> int:
    """Return the index just past a Basic string, which ends at the line break."""
    length = len(source)
    index += 1
    while index < length:
        char = source[index]
        if char == "\n":
            return index
        if char == '"':
            if source[index + 1 : index + 2] == '"':
                index += 2
                continue
            return index + 1
        index += 1
    return length


POWERSHELL_SCANNERS = (
    scan_requires,
    scan_using,
    scan_commands,
    scan_assembly_loads,
    scan_dll_imports,
    scan_manifest_keys,
    scan_com_objects,
    scan_domain_types,
    scan_cmdlets,
    scan_macro_indicators,
    scan_connection_strings,
)
BASIC_SCANNERS = (
    scan_basic_objects,
    scan_basic_declares,
    scan_basic_types,
    scan_macro_indicators,
    scan_connection_strings,
)
LANGUAGES = {
    "powershell": (classify_characters, POWERSHELL_SCANNERS),
    "basic": (classify_characters_basic, BASIC_SCANNERS),
}


def language_of(path: str) -> str:
    """
    >>> language_of("tools/build.psm1"), language_of("Macros/Ledger.bas")
    ('powershell', 'basic')
    """
    return "basic" if Path(path).suffix.casefold() in BASIC_SUFFIXES else "powershell"


def extract_references(
    source: str, path: str = "", language: str = ""
) -> list[Reference]:
    """
    Extract every library reference from one source text.  The language is
    taken from the file name unless it is given.

    >>> source = '''
    ... #Requires -Modules Pester, @{ModuleName='PSReadLine';ModuleVersion='2.0'}
    ... using namespace System.Text
    ... Import-Module -Name Microsoft.PowerShell.Archive -Force
    ... Add-Type -AssemblyName System.Windows.Forms
    ... Write-Host "Import-Module is only mentioned inside this string"
    ... # Import-Module NotADependency
    ... '''
    >>> for reference in extract_references(source):
    ...     print(reference.kind, reference.name, reference.origin, sep=" | ")
    module | Pester | #requires
    module | PSReadLine | #requires
    namespace | System.Text | using namespace
    module | Microsoft.PowerShell.Archive | Import-Module
    assembly | System.Windows.Forms | Add-Type

    >>> vba = 'Dim book As Excel.Workbook'
    >>> [(r.kind, r.name, r.domain, r.availability) for r in
    ...  extract_references(vba, "Ledger.bas")]
    [('com', 'Excel.Workbook', 'excel', 'office')]
    """
    mask_of, scanners = LANGUAGES[language or language_of(path)]
    mask = mask_of(source)
    found: list[Reference] = []
    for scanner in scanners:
        scanner(source, mask, found)
    found.sort(
        key=lambda reference: (
            reference.line,
            reference.kind,
            reference.name.casefold(),
        )
    )
    if not path:
        return found
    return [dataclasses.replace(reference, path=path) for reference in found]


def read_source(path: Path) -> str:
    """Read a PowerShell file, tolerating byte order marks and UTF-16."""
    raw = path.read_bytes()
    for bom, encoding in ((b"\xff\xfe", "utf-16-le"), (b"\xfe\xff", "utf-16-be")):
        if raw.startswith(bom):
            return raw[2:].decode(encoding, errors="replace")
    return raw.decode("utf-8-sig", errors="replace")


def powershell_files(paths: list[str], suffixes: tuple[str, ...]) -> Iterator[Path]:
    """Yield every PowerShell file under the given files and directories."""
    for entry in paths:
        path = Path(entry)
        if path.is_file():
            yield path
        elif path.is_dir():
            for candidate in sorted(path.rglob("*")):
                if ".git" in candidate.parts:
                    continue
                if candidate.is_file() and candidate.suffix.lower() in suffixes:
                    yield candidate
        else:
            print(f"warning: {entry} does not exist", file=sys.stderr)


def scan_files(
    paths: list[str], suffixes: tuple[str, ...] = DEFAULT_SUFFIXES
) -> tuple[list[Reference], Counter[str]]:
    """Return every reference found under ``paths``, plus a file count by suffix."""
    references: list[Reference] = []
    scanned: Counter[str] = Counter()
    for path in powershell_files(paths, suffixes):
        scanned[path.suffix.lower()] += 1
        references.extend(extract_references(read_source(path), str(path)))
    return references, scanned


Groups = dict[tuple[str, str], list[Reference]]


def group_references(references: list[Reference]) -> Groups:
    """Group references by kind and case insensitive name, most used first."""
    groups: Groups = {}
    for reference in references:
        key = (reference.kind, reference.name.casefold())
        groups.setdefault(key, []).append(reference)
    return dict(
        sorted(
            groups.items(),
            key=lambda item: (
                KIND_ORDER.index(item[0][0]) if item[0][0] in KIND_ORDER else 9,
                -len(item[1]),
                item[0][1],
            ),
        )
    )


def report_text(groups: Groups, details: bool = False) -> str:
    """
    Render the grouped references as a plain text report.

    >>> references = extract_references("Import-Module ImportExcel", "a.ps1")
    >>> print(report_text(group_references(references)))
    <BLANKLINE>
    module (1)
      ImportExcel                              1  excel   external  Import-Module
    """
    per_kind = Counter(kind for kind, _ in groups)
    lines: list[str] = []
    current_kind = ""
    for (kind, _), references in groups.items():
        if kind != current_kind:
            current_kind = kind
            lines.append(f"\n{kind} ({per_kind[kind]})")
        first = references[0]
        origins = ", ".join(sorted({reference.origin for reference in references}))
        lines.append(
            f"  {first.name:<38} {len(references):>3}  "
            f"{first.domain or '-':<7} {first.availability:<9} {origins}"
        )
        if details:
            lines.extend(f"      {ref.path}:{ref.line}" for ref in references)
    return "\n".join(lines)


def report_staging(groups: Groups) -> str:
    """
    Say what it takes to run the scanned code on a machine with no internet.

    >>> code = "Import-Module ImportExcel; [xml]$x = Get-Content f.xml"
    >>> print(report_staging(group_references(extract_references(code))))
    <BLANKLINE>
    offline readiness
      builtin     1  ships with Windows, .NET or PowerShell
      external    1  stage it first: Save-Module / nuget install on a connected box
          ImportExcel
    """
    names: dict[str, set[str]] = {}
    for references in groups.values():
        names.setdefault(references[0].availability, set()).add(references[0].name)
    lines = ["\noffline readiness"]
    for availability in AVAILABILITIES:
        found = names.get(availability)
        if not found:
            continue
        advice = STAGING_ADVICE[availability]
        lines.append(f"  {availability:<10} {len(found):>2}  {advice}")
        if availability in {"external", "office", "unknown"}:
            lines.append(f"      {', '.join(sorted(found))}")
    return "\n".join(lines)


def report_json(groups: Groups) -> str:
    """Render the grouped references as JSON."""
    payload = [
        {
            "kind": kind,
            "name": references[0].name,
            "domain": references[0].domain,
            "availability": references[0].availability,
            "dynamic": references[0].dynamic,
            "count": len(references),
            "origins": sorted({reference.origin for reference in references}),
            "occurrences": [
                {"file": reference.path, "line": reference.line}
                for reference in references
            ],
        }
        for (kind, _), references in groups.items()
    ]
    return json.dumps(payload, indent=2)


def report_csv(groups: Groups) -> str:
    """
    Render the grouped references as CSV.

    >>> print(report_csv(group_references(extract_references("using module Foo"))))
    kind,name,domain,availability,dynamic,count,origins,files
    module,Foo,,external,false,1,using module,1
    """
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\n")
    writer.writerow(
        [
            "kind",
            "name",
            "domain",
            "availability",
            "dynamic",
            "count",
            "origins",
            "files",
        ]
    )
    for (kind, _), references in groups.items():
        writer.writerow(
            [
                kind,
                references[0].name,
                references[0].domain,
                references[0].availability,
                str(references[0].dynamic).lower(),
                len(references),
                ";".join(sorted({reference.origin for reference in references})),
                len({reference.path for reference in references}),
            ]
        )
    return buffer.getvalue().rstrip("\n")


def parse_arguments(argv: list[str] | None = None) -> argparse.Namespace:
    """Build the command line interface."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument(
        "paths", nargs="*", default=["."], help="files or directories to scan"
    )
    parser.add_argument(
        "--format", dest="output", choices=("text", "json", "csv"), default="text"
    )
    parser.add_argument(
        "--details", action="store_true", help="list every file and line"
    )
    parser.add_argument(
        "--include-dynamic",
        action="store_true",
        help="also report names that are runtime expressions",
    )
    parser.add_argument(
        "--kinds", default="", help=f"comma separated subset of {','.join(KIND_ORDER)}"
    )
    parser.add_argument(
        "--domains",
        default="",
        help=f"keep only these domains: {','.join(DOMAINS)}",
    )
    parser.add_argument(
        "--suffixes",
        default=",".join(DEFAULT_SUFFIXES),
        help="comma separated file suffixes to scan",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Scan the requested paths and print the report."""
    arguments = parse_arguments(argv)
    suffixes = tuple(
        suffix if suffix.startswith(".") else f".{suffix}"
        for suffix in arguments.suffixes.casefold().split(",")
    )
    references, scanned = scan_files(arguments.paths or ["."], suffixes)
    if not scanned:
        print("no PowerShell files found", file=sys.stderr)
        return 1
    dynamic = sum(1 for reference in references if reference.dynamic)
    if not arguments.include_dynamic:
        references = [reference for reference in references if not reference.dynamic]
    if arguments.kinds:
        wanted = {kind.strip().casefold() for kind in arguments.kinds.split(",")}
        references = [reference for reference in references if reference.kind in wanted]
    if arguments.domains:
        domains = {domain.strip().casefold() for domain in arguments.domains.split(",")}
        references = [
            reference for reference in references if reference.domain in domains
        ]
    groups = group_references(references)
    if arguments.output == "json":
        print(report_json(groups))
        return 0
    if arguments.output == "csv":
        print(report_csv(groups))
        return 0
    by_suffix = ", ".join(
        f"{count} {suffix}" for suffix, count in scanned.most_common()
    )
    skipped = "" if arguments.include_dynamic else f", {dynamic} dynamic skipped"
    print(f"{sum(scanned.values())} files scanned ({by_suffix})")
    print(f"{len(references)} references, {len(groups)} distinct libraries{skipped}")
    print(report_text(groups, arguments.details))
    print(report_staging(groups))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except BrokenPipeError:  # a pager or `head` closed the pipe
        os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
        raise SystemExit(1)
