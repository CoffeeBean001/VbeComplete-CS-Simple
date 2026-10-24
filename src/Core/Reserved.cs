// VBA 保留字 / 内建名：**只用于语法过滤，不再作为候选提示**。
//
// （口径）：候选里只留"自己写的名字"—— 过程名、变量名、参数名、
//   常量名。VBA 关键字、内建函数、内建类型、模块/表/窗体/类名一律不提示。
//   于是这里的两张表退化成"解析时的挡板"，职责只剩一条：
//   收集"没用 Dim 声明过的自由标识符"时，把它们挡掉，别让 `MsgBox "x"`
//   里的 MsgBox 变成一条隐式变量。
//
// 为什么不能删干净：无 Option Explicit 的模块里，`MsgBox "x"`（不带括号）
//   会被当成"当值用过"，如果这里不挡，MsgBox 就会以"变量"的身份进候选池 ——
//   那正是用户明确不要的东西。关键字同理：`Dim`/`As`/`End` 不是变量名。
//
// 枚举常量（vbRed / xlUp / msoTrue…）本来就不在这两张表里，
//   由 Scope 的前缀判据 + 宿主对象名清单单独挡（见 Scope.IsHostEnumConst）。

using System;
using System.Collections.Generic;

namespace VbeComplete
{
    internal static class Reserved
    {
        /// <summary>VBA 语言关键字（79 条）。</summary>
        private static readonly string[] Keywords =
        {
            "Sub", "Function", "Property", "Declare", "PtrSafe", "Alias",
            "Lib", "Public", "Private", "Friend", "Global", "Static",
            "Const", "Dim", "ReDim", "Preserve", "Implements", "Event",
            "WithEvents", "RaiseEvent", "Optional", "ParamArray", "ByVal", "ByRef",
            "Attribute", "Option", "Explicit", "Enum", "Type", "As",
            "New", "Me", "Is", "Like", "Set", "Let",
            "Call", "If", "Then", "Else", "ElseIf", "End",
            "Select", "Case", "For", "Each", "In", "To",
            "Step", "Next", "Do", "Loop", "While", "Wend",
            "Until", "Exit", "GoTo", "GoSub", "Return", "On",
            "Resume", "Stop", "With", "TypeOf", "Rem", "Debug",
            "True", "False", "Nothing", "Empty", "Null", "And",
            "Or", "Not", "Xor", "Eqv", "Imp", "Mod",
            "AddressOf",
        };

        /// <summary>VBA 内建函数（171 条）。</summary>
        private static readonly string[] Functions =
        {
            "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec",
            "CInt", "CLng", "CLngLng", "CLngPtr", "CSng", "CStr",
            "CVar", "CVErr", "IsArray", "IsDate", "IsEmpty", "IsError",
            "IsMissing", "IsNull", "IsNumeric", "IsObject", "TypeName", "VarType",
            "Asc", "AscB", "AscW", "Chr", "ChrB", "ChrW",
            "Format", "FormatCurrency", "FormatDateTime", "FormatNumber", "FormatPercent", "Hex",
            "Oct", "InStr", "InStrRev", "Join", "LCase", "Left",
            "LeftB", "Len", "LenB", "LTrim", "Mid", "MidB",
            "Replace", "Right", "RightB", "RTrim", "Space", "Split",
            "Str", "StrComp", "StrConv", "String", "StrReverse", "Trim",
            "UCase", "Val", "Filter", "Abs", "Atn", "Cos",
            "Exp", "Fix", "Int", "Log", "Rnd", "Round",
            "Sgn", "Sin", "Sqr", "Tan", "Date", "DateAdd",
            "DateDiff", "DatePart", "DateSerial", "DateValue", "Day", "Hour",
            "Minute", "Month", "MonthName", "Now", "Second", "Time",
            "Timer", "TimeSerial", "TimeValue", "Weekday", "WeekdayName", "Year",
            "Array", "LBound", "UBound", "MsgBox", "InputBox", "DoEvents",
            "Shell", "AppActivate", "SendKeys", "Command", "Environ", "CurDir",
            "CallByName", "Choose", "CreateObject", "GetObject", "IIf", "Partition",
            "QBColor", "RGB", "Switch", "Error", "Dir", "EOF",
            "FileAttr", "FileDateTime", "FileLen", "FreeFile", "GetAttr", "SetAttr",
            "FileCopy", "Kill", "MkDir", "RmDir", "ChDir", "ChDrive",
            "Open", "Close", "Get", "Put", "Input", "Seek",
            "Lock", "Unlock", "Loc", "LOF", "Print", "Write",
            "Reset", "LSet", "RSet", "LoadPicture", "SavePicture", "SaveSetting",
            "GetSetting", "GetAllSettings", "DeleteSetting", "DDB", "FV", "IPmt",
            "IRR", "MIRR", "NPer", "NPV", "Pmt", "PPmt",
            "PV", "Rate", "SLN", "SYD", "Load", "Unload",
            "Beep", "Erase", "Randomize",
        };

        /// <summary>VBA 内建数据类型（13 条）。</summary>
        private static readonly string[] TypeNames =
        {
            "Boolean", "Byte", "Currency", "Date", "Double", "Integer",
            "Long", "LongLong", "LongPtr", "Object", "Single", "String",
            "Variant",
        };

        private static readonly HashSet<string> _kw =
            new HashSet<string>(Keywords, StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> _builtin =
            BuildBuiltin();

        private static HashSet<string> BuildBuiltin()
        {
            var set = new HashSet<string>(Functions, StringComparer.OrdinalIgnoreCase);
            foreach (var t in TypeNames) set.Add(t);
            return set;
        }

        /// <summary>是不是语言关键字。</summary>
        public static bool IsReserved(string name) =>
            !string.IsNullOrEmpty(name) && _kw.Contains(name);

        /// <summary>是不是内建函数 / 内建类型名。</summary>
        public static bool IsBuiltinName(string name) =>
            !string.IsNullOrEmpty(name) && _builtin.Contains(name);
    }
}
