Attribute VB_Name = "mdConsts"
'=========================================================================
'
' VbPeg (c) 2018 by wqweto@gmail.com
'
' PEG parser generator for VB6
'
'=========================================================================
Option Explicit
DefObj A-Z

'-- Generates class module with public visibility
Public Const STR_SETTING_PUBLIC         As String = "public"

'-- Generates class module with private visibility
Public Const STR_SETTING_PRIVATE        As String = "private"

'-- Generates class/standard module with this name
Public Const STR_SETTING_MODULENAME     As String = "modulename"

'-- Sets `ctx.UseData` member data-type
Public Const STR_SETTING_USERDATATYPE   As String = "userdatatype"

'-- Sets `ctx.VarResult` and `ctx.VarStack` member data-type
Public Const STR_SETTING_VARDATATYPE    As String = "vardatatype"

'-- Sets grammar start rule. If not set uses first rule
Public Const STR_SETTING_START          As String = "start"

'-- Sets default case-sensitivity matching for the grammar
Public Const STR_SETTING_IGNORECASE     As String = "ignorecase"

'-- Injects prolog (Dim xxx, With yyy) into actions impl function
Public Const STR_SETTING_PROLOG         As String = "prolog"

'-- Injects epilog (Wnd With) into actions impl function
Public Const STR_SETTING_EPILOG         As String = "epilog"

'-- Adds members to `UcsParserType` struct
Public Const STR_SETTING_MEMBERS        As String = "members"

'-- Adds member variables, enums or API declares in the header of the module
Public Const STR_SETTING_DECLARES       As String = "declares"

'-- Sets tracing function to be invoked on rule entry/exit
Public Const STR_SETTING_TRACE          As String = "trace"
