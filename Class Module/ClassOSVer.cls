VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOSVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***************************************************************************
'  Module:        cOperSystem  (clsOperSystem.cls)
'
'  Description:   Determines the operating system.  When this class is
'                 initialized, the GetWindowsVersion() routine is called.
'                 This will instantly set the flags to their appropriate
'                 value.  This module does not test with operating systems
'                 earlier than Windows 2000.
'
'                 Identifying the current operating system is usually not
'                 the best way to determine whether a particular operating
'                 system feature is present. This is because the operating
'                 system may have had new features added in a redistributable
'                 DLL. Rather than using GetVersionEx to determine the
'                 operating system platform or version number, test for the
'                 presence of the feature itself.
'                 http://msdn.microsoft.com/en-us/library/ms724832(VS.85).aspx
'
'                 Applications not manifested specifically for 8.1 or newer
'                 will return the Windows 8.0 operating system values instead
'                 the correct values.
'                 http://msdn.microsoft.com/en-us/library/windows/desktop/ms724834(v=vs.85).aspx
'
'                 Support Statement for Visual Basic 6.0 on Windows Vista,
'                 Windows Server 2008, Windows 7, and Windows 8
'                 http://msdn.microsoft.com/en-us/vstudio/ms788708
'                 "The Visual Basic team’s goal is that Visual Basic 6.0
'                 applications that run on Windows XP will also run on
'                 Windows Vista, Windows Server 2008, Windows 7, and
'                 Windows 8."
'
'                August 2015 - Classic VB6 to be supported at through 2024
'
'                Support Statement for Visual Basic 6.0 on Windows Vista,
'                Windows Server 2008, Windows 7, Windows 8 and Windows 8.1,
'                Windows Server 2012, and Windows 10
'                https://msdn.microsoft.com/en-us/vstudio/ms788708.aspx
'
'                Windows XP support has ended as of April 8, 2014
'                http://windows.microsoft.com/en-us/windows/end-support-help
'
'                Windows Server 2003 extended support ended on July 14, 2015
'                https://www.microsoft.com/en-us/server-cloud/products/windows-server-2003/
'
' AddIn tools    Callers Add-in v2.9 by RD Edwards (RDE)
' for VB6:       Fantastic VB6 add-in to indentify if a routine calls
'                another routine or is called by other routines within
'                a project.  A must have tool for any VB6 programmer.
'                http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=74734&lngWId=1
'
'                NOTE:  Under Windows 10, if you have problems recognizing
'                a VB6 addin, try recompiling it directly into the System32
'                folder.
'
' -----------------------------------------------------------------------
' Base version information for Windows versions:
'
' Version                     Number  Major  Minor  Build   Other info
'  -----------------------------------------------------------------------
' Windows Server 2016         10.0     10      0   >10000   OSVERSIONINFOEXW.wProductType <> VER_NT_WORKSTATION
' Windows 10                  10.0     10      0   >10000   OSVERSIONINFOEXW.wProductType = VER_NT_WORKSTATION
' Windows Server 10 (Beta)     6.4      6      4    >9000   OSVERSIONINFOEXW.wProductType <> VER_NT_WORKSTATION
' Windows 10 (Beta)            6.4      6      4    >9000   OSVERSIONINFOEXW.wProductType = VER_NT_WORKSTATION
' Windows Server 2012 R2       6.3      6      3     9600   OSVERSIONINFOEXW.wProductType <> VER_NT_WORKSTATION
' Windows 8.1                  6.3      6      3     9600   OSVERSIONINFOEXW.wProductType = VER_NT_WORKSTATION
' Windows Server 2012          6.2      6      2   >=9200   OSVERSIONINFOEXW.wProductType <> VER_NT_WORKSTATION
' Windows 8                    6.2      6      2     9200   OSVERIONINFOEX.wProductType = VER_NT_WORKSTATION
' Windows Server 2008 R2       6.1      6      1     7601   OSVERSIONINFOEXW.wProductType <> VER_NT_WORKSTATION
' Windows 7                    6.1      6      1     7601   OSVERSIONINFOEXW.wProductType = VER_NT_WORKSTATION
' Windows Server 2008          6.0      6      0     6001   OSVERSIONINFOEXW.wProductType <> VER_NT_WORKSTATION
' Windows Vista                6.0      6      0     6000   OSVERSIONINFOEXW.wProductType = VER_NT_WORKSTATION
' Windows Server 2003 R2       5.2      5      2     3790   GetSystemMetrics(SM_SERVERR2) <> 0
' Windows Server 2003          5.2      5      2     3790   GetSystemMetrics(SM_SERVERR2) = 0
' Windows XP Home Server       5.2      5      2     3790   OSVERSIONINFOEXW.wSuiteMask = VER_SUITE_WH_SERVER
' Windows XP Pro x64 Edition   5.2      5      2     3790   (OSVERSIONINFOEXW.wProductType = VER_NT_WORKSTATION) And
'                                                           ProcessorArchitecture = 64-bit
' Windows XP                   5.1      5      1     2600
' Windows 2000                 5.0      5      0     2195
'  -----------------------------------------------------------------------
'
' Reference:   Randy Birch
'              GetVersionEx: Windows Version Info (Wrapper Routines)
'              http://vbnet.mvps.org/code/helpers/iswinversion.htm
'
'              Karl Peterson's Classic VB
'              http://www.mvps.org/vb/
'
'              GetVersionEx Function
'              http://msdn2.microsoft.com/en-us/library/ms724451(VS.85).aspx
'
'              Getting the System Version (Source code)
'              http://msdn2.microsoft.com/en-us/library/ms724429(VS.85).aspx
'
'              Operating System Version
'              http://msdn.microsoft.com/en-us/library/windows/desktop/ms724832%28v=vs.85%29.aspx
'
'              GetSystemMetrics Function
'              http://msdn2.microsoft.com/en-us/library/ms724385(VS.85).aspx
'
'              OSVERSIONINFOEXW structure
'              http://msdn.microsoft.com/en-us/library/windows/desktop/ms724833(v=vs.85).aspx
'
'              Win32_OperatingSystem class
'              https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
'
'              Operating System Property Values
'              https://msdn.microsoft.com/en-us/library/aa370556(v=vs.85).aspx
'
'              List of Microsoft Windows versions
'              https://en.wikipedia.org/wiki/List_of_Microsoft_Windows_versions
'
'              The 2016 Microsoft Product Roadmap
'              https://rcpmag.com/articles/2011/02/01/the-2011-microsoft-product-roadmap.aspx
'
' Important:   Developers working in the Windows Vista or newer environment
'              should be aware that if this code is run in the VB IDE and
'              you are notified that the operating system is other than
'              the correct Windows version (i.e., dwVerMajor=5 instead of
'              dwVerMajor=6), do the following:
'
'              1.  Close the VB IDE (Visual BASIC Integrated Development Environment)
'              2.  Locate the file VB.EXE
'                  ex:  C:\Program Files\Microsoft Visual Studio\VB98\VB.EXE or
'                       C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB.EXE
'              3.  Right mouse click and select Properties >> Compatibility tab
'              4.  Remove compatibility checkmark and close the window
'              5.  Restart VB IDE
'
'  -----------------------------------------------------------------------
'
'  Example:    Set objOS = new cOperSystem  ' Instantiate class object
'
'              ' Is this Windows 10?
'              If objOS.bWindows10 then
'                  Msgbox objOS.VersionData & vbNewLine & _
'                         "Ver " & objOS.VersionNumber
'              Else
'                  Msgbox "This is not Windows 10"
'              End If
'
'              Set objOS = Nothing  ' Free class object from memory
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Oct-2006  Kenneth Ives  kenaso@tx.rr.com
'              Converted to a class, renamed some of the variables,
'              Combined some of the routines for speed.  When this class is
'              initialized, the GetWindowsVersion() routine is called.  This
'              will instantly set the various flags to an appropriate value.
' 05-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Maria AAM for pointing out that I needed to add an
'              additional property named bOperSystem64() denoting a 64-bit
'              operating system in general.
' 12-Jul-2015  Kenneth Ives  kenaso@tx.rr.com
'              Based on input from Wieslaw Smagur, I was inspired to make
'              this module more efficient so I decided to rewrite it using
'              WMI and Visual Basic to get a more accurate description of
'              the operating system.
' 09-Aug-2015  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Domenico Galliano for his valuable input concerning
'              the XP portion and its testing.
' 12-Aug-2015  Kenneth Ives  kenaso@tx.rr.com
'              - Renamed TestMajorVer_6() to TestOtherVer() to accomodate
'                Testing for Windows 10.
'              - Added testing for Windows 10 Beta (pre-release) version.
'              - Renamed RefineDataSearch() to RefineNonServerData().
'              - Split routine RefineNonServerData() into two routines.
'                One for workstations [RefineNonServerData()] and
'                another for servers [RefineServerData()].
' 16-Aug-2015  Kenneth Ives  kenaso@tx.rr.com
'              Added testing for Windows 10 Beta (pre-release) version.
' 10-Jan-2016  Kenneth Ives  kenaso@tx.rr.com
'              - Separated Windows 10 (Official release) testing
'              - Renamed TestOtherVer() to TestMajorVer_6()
'              - Created TestMajorVer_10() routine
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  ' Used by OSVERSIONINFOEXW.PlatformID
  Private Const VER_PLATFORM_WIN32_NT                  As Long = 2

  ' used by OSVERSIONINFOEXW.wProductType
  Private Const VER_NT_WORKSTATION                     As Long = &H1         ' Windows 10, Windows 8.1, Windows 8, Windows 7, Windows Vista,
                                                                             ' Windows XP Pro, Windows XP Home Edition, or Windows 2000 Pro
  Private Const VER_NT_DOMAIN_CONTROLLER               As Long = &H2         ' Windows Server 10, Windows Server 2012, Windows Server 2008 R2,
                                                                             ' Windows Server 2008, Windows Server 2003, or Windows 2000 Server
  Private Const VER_NT_SERVER                          As Long = &H3         ' Windows Server 10, Windows Server 2012, Windows Server 2008 R2,
                                                                             ' Windows Server 2008, Windows Server 2003, or Windows 2000 Server
                                                                                  
  ' used by OSVERSIONINFOEXW.wSuiteMask
  Private Const VER_SUITE_SMALLBUSINESS                As Long = &H1         ' Windows Server 2003: Small Business
  Private Const VER_SUITE_ENTERPRISE                   As Long = &H2         ' Windows Server 2003: Enterprise Edition or
                                                                             ' Windows Server 2008: Enterprise or
                                                                             ' Windows 2000 Advanced Server
  Private Const VER_SUITE_COMMUNICATION                As Long = &H3         ' Windows Server 2003: Communication
  Private Const VER_SUITE_BACKOFFICE                   As Long = &H4         ' Windows Server 2003: BackOffice
  Private Const VER_SUITE_TERMINAL                     As Long = &H10        ' Windows Server 2003: Terminal
  Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED     As Long = &H20        ' Windows Server 2003: Small Business Restricted
  Private Const VER_SUITE_EMBEDDEDNT                   As Long = &H40        ' Windows XP Embedded
  Private Const VER_SUITE_DATACENTER                   As Long = &H80        ' Windows Server 2003: DataCenter or
                                                                             ' Windows Server 2008: Datacenter or
                                                                             ' Windows 2000 Datacenter Server
  Private Const VER_SUITE_SINGLEUSERTS                 As Long = &H100       ' Remote Desktop is supported
  Private Const VER_SUITE_PERSONAL                     As Long = &H200       ' Windows Vista Home Premium or
                                                                             ' Windows Vista Home Basic or
                                                                             ' Windows XP Home Edition
  Private Const VER_SUITE_BLADE                        As Long = &H400       ' Windows Server 2003: Web Edition
  Private Const VER_SUITE_STORAGE_SERVER               As Long = &H2000      ' Windows Storage Server 2003 R2 or
                                                                             ' Windows Storage Server 2003
  Private Const VER_SUITE_COMPUTE_SERVER               As Long = &H4000      ' Windows Compute Cluster Edition Server
  Private Const VER_SUITE_WH_SERVER                    As Long = &H8000      ' Windows Home Server

  ' Used by API GetSystemMetrics()
  Private Const SM_SERVERR2                            As Long = 89          ' Windows Server 2003 R2

  ' Used to refine product type
  Private Const PRODUCT_UNDEFINED                      As Long = &H0         '  0-An unknown product
  Private Const PRODUCT_ULTIMATE                       As Long = &H1         '  1-Ultimate
  Private Const PRODUCT_HOME_BASIC                     As Long = &H2         '  2-Home Basic
  Private Const PRODUCT_HOME_PREMIUM                   As Long = &H3         '  3-Home Premium
  Private Const PRODUCT_ENTERPRISE                     As Long = &H4         '  4-Enterprise
  Private Const PRODUCT_HOME_BASIC_N                   As Long = &H5         '  5-Home Basic N (without Windows Media Player)
  Private Const PRODUCT_BUSINESS                       As Long = &H6         '  6-Business
  Private Const PRODUCT_STANDARD_SERVER                As Long = &H7         '  7-Server Standard
  Private Const PRODUCT_DATACENTER_SERVER              As Long = &H8         '  8-Server Datacenter (full installation)
  Private Const PRODUCT_SMALLBUSINESS_SERVER           As Long = &H9         '  9-Windows Small Business Server
  Private Const PRODUCT_ENTERPRISE_SERVER              As Long = &HA         ' 10-Server Enterprise (full installation)
  Private Const PRODUCT_STARTER                        As Long = &HB         ' 11-Starter
  Private Const PRODUCT_DATACENTER_SERVER_CORE         As Long = &HC         ' 12-Server Datacenter (core installation)
  Private Const PRODUCT_STANDARD_SERVER_CORE           As Long = &HD         ' 13-Server Standard (core installation)
  Private Const PRODUCT_ENTERPRISE_SERVER_CORE         As Long = &HE         ' 14-Server Enterprise (core installation)
  Private Const PRODUCT_ENTERPRISE_SERVER_ITANIUM      As Long = &HF         ' 15-Enterprise Server Edition for Itanium-Based Systems
  Private Const PRODUCT_BUSINESS_N                     As Long = &H10        ' 16-Business N (without Windows Media Player)
  Private Const PRODUCT_WEB_SERVER                     As Long = &H11        ' 17-Web Server (full installation)
  Private Const PRODUCT_CLUSTER_SERVER                 As Long = &H12        ' 18-Cluster Server
  Private Const PRODUCT_STORAGE_SERVER_2008_R2_ESS     As Long = &H13        ' 19-Windows Storage Server 2008 R2 Essentials
  Private Const PRODUCT_STORAGE_SERVER_EXPRESS         As Long = &H14        ' 20-Storage Server Express
  Private Const PRODUCT_STORAGE_SERVER_STANDARD        As Long = &H15        ' 21-Storage Server Standard
  Private Const PRODUCT_STORAGE_SERVER_WORKGROUP       As Long = &H16        ' 22-Storage Server Workgroup
  Private Const PRODUCT_STORAGE_SERVER_ENTERPRISE      As Long = &H17        ' 23-Storage Server Enterprise
  Private Const PRODUCT_SERVER_2008_ESS_SVCS           As Long = &H18        ' 24-Windows Server 2008 for Windows Essential Server Solutions
  Private Const PRODUCT_SMALL_BUS_SERVER_PREMIUM       As Long = &H19        ' 25-Small Business Server Premium
  Private Const PRODUCT_HOME_PREMIUM_N                 As Long = &H1A        ' 26-Home Premium N (without Windows Media Player)
  Private Const PRODUCT_ENTERPRISE_N                   As Long = &H1B        ' 27-Enterprise N (without Windows Media Player)
  Private Const PRODUCT_ULTIMATE_N                     As Long = &H1C        ' 28-Ultimate N (without Windows Media Player)
  Private Const PRODUCT_WEB_SERVER_CORE                As Long = &H1D        ' 29-Web Server (core installation)
  Private Const PRODUCT_ESS_BUS_SERVER_MANAGEMENT      As Long = &H1E        ' 30-Windows Essential Business Server Management Server
  Private Const PRODUCT_ESS_BUS_SERVER_SECURITY        As Long = &H1F        ' 31-Windows Essential Business Server Security Server
  Private Const PRODUCT_ESS_BUS_SERVER_MESSAGING       As Long = &H20        ' 32-Windows Essential Business Server Messaging Server
  Private Const PRODUCT_SERVER_FOUNDATION              As Long = &H21        ' 33-Server Foundation
  Private Const PRODUCT_HOME_SERVER_2011               As Long = &H22        ' 34-Windows Home Server 2011
  Private Const PRODUCT_SERVER_2008_WO_V               As Long = &H23        ' 35-Windows Server 2008 without Hyper-V for Windows Essential Server Solutions
  Private Const PRODUCT_STANDARD_SERVER_WO_V           As Long = &H24        ' 36-Server Standard without Hyper-V
  Private Const PRODUCT_DATACENTER_SERVER_WO_V         As Long = &H25        ' 37-Server Datacenter without Hyper-V (core installation)
  Private Const PRODUCT_ENTERPRISE_SVR_WO_V            As Long = &H26        ' 38-Server Enterprise without Hyper-V (core installation)
  Private Const PRODUCT_DATACENTER_SVR_CORE_WO_V       As Long = &H27        ' 39-Server Datacenter without Hyper-V (core installation)
  Private Const PRODUCT_STANDARD_SVR_CORE_WO_V         As Long = &H28        ' 40-Server Standard without Hyper-V (core installation)
  Private Const PRODUCT_ENTERPRISE_SVR_CORE_WO_V       As Long = &H29        ' 41-Server Enterprise without Hyper-V (core installation)
  Private Const PRODUCT_HYPER_V                        As Long = &H2A        ' 42-Microsoft Hyper-V Server
  Private Const PRODUCT_STORAGE_SERVER_EXPRESS_CORE    As Long = &H2B        ' 43-Storage Server Express (core installation)
  Private Const PRODUCT_STORAGE_SERVER_STANDARD_CORE   As Long = &H2C        ' 44-Storage Server Standard (core installation)
  Private Const PRODUCT_STORAGE_SERVER_WORKGROUP_CORE  As Long = &H2D        ' 45-Storage Server Workgroup (core installation)
  Private Const PRODUCT_STORAGE_SERVER_ENTERPRISE_CORE As Long = &H2E        ' 46-Storage Server Enterprise (core installation)
  Private Const PRODUCT_STARTER_N                      As Long = &H2F        ' 47-Starter N
  Private Const PRODUCT_PROFESSIONAL                   As Long = &H30        ' 48-Professional
  Private Const PRODUCT_PROFESSIONAL_N                 As Long = &H31        ' 49-Professional N
  Private Const PRODUCT_SB_SERVER_2011_ESS             As Long = &H32        ' 50-Windows Small Business Server 2011 Essentials
  Private Const PRODUCT_SERVER_FOR_SB_SOLUTIONS        As Long = &H33        ' 51-Server For SB Solutions
  Private Const PRODUCT_SERVER_SOLUTIONS_PREM          As Long = &H34        ' 52-Server Solutions Premium
  Private Const PRODUCT_SERVER_SOLUTIONS_PREM_CORE     As Long = &H35        ' 53-Server Solutions Premium (core installation)
  Private Const PRODUCT_SERVER_SB_SOLUTIONS_EM_1       As Long = &H36        ' 54-Server For SB Solutions EM
  Private Const PRODUCT_SERVER_SB_SOLUTIONS_EM_2       As Long = &H37        ' 55-Server For SB Solutions EM
  Private Const PRODUCT_MULTIPOINT_SERVER              As Long = &H38        ' 56-Windows MultiPoint Server
  Private Const PRODUCT_ESS_SERVER_SOLUTION_MGMT       As Long = &H3B        ' 59-Windows Essential Server Solution Management
  Private Const PRODUCT_ESS_SERVER_SOLUTION_ADDL       As Long = &H3C        ' 60-Windows Essential Server Solution Additional
  Private Const PRODUCT_ESS_SERVER_SOLUTION_MGMT_SVC   As Long = &H3D        ' 61-Windows Essential Server Solution Management Services
  Private Const PRODUCT_ESS_SERVER_SOLUTION_ADDL_SVC   As Long = &H3E        ' 62-Windows Essential Server Solution Additional Services
  Private Const PRODUCT_SMALL_BUS_SERVER_PREMIUM_CORE  As Long = &H3F        ' 63-Small Business Server Premium (core installation)
  Private Const PRODUCT_SERVER_HYPER_V                 As Long = &H40        ' 64-Server Hyper Core V
  Private Const PRODUCT_STARTER_E                      As Long = &H42        ' 66-Starter E
  Private Const PRODUCT_HOME_BASIC_E                   As Long = &H43        ' 67-Home Basic E
  Private Const PRODUCT_HOME_PREMIUM_E                 As Long = &H44        ' 68-Home Premium E
  Private Const PRODUCT_PROFESSIONAL_E                 As Long = &H45        ' 69-Professional E
  Private Const PRODUCT_ENTERPRISE_E                   As Long = &H46        ' 70-Enterprise E
  Private Const PRODUCT_ULTIMATE_E                     As Long = &H47        ' 71-Ultimate E
  Private Const PRODUCT_SVR_ENTERPRISE_EVAL            As Long = &H48        ' 72-Server Enterprise (evaluation installation)
  Private Const PRODUCT_MULTIPOINT_SERVER_STANDARD     As Long = &H4C        ' 76-Windows MultiPoint Server Standard (full installation)
  Private Const PRODUCT_MULTIPOINT_SERVER_PREMIUM      As Long = &H4D        ' 77-Windows MultiPoint Server Premium (full installation)
  Private Const PRODUCT_SERVER_STANDARD_EVAL           As Long = &H4F        ' 79-Server Standard (evaluation installation)
  Private Const PRODUCT_SERVER_DATACENTER_EVAL         As Long = &H50        ' 80-Server Datacenter (evaluation installation)
  Private Const PRODUCT_ENTERPRISE_N_EVAL              As Long = &H54        ' 84-Enterprise N (evaluation installation)
  Private Const PRODUCT_STORAGE_SERVER_WORKGROUP_EVAL  As Long = &H5F        ' 95-Storage Server Workgroup (evaluation installation)
  Private Const PRODUCT_STORAGE_SERVER_STANDARD_EVAL   As Long = &H60        ' 96-Storage Server Standard (evaluation installation)
  Private Const PRODUCT_WINDOWS_8_N                    As Long = &H62        ' 98-Windows 8 N
  Private Const PRODUCT_WINDOWS_8_SNG_LANG             As Long = &H64        ' 100-Windows 8 Single Language
  Private Const PRODUCT_WINDOWS_EDITION                As Long = &H65        ' 101-Home Edition
  Private Const PRODUCT_PROFESSIONAL_WMC               As Long = &H67        ' 103-Professional with Media Center
  Private Const PRODUCT_UNLICENSED                     As Long = &HABCDABCD  ' Unlicensed product

' ***************************************************************************
' Type structures
' ***************************************************************************
  ' The OSVERSIONINFOEXW data structure contains operating system version
  ' information. The information includes major and minor version numbers,
  ' a build number, a platform identifier, and information about product
  ' suites and the latest Service Pack installed on the system. This structure
  ' is used with the GetVersionEx and VerifyVersionInfo functions.
  Private Type OSVERSIONINFOEXW
      OSVSize           As Long           ' size of this data structure (in bytes)
      dwVerMajor        As Long           ' major version number of the operating system (ex: 5)
      dwVerMinor        As Long           ' minor version number of the operating system (ex: 01)
      dwBuildNumber     As Long           ' build number of the operating system (ex: 2600)
      PlatformID        As Long           ' Identifies operating system platform
      szCSDVersion      As String * 128   ' indicates the latest Service Pack installed on the system (ex: "Service Pack 3")
      wServicePackMajor As Integer        ' major version number of the latest Service Pack installed on the system
      wServicePackMinor As Integer        ' minor version number of the latest Service Pack installed on the system
      wSuiteMask        As Integer        ' bit mask that identifies the product suites available on the system
      wProductType      As Byte           ' additional information about the system
      wReserved         As Byte           ' Reserved for future use
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' This function obtains extended information about the version of the
  ' operating system that is currently running.
  Private Declare Function GetSystemMetrics Lib "user32" _
          (ByVal nIndex As Long) As Long

  ' This function obtains extended information about the version of the
  ' operating system that is currently running.
  Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
          (lpVersionInformation As Any) As Boolean

  ' ZeroMemory function fills a block of memory with zeros.
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Boolean)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m bln WindowsNT
' Variable name:     mblnWindowsNT
' ***************************************************************************
  Private mlngMajorVer                    As Long
  Private mlngMinorVer                    As Long
  Private mlngBuildNumber                 As Long
  Private mintSvrPackMajor                As Integer
  Private mintSvrPackMinor                As Integer
  Private mstrAppend                      As String
  Private mstrVerName                     As String
  Private mstrProcArch                    As String
  Private mstrBitStruct                   As String
  Private mstrVerNumber                   As String
  Private mblnCaptionIsCentered           As Boolean    ' Used by calling application to see if
                                                        ' form captions are centered automatically
  
  Private mtypOSVerInfoW                  As OSVERSIONINFOEXW

  ' Generalized flags
  Private mblnMajorVersion6               As Boolean
  Private mblnMajorVersion10              As Boolean
  Private mblnUndefinedProduct            As Boolean
  Private mblnUnlicensedProduct           As Boolean
  Private mblnWin2000orNewer              As Boolean
  Private mblnWin2000_SP1orNewer          As Boolean
  Private mblnWin2000_SP2orNewer          As Boolean
  Private mblnWin2000_SP3orNewer          As Boolean
  Private mblnWin2000_SP4orNewer          As Boolean
  Private mblnWinVistaOrNewer             As Boolean
  Private mblnWinVista_SP1orNewer         As Boolean
  Private mblnWinVista_SP2orNewer         As Boolean
  Private mblnWinXPorNewer                As Boolean
  Private mblnWinXP_SP1orNewer            As Boolean
  Private mblnWinXP_SP2orNewer            As Boolean
  Private mblnWinXP_SP3orNewer            As Boolean
  Private mblnWindows7orNewer             As Boolean
  Private mblnWindows7_SP1orNewer         As Boolean
  Private mblnWindows8orNewer             As Boolean
  Private mblnWindows81orNewer            As Boolean
  Private mblnWindows10orNewer            As Boolean
  Private mblnWindowsNT                   As Boolean

  ' Service packs
  Private mblnServicePack1                As Boolean
  Private mblnServicePack2                As Boolean
  Private mblnServicePack3                As Boolean
  Private mblnServicePack4                As Boolean
                                     
  ' Base operating systems
  Private mblnWin2000                     As Boolean
  Private mblnWinVista                    As Boolean
  Private mblnWinXP                       As Boolean
  Private mblnWindows7                    As Boolean
  Private mblnWindows8                    As Boolean
  Private mblnWindows81                   As Boolean
  Private mblnWindows10                   As Boolean
  Private mblnWindows10_Beta              As Boolean

  Private mblnWin2000_Pro                 As Boolean
  Private mblnWinXP_Pro                   As Boolean
  Private mblnWinVista_Pro                As Boolean
  Private mblnWindows7_Pro                As Boolean
  Private mblnWindows8_Pro                As Boolean
  Private mblnWindows81_Pro               As Boolean
  Private mblnWindows10_Pro_Beta          As Boolean
  Private mblnWindows10_Pro               As Boolean

  ' Workstations
  Private mblnBusinessEd                  As Boolean
  Private mblnXPEmbedded                  As Boolean
  Private mblnEnterprise                  As Boolean
  Private mblnHomeBasic                   As Boolean
  Private mblnHomePremium                 As Boolean
  Private mblnMediaCenter                 As Boolean
  Private mblnStarterEd                   As Boolean
  Private mblnProfessional                As Boolean
  Private mblnUltimateEd                  As Boolean
  Private mblnWorkstation                 As Boolean
  Private mblnSingleUser                  As Boolean
  Private mblnXPHomeEd                    As Boolean
  Private mblnWindows8SL                  As Boolean
  
  ' "_E" versions are meant for European Commission countries,
  ' including UK, and includes the same functionality as
  ' ordinary standard flavor of Windows 7, except that it does
  ' not include Internet Explorer 8 (IE8).
  Private mblnUltimateEd_E                As Boolean
  Private mblnHomeBasic_E                 As Boolean
  Private mblnHomePremium_E               As Boolean
  Private mblnEnterprise_E                As Boolean
  Private mblnStarterEd_E                 As Boolean
  Private mblnProfessional_E              As Boolean
  
  ' "_N" versions are meant for European market, and include
  ' the same functionality as Windows, except that it does
  ' not include Windows Media Player and related technologies
  ' such as Windows Movie Maker.
  Private mblnUltimateEd_N                As Boolean
  Private mblnHomeBasic_N                 As Boolean
  Private mblnHomePremium_N               As Boolean
  Private mblnEnterprise_N                As Boolean
  Private mblnBusinessEd_N                As Boolean
  Private mblnStarterEd_N                 As Boolean
  Private mblnProfessional_N              As Boolean
  Private mblnWindows8_N                  As Boolean

  ' Servers
  Private mblnProfessional_WMC            As Boolean
  Private mblnBackOfficeServer            As Boolean
  Private mblnBackOfficeTermSvr           As Boolean
  Private mblnCommServer                  As Boolean
  Private mblnClusterServer               As Boolean
  Private mblnDatacenterServer            As Boolean
  Private mblnDatacenterServerCore        As Boolean
  Private mblnDomainController            As Boolean
  Private mblnEnterpriseServer            As Boolean
  Private mblnEnterpriseServerCore        As Boolean
  Private mblnEssBusServerSolAddl         As Boolean
  Private mblnEssBusServerSolAddlSvc      As Boolean
  Private mblnEssBusServerSolMgmt         As Boolean
  Private mblnEssBusServerSolMgmtSvc      As Boolean
  Private mblnHomePremiumServer           As Boolean
  Private mblnHomeEdServer                As Boolean
  Private mblnMedBusServerMgmt            As Boolean
  Private mblnMedBusServerMessaging       As Boolean
  Private mblnMedBusServerSecurity        As Boolean
  Private mblnMultiPointServer            As Boolean
  Private mblnSB_SolutionServer2011       As Boolean
  Private mblnServerForSB_Solution        As Boolean
  Private mblnServerForSB_Solution_EM     As Boolean
  Private mblnServerFoundation            As Boolean
  Private mblnSmallBusRestrictedServer    As Boolean
  Private mblnSmallBusServer              As Boolean
  Private mblnSmallBusServerPremCore      As Boolean
  Private mblnSmallBusServerPremium       As Boolean
  Private mblnStandardServer              As Boolean
  Private mblnStandardServerCore          As Boolean
  Private mblnStandardServerSolution      As Boolean
  Private mblnStandardSvrSolutionCore     As Boolean
  Private mblnSolutionSvrForSB_Sol_EM     As Boolean
  Private mblnStorageServer               As Boolean
  Private mblnTerminalServer              As Boolean
  Private mblnWebServer                   As Boolean
  Private mblnWebServerCore               As Boolean
  Private mblnWin2000Server               As Boolean
  Private mblnWinServer2003               As Boolean
  Private mblnWinServer2003R2             As Boolean
  Private mblnWinServer2008               As Boolean
  Private mblnWinServer2008R2             As Boolean
  Private mblnWinServer2012               As Boolean
  Private mblnWinServer2012R2             As Boolean
  Private mblnWinServer2016               As Boolean
  Private mblnWinServer10_Beta            As Boolean
  Private mblnWinXPServer                 As Boolean
  Private mblnWindowsServer               As Boolean

  ' Hyper-V based
  Private mblnClusterServer_V             As Boolean
  Private mblnDatacenterServerCore_V      As Boolean
  Private mblnDatacenterServer_V          As Boolean
  Private mblnEnterpriseServerCore_V      As Boolean
  Private mblnEnterpriseServer_V          As Boolean
  Private mblnServerForSmallBus_V         As Boolean
  Private mblnStandardServerCore_V        As Boolean
  Private mblnStandardServer_V            As Boolean
    
  ' Storage servers
  Private mblnStorageEnterpriseServer     As Boolean
  Private mblnStorageEnterpriseServerCore As Boolean
  Private mblnStorageExpressServer        As Boolean
  Private mblnStorageExpressServerCore    As Boolean
  Private mblnStorageServer2008R2         As Boolean
  Private mblnStorageStandardServer       As Boolean
  Private mblnStorageStandardServerCore   As Boolean
  Private mblnStorageWorkgrpServer        As Boolean
  Private mblnStorageWorkgrpServerCore    As Boolean
        
  ' Evaluation servers
  Private mblnDatacenterServerEval        As Boolean
  Private mblnEnterpriseServerEval        As Boolean
  Private mblnEnterpriseServerN_Eval      As Boolean
  Private mblnMultiPointPremiumServerEval As Boolean
  Private mblnMultiPointStandardServer    As Boolean
  Private mblnStandardServerEval          As Boolean
  Private mblnStorageStandardServerEval   As Boolean
  Private mblnStorageWorkgroupServerEval  As Boolean
 
  ' 64-Bit Versions
  Private mblnClusterServer64             As Boolean
  Private mblnDatacenterServer64          As Boolean
  Private mblnEnterpriseItanium64         As Boolean
  Private mblnEnterpriseServer64          As Boolean
  Private mblnOperSystem64                As Boolean
  Private mblnStandardServer64            As Boolean
  Private mblnWinXP_Pro64                 As Boolean
  Private mblnWinVista64                  As Boolean


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Get BaseBitStructure() As String
    BaseBitStructure = mstrBitStruct   ' 32-bit | 64-bit
End Property

Public Property Get BuildNumber() As String
    BuildNumber = CStr(mlngBuildNumber)
End Property

Public Property Get MajorVersion() As String
    MajorVersion = CStr(mlngMajorVer)
End Property

Public Property Get MinorVersion() As String
    MinorVersion = CStr(mlngMinorVer)
End Property

Public Property Get PlatformID() As Long
    PlatformID = mtypOSVerInfoW.PlatformID
End Property

Public Property Get ProcessArchitecture() As String
    ProcessArchitecture = mstrProcArch   ' x86 | IA64 | AMD64
End Property

Public Property Get ServicePack() As String
    
    Dim strData As String
    
    strData = vbNullString
    
    If mintSvrPackMajor > 0 Then
        
        strData = " SP" & CStr(mintSvrPackMajor)
        
        If mintSvrPackMinor > 0 Then
            strData = strData & "." & CStr(mintSvrPackMinor)
        End If
        
    End If
    
    ServicePack = strData
    
End Property

Public Property Get VersionName() As String
    VersionName = mstrVerName
End Property

Public Property Get VersionNumber() As String
    VersionNumber = mstrVerNumber
End Property

Public Property Get VersionData() As String
    VersionData = TrimStr(VersionNumber & ServicePack)
End Property

' Generalized flags
Public Property Get bCaptionIsCentered() As Boolean
    bCaptionIsCentered = mblnCaptionIsCentered
End Property

Public Property Get bMajorVersion6() As Boolean
    bMajorVersion6 = mblnMajorVersion6
End Property

Public Property Get bMajorVersion10() As Boolean
    bMajorVersion10 = mblnMajorVersion10
End Property

' Service packs
Public Property Get bServicePack1() As Boolean
    bServicePack1 = mblnServicePack1
End Property

Public Property Get bServicePack2() As Boolean
    bServicePack2 = mblnServicePack2
End Property

Public Property Get bServicePack3() As Boolean
    bServicePack3 = mblnServicePack3
End Property

Public Property Get bServicePack4() As Boolean
    bServicePack4 = mblnServicePack4
End Property

Public Property Get bUndefinedProduct() As Boolean
    bUndefinedProduct = mblnUndefinedProduct
End Property

Public Property Get bUnlicensedProduct() As Boolean
    bUnlicensedProduct = mblnUnlicensedProduct
End Property

' Major version 5
Public Property Get bWin2000() As Boolean
    bWin2000 = mblnWin2000
End Property

Public Property Get bWin2000_Pro() As Boolean
    bWin2000_Pro = mblnWin2000_Pro
End Property

Public Property Get bWin2000orNewer() As Boolean
    bWin2000orNewer = mblnWin2000orNewer
End Property

Public Property Get bWin2000_SP1orNewer() As Boolean
    bWin2000_SP1orNewer = mblnWin2000_SP1orNewer
End Property

Public Property Get bWin2000_SP2orNewer() As Boolean
    bWin2000_SP2orNewer = mblnWin2000_SP2orNewer
End Property

Public Property Get bWin2000_SP3orNewer() As Boolean
    bWin2000_SP3orNewer = mblnWin2000_SP3orNewer
End Property

Public Property Get bWin2000_SP4orNewer() As Boolean
    bWin2000_SP4orNewer = mblnWin2000_SP4orNewer
End Property

Public Property Get bWinXP() As Boolean
    bWinXP = mblnWinXP
End Property

Public Property Get bWinXP_Pro() As Boolean
    bWinXP_Pro = mblnWinXP_Pro
End Property

Public Property Get bWinXPorNewer() As Boolean
    bWinXPorNewer = mblnWinXPorNewer
End Property

Public Property Get bWinXP_SP1orNewer() As Boolean
    bWinXP_SP1orNewer = mblnWinXP_SP1orNewer
End Property

Public Property Get bWinXP_SP2orNewer() As Boolean
    bWinXP_SP2orNewer = mblnWinXP_SP2orNewer
End Property

Public Property Get bWinXP_SP3orNewer() As Boolean
    bWinXP_SP3orNewer = mblnWinXP_SP3orNewer
End Property

' Major version 6
Public Property Get bWinVista() As Boolean
    bWinVista = mblnWinVista
End Property

Public Property Get bWinVista_Pro() As Boolean
    bWinVista_Pro = mblnWinVista_Pro
End Property

Public Property Get bWinVistaOrNewer() As Boolean
    bWinVistaOrNewer = mblnWinVistaOrNewer
End Property

Public Property Get bWinVista_SP1orNewer() As Boolean
    bWinVista_SP1orNewer = mblnWinVista_SP1orNewer
End Property

Public Property Get bWinVista_SP2orNewer() As Boolean
    bWinVista_SP2orNewer = mblnWinVista_SP2orNewer
End Property

Public Property Get bWindows7() As Boolean
    bWindows7 = mblnWindows7
End Property

Public Property Get bWindows7_Pro() As Boolean
    bWindows7_Pro = mblnWindows7_Pro
End Property

Public Property Get bWindows7orNewer() As Boolean
    bWindows7orNewer = mblnWindows7orNewer
End Property

Public Property Get bWindows7SP1orNewer() As Boolean
    bWindows7SP1orNewer = mblnWindows7_SP1orNewer
End Property

Public Property Get bWindows8() As Boolean
    bWindows8 = mblnWindows8
End Property

Public Property Get bWindows8_Pro() As Boolean
    bWindows8_Pro = mblnWindows8_Pro
End Property

Public Property Get bWindows8orNewer() As Boolean
    bWindows8orNewer = mblnWindows8orNewer
End Property

Public Property Get bWindows81() As Boolean
    bWindows81 = mblnWindows81
End Property

Public Property Get bWindows81_Pro() As Boolean
    bWindows81_Pro = mblnWindows81_Pro
End Property

Public Property Get bWindows81orNewer() As Boolean
    bWindows81orNewer = mblnWindows81orNewer
End Property

Public Property Get bWindows10() As Boolean
    bWindows10 = mblnWindows10
End Property

Public Property Get bWindows10_Beta() As Boolean
    bWindows10_Beta = mblnWindows10_Beta
End Property

Public Property Get bWindows10_Pro_Beta() As Boolean
    bWindows10_Pro_Beta = mblnWindows10_Pro_Beta
End Property

Public Property Get bWindows10_Pro() As Boolean
    bWindows10_Pro = mblnWindows10_Pro
End Property

Public Property Get bWindows10orNewer() As Boolean
    bWindows10orNewer = mblnWindows10orNewer
End Property

' Workstations
Public Property Get bBusinessEd() As Boolean
    bBusinessEd = mblnBusinessEd
End Property

Public Property Get bEmbedded() As Boolean
    bEmbedded = mblnXPEmbedded
End Property

Public Property Get bEnterprise() As Boolean
    bEnterprise = mblnEnterprise
End Property

Public Property Get bHomeBasic() As Boolean
    bHomeBasic = mblnHomeBasic
End Property

Public Property Get bHomePremium() As Boolean
    bHomePremium = mblnHomePremium
End Property

Public Property Get bProfessional() As Boolean
    bProfessional = mblnProfessional
End Property

Public Property Get bProfessional_WMC() As Boolean
    bProfessional_WMC = mblnProfessional_WMC
End Property

Public Property Get bMediaCenter() As Boolean
    bMediaCenter = mblnMediaCenter
End Property

Public Property Get bStarterEd() As Boolean
    bStarterEd = mblnStarterEd
End Property

Public Property Get bUltimateEd() As Boolean
    bUltimateEd = mblnUltimateEd
End Property

Public Property Get bWindowsNT() As Boolean
    bWindowsNT = mblnWindowsNT
End Property

Public Property Get bWorkstation() As Boolean
    bWorkstation = mblnWorkstation
End Property

Public Property Get bSingleUser() As Boolean
    bSingleUser = mblnSingleUser
End Property

Public Property Get bXPHomeEd() As Boolean
    bXPHomeEd = mblnXPHomeEd
End Property

Public Property Get bWindows8SL() As Boolean
    bWindows8SL = mblnWindows8SL
End Property

Public Property Get bUltimateEd_E() As Boolean
    bUltimateEd_E = mblnUltimateEd_E
End Property

Public Property Get bHomeBasic_E() As Boolean
    bHomeBasic_E = mblnHomeBasic_E
End Property

Public Property Get bHomePremium_E() As Boolean
    bHomePremium_E = mblnHomePremium_E
End Property

Public Property Get bEnterprise_E() As Boolean
    bEnterprise_E = mblnEnterprise_E
End Property

Public Property Get bStarterEd_E() As Boolean
    bStarterEd_E = mblnStarterEd_E
End Property

Public Property Get bProfessional_E() As Boolean
    bProfessional_E = mblnProfessional_E
End Property

Public Property Get bUltimateEd_N() As Boolean
    bUltimateEd_N = mblnUltimateEd_N
End Property

Public Property Get bHomeBasic_N() As Boolean
    bHomeBasic_N = mblnHomeBasic_N
End Property

Public Property Get bHomePremium_N() As Boolean
    bHomePremium_N = mblnHomePremium_N
End Property

Public Property Get bEnterprise_N() As Boolean
    bEnterprise_N = mblnEnterprise_N
End Property

Public Property Get bBusinessEd_N() As Boolean
    bBusinessEd_N = mblnBusinessEd_N
End Property

Public Property Get bStarterEd_N() As Boolean
    bStarterEd_N = mblnStarterEd_N
End Property

Public Property Get bProfessional_N() As Boolean
    bProfessional_N = mblnProfessional_N
End Property

Public Property Get bWindows8_N() As Boolean
    bWindows8_N = mblnWindows8_N
End Property

' Servers
Public Property Get bBackOfficeTermSvr() As Boolean
    bBackOfficeTermSvr = mblnBackOfficeTermSvr
End Property

Public Property Get bBackOfficeServer() As Boolean
    bBackOfficeServer = mblnBackOfficeServer
End Property

Public Property Get bCommServer() As Boolean
    bCommServer = mblnCommServer
End Property

Public Property Get bClusterServer() As Boolean
    bClusterServer = mblnClusterServer
End Property

Public Property Get bDatacenterServer() As Boolean
    bDatacenterServer = mblnDatacenterServer
End Property

Public Property Get bDatacenterServerCore() As Boolean
    bDatacenterServerCore = mblnDatacenterServerCore
End Property

Public Property Get bDomainController() As Boolean
    bDomainController = mblnDomainController
End Property

Public Property Get bEnterpriseServer() As Boolean
    bEnterpriseServer = mblnEnterpriseServer
End Property

Public Property Get bEnterpriseServerCore() As Boolean
    bEnterpriseServerCore = mblnEnterpriseServerCore
End Property

Public Property Get bEssBusServerSolAddl() As Boolean
    bEssBusServerSolAddl = mblnEssBusServerSolAddl
End Property

Public Property Get bEssBusServerSolAddlSvc() As Boolean
    bEssBusServerSolAddlSvc = mblnEssBusServerSolAddlSvc
End Property

Public Property Get bEssBusServerSolMgmt() As Boolean
    bEssBusServerSolMgmt = mblnEssBusServerSolMgmt
End Property

Public Property Get bEssBusServerSolMgmtSvc() As Boolean
    bEssBusServerSolMgmtSvc = mblnEssBusServerSolMgmtSvc
End Property

Public Property Get bHomePremiumServer() As Boolean
    bHomePremiumServer = mblnHomePremiumServer
End Property

Public Property Get bHomeEdServer() As Boolean
    bHomeEdServer = mblnHomeEdServer
End Property

Public Property Get bMedBusServerMgmt() As Boolean
    bMedBusServerMgmt = mblnMedBusServerMgmt
End Property

Public Property Get bMedBusServerMessaging() As Boolean
    bMedBusServerMessaging = mblnMedBusServerMessaging
End Property

Public Property Get bMedBusServerSecurity() As Boolean
    bMedBusServerSecurity = mblnMedBusServerSecurity
End Property

Public Property Get bMultiPointServer() As Boolean
    bMultiPointServer = mblnMultiPointServer
End Property

Public Property Get bSB_SolutionServer2011() As Boolean
    bSB_SolutionServer2011 = mblnSB_SolutionServer2011
End Property

Public Property Get bServerForSB_Solution() As Boolean
    bServerForSB_Solution = mblnServerForSB_Solution
End Property

Public Property Get bServerForSB_Solution_EM() As Boolean
    bServerForSB_Solution_EM = mblnServerForSB_Solution_EM
End Property

Public Property Get bServerFoundation() As Boolean
    bServerFoundation = mblnServerFoundation
End Property

Public Property Get bSmallBusRestrictedServer() As Boolean
    bSmallBusRestrictedServer = mblnSmallBusRestrictedServer
End Property

Public Property Get bSmallBusServer() As Boolean
    bSmallBusServer = mblnSmallBusServer
End Property

Public Property Get bSmallBusServerPremCore() As Boolean
    bSmallBusServerPremCore = mblnSmallBusServerPremCore
End Property

Public Property Get bSmallBusServerPremium() As Boolean
    bSmallBusServerPremium = mblnSmallBusServerPremium
End Property

Public Property Get bStandardServer() As Boolean
    bStandardServer = mblnStandardServer
End Property

Public Property Get bStandardServerCore() As Boolean
    bStandardServerCore = mblnStandardServerCore
End Property

Public Property Get bStandardServerSolution() As Boolean
    bStandardServerSolution = mblnStandardServerSolution
End Property

Public Property Get bStandardSvrSolutionCore() As Boolean
    bStandardSvrSolutionCore = mblnStandardSvrSolutionCore
End Property

Public Property Get bSolutionSvrForSB_Sol_EM() As Boolean
    bSolutionSvrForSB_Sol_EM = mblnSolutionSvrForSB_Sol_EM
End Property

Public Property Get bStorageServer() As Boolean
    bStorageServer = mblnStorageServer
End Property

Public Property Get bTerminalServer() As Boolean
    bTerminalServer = mblnTerminalServer
End Property

Public Property Get bWebServer() As Boolean
    bWebServer = mblnWebServer
End Property

Public Property Get bWebServerCore() As Boolean
    bWebServerCore = mblnWebServerCore
End Property

Public Property Get bWin2000Server() As Boolean
    bWin2000Server = mblnWin2000Server
End Property

Public Property Get bWinXPServer() As Boolean
    bWinXPServer = mblnWinXPServer
End Property

Public Property Get bWinServer2003() As Boolean
    bWinServer2003 = mblnWinServer2003
End Property

Public Property Get bWinServer2003R2() As Boolean
    bWinServer2003R2 = mblnWinServer2003R2
End Property

Public Property Get bWinServer2008() As Boolean
    bWinServer2008 = mblnWinServer2008
End Property

Public Property Get bWinServer2008R2() As Boolean
    bWinServer2008R2 = mblnWinServer2008R2
End Property

Public Property Get bWinServer2012() As Boolean
    bWinServer2012 = mblnWinServer2012
End Property

Public Property Get bWinServer2012R2() As Boolean
    bWinServer2012R2 = mblnWinServer2012R2
End Property

Public Property Get bWinServer2016() As Boolean
    bWinServer2016 = mblnWinServer2016
End Property

Public Property Get bWinServer10_Beta() As Boolean
    bWinServer10_Beta = mblnWinServer10_Beta
End Property

Public Property Get bWindowsServer() As Boolean
    bWindowsServer = mblnWindowsServer
End Property

' Hyper-V based
Public Property Get bClusterServer_V() As Boolean
    bClusterServer_V = mblnClusterServer_V
End Property

Public Property Get bDatacenterServerCore_V() As Boolean
    bDatacenterServerCore_V = mblnDatacenterServerCore_V
End Property

Public Property Get bDatacenterServer_V() As Boolean
    bDatacenterServer_V = mblnDatacenterServer_V
End Property

Public Property Get bEnterpriseServerCore_V() As Boolean
    bEnterpriseServerCore_V = mblnEnterpriseServerCore_V
End Property

Public Property Get bEnterpriseServer_V() As Boolean
    bEnterpriseServer_V = mblnEnterpriseServer_V
End Property

Public Property Get bServerForSmallBus_V() As Boolean
    bServerForSmallBus_V = mblnServerForSmallBus_V
End Property

Public Property Get bStandardServerCore_V() As Boolean
    bStandardServerCore_V = mblnStandardServerCore_V
End Property

Public Property Get bStandardServer_V() As Boolean
    bStandardServer_V = mblnStandardServer_V
End Property

' Storage servers
Public Property Get bStorageEnterpriseServer() As Boolean
    bStorageEnterpriseServer = mblnStorageEnterpriseServer
End Property

Public Property Get bStorageEnterpriseServerCore() As Boolean
    bStorageEnterpriseServerCore = mblnStorageEnterpriseServerCore
End Property

Public Property Get bStorageExpressServer() As Boolean
    bStorageExpressServer = mblnStorageExpressServer
End Property

Public Property Get bStorageExpressServerCore() As Boolean
    bStorageExpressServerCore = mblnStorageExpressServerCore
End Property

Public Property Get bStorageServer2008R2() As Boolean
    bStorageServer2008R2 = mblnStorageServer2008R2
End Property

Public Property Get bStorageStandardServer() As Boolean
    bStorageStandardServer = mblnStorageStandardServer
End Property

Public Property Get bStorageStandardServerCore() As Boolean
    bStorageStandardServerCore = mblnStorageStandardServerCore
End Property

Public Property Get bStorageWorkgrpServer() As Boolean
    bStorageWorkgrpServer = mblnStorageWorkgrpServer
End Property

Public Property Get bStorageWorkgrpServerCore() As Boolean
    bStorageWorkgrpServerCore = mblnStorageWorkgrpServerCore
End Property

' Evaluation servers
Public Property Get bDatacenterServerEval() As Boolean
    bDatacenterServerEval = mblnDatacenterServerEval
End Property

Public Property Get bEnterpriseServerEval() As Boolean
    bEnterpriseServerEval = mblnEnterpriseServerEval
End Property

Public Property Get bEnterpriseServerN_Eval() As Boolean
    bEnterpriseServerN_Eval = mblnEnterpriseServerN_Eval
End Property

Public Property Get bMultiPointPremiumServerEval() As Boolean
    bMultiPointPremiumServerEval = mblnMultiPointPremiumServerEval
End Property

Public Property Get bMultiPointStandardServerEval() As Boolean
    bMultiPointStandardServerEval = mblnMultiPointStandardServer
End Property

Public Property Get bStandardServerEval() As Boolean
    bStandardServerEval = mblnStandardServerEval
End Property

Public Property Get bStorageStandardServerEval() As Boolean
    bStorageStandardServerEval = mblnStorageStandardServerEval
End Property

Public Property Get bStorageWorkgroupServerEval() As Boolean
    bStorageWorkgroupServerEval = mblnStorageWorkgroupServerEval
End Property

' 64-Bit Versions
Public Property Get bClusterServer64() As Boolean
    bClusterServer64 = mblnClusterServer64
End Property

Public Property Get bDatacenterServer64() As Boolean
    bDatacenterServer64 = mblnDatacenterServer64
End Property

Public Property Get bEnterpriseItanium64() As Boolean
    bEnterpriseItanium64 = mblnEnterpriseItanium64
End Property

Public Property Get bEnterpriseServer64() As Boolean
    bEnterpriseServer64 = mblnEnterpriseServer64
End Property

Public Property Get bOperSystem64() As Boolean
    bOperSystem64 = mblnOperSystem64
End Property

Public Property Get bStandardServer64() As Boolean
    bStandardServer64 = mblnStandardServer64
End Property

Public Property Get bWinVista64() As Boolean
    bWinVista64 = mblnWinVista64
End Property

Public Property Get bWinXP_Pro64() As Boolean
    bWinXP_Pro64 = mblnWinXP_Pro64
End Property

' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
'  Routine:      GetWindowsVersion
'
'  Description:  Get the current windows operating system version.  This
'                procedure is called by Class_Initialize() when this class
'                is instantiated.  The subsequent called routines may
'                perform additional testing to further refine the type of
'                current operating system.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 04-May-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
'              Added testing for additional Windows operating systems.
' ***************************************************************************
Private Sub GetWindowsVersion()

    ' Called by Class_Initialize()
    
    ResetVariables   ' Initialize module variables
    
    ' Size data structure before filling
    mtypOSVerInfoW.OSVSize = Len(mtypOSVerInfoW)

    ' Make call to load basic information into data structure
    If GetVersionEx(mtypOSVerInfoW) Then

        ' Test for specific features
        With mtypOSVerInfoW

            ' Is this NT based Windows operating system?
            If (.PlatformID = VER_PLATFORM_WIN32_NT) And _
               (.dwVerMajor > 4) Then

                mblnWindowsNT = True   ' Windows NT based operating system
                GetBasicData           ' Capture basic OS data
        
                ' Evaluate major version
                Select Case mlngMajorVer
                       Case 5:  TestMajorVer_5    ' Windows 2000, XP, XP64
                       Case 6:  TestMajorVer_6    ' Windows Vista, 7, 8, 8.1, 10 (Beta)
                       Case 10: TestMajorVer_10   ' Windows 10
                       Case Else: Exit Sub        ' Windows NT4 or earlier (not used by this module)
                End Select

            End If
        End With
    End If

End Sub

Private Sub GetBasicData()

    ' Called by GetWindowsVersion()
    
    Dim lngStart      As Long
    Dim lngPointer    As Long
    Dim strKeyPath    As String
    Dim strValueName  As String
    Dim objReg        As Object
    Dim objItem       As Object
    Dim objOperSys    As Object
    Dim objWMIService As Object
        
    Const HKEY_LOCAL_MACHINE = &H80000002
            
    On Error Resume Next
            
    ' Use default authentication and impersonation settings
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!//./root/cimv2")
    Set objOperSys = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    
    For Each objItem In objOperSys
        
        With objItem
            If Len(.Name) > 0 Then
                lngPointer = InStr(1, .Name, "|", vbBinaryCompare)
                
                ' ex: Microsoft Windows 8.1
                If lngPointer = 0 Then
                    mstrVerName = .Name
                Else
                    mstrVerName = Mid$(.Name, 1, lngPointer - 1)
                End If
                
                mstrVerNumber = .Version                      ' ex: 6.3.9600 | 6.3.99999.9600 | 6.3.9600.99999
                mlngBuildNumber = .BuildNumber                ' ex: 9600
                mintSvrPackMajor = .ServicePackMajorVersion   ' ex: 0
                mintSvrPackMinor = .ServicePackMinorVersion   ' ex: 0
                
                ' Test for workstation or server
                Select Case .ProductType
                       
                       Case VER_NT_WORKSTATION
                            mtypOSVerInfoW.wProductType = VER_NT_WORKSTATION         ' Update type structure
                            mblnWorkstation = True                                   ' Set flag to TRUE
                            mblnWindowsServer = False                                ' Verify flag = FALSE
        
                       Case VER_NT_DOMAIN_CONTROLLER
                            mtypOSVerInfoW.wProductType = VER_NT_DOMAIN_CONTROLLER   ' Update type structure
                            mblnWindowsServer = True                                 ' Set flag to TRUE
                            mblnDomainController = True
                       
                       Case VER_NT_SERVER
                            mtypOSVerInfoW.wProductType = VER_NT_SERVER              ' Update type structure
                            mblnWindowsServer = True                                 ' Set flag to TRUE
                End Select
                
                Exit For   ' exit FOR..NEXT loop
                
            End If
        End With
    
    Next objItem
    
    ' Test for beta version of Windows 10
    If StrComp("6.4", left$(mstrVerNumber, 3), vbBinaryCompare) = 0 Then
        
        ' Professional version
        If InStr(1, UCase$(mstrVerName), "PROF", vbBinaryCompare) > 0 Then
            mstrVerName = mstrVerName & " (Beta)"
        Else
            ' Home version, etc.
            mstrVerName = "Microsoft Windows 10 (Beta)"
        End If
    End If
        
    ' Parse version data
    ' ex: 6.3 | 6.3.9600 | 6.3.99999.9600 | 6.3.9600.99999
    '
    ' Capture major value
    lngStart = 1
    lngPointer = InStr(lngStart, mstrVerNumber, ".", vbBinaryCompare)
    mlngMajorVer = CLng(Mid$(mstrVerNumber, lngStart, lngPointer - lngStart))

    ' Capture minor value
    lngStart = lngPointer + 1
    lngPointer = InStr(lngStart, mstrVerNumber, ".", vbBinaryCompare)
    
    If lngPointer > 0 Then
        mlngMinorVer = CLng(Mid$(mstrVerNumber, lngStart, lngPointer - lngStart))
    Else
        mlngMinorVer = CLng(Mid$(mstrVerNumber, lngStart))
    End If

    mstrVerNumber = vbNullString

    ' Reformat complete version number (ex: 6.3.9600)
    mstrVerNumber = CStr(mlngMajorVer) & "." & _
                    CStr(mlngMinorVer) & "." & _
                    CStr(mlngBuildNumber)
    
    ' Update type structure
    With mtypOSVerInfoW
        .dwVerMajor = mlngMajorVer              ' ex: 6
        .dwVerMinor = mlngMinorVer              ' ex: 3
        .dwBuildNumber = mlngBuildNumber        ' ex: 9600
        .wServicePackMajor = mintSvrPackMajor   ' ex: 0
        .wServicePackMinor = mintSvrPackMinor   ' ex: 0
    End With
        
    ' Set boolean flags
    Select Case mlngMajorVer
           Case 6:  mblnMajorVersion6 = True    ' Vista, 7, 8, 8.1, 10 (Beta)
           Case 10: mblnMajorVersion10 = True   ' 10 (Official Release)
    End Select
    
    ' Capture architecture from environment variable
    ' ex:  IA64 | AMD64 | x86
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!//./root/default:StdRegProv")
    strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    strValueName = "PROCESSOR_ARCHITECTURE"
    objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, mstrProcArch
    Set objReg = Nothing
    
    ' Determine 32-bit | 64-bit
    If InStr(1, mstrProcArch, "64", vbBinaryCompare) > 0 Then
        mblnOperSystem64 = True
        mstrBitStruct = "64-bit"
    Else
        mstrBitStruct = "32-bit"
    End If
    
    ' See if Professional version
    If InStr(1, UCase$(mstrVerName), "PROF", vbBinaryCompare) > 0 Then
                
        mblnProfessional = True   ' Set flag = TRUE
        
        Select Case mlngMajorVer
               Case 5   ' Win2000, XP
                    Select Case mlngMinorVer
                           Case 0
                                mblnWin2000_Pro = True
                           Case Else
                                mblnWinXP_Pro = True
                                mblnWinXP_Pro64 = mblnOperSystem64
                    End Select
                    
               Case 6   ' Vista, 7, 8, 8.1, 10 (Beta)
                    Select Case mlngMinorVer
                           Case 0
                                mblnWinVista_Pro = True
                                mblnWinVista64 = mblnOperSystem64
                                
                           Case 1: mblnWindows7_Pro = True
                           Case 2: mblnWindows8_Pro = True
                           Case 3: mblnWindows81_Pro = True
                           Case 4: mblnWindows10_Pro_Beta = True
                    End Select
        
               Case 10   ' Windows 10 (Official Release)
                    mblnWindows10_Pro = True
        End Select
    End If
        
    Set objReg = Nothing          ' Free objects from memory
    Set objItem = Nothing
    Set objOperSys = Nothing
    Set objWMIService = Nothing
    On Error GoTo 0               ' Nullify this error trap
    
End Sub

Private Sub TestMajorVer_5()

    ' Called by GetWindowsVersion()
    '
    ' Windows XP support has ended as of April 8, 2014
    ' http://windows.microsoft.com/en-us/windows/end-support-help
    
    With mtypOSVerInfoW
        
        Select Case .dwVerMinor

               Case 0
                    ' Windows 2000 testing
                    mblnWin2000 = True
                    mblnWin2000orNewer = True
                    
                    ' Windows 2000 testing to see which
                    ' service packs are installed
                    Select Case mintSvrPackMajor
                           Case 0   ' No service pack installed
                           Case 1   ' Service pack 1 installed
                                mblnWin2000_SP1orNewer = True
                                mblnServicePack1 = True
                           Case 2   ' Service pack 2 installed
                                mblnWin2000_SP2orNewer = True
                                mblnServicePack1 = True
                                mblnServicePack2 = True
                           Case 3   ' Service pack 3 installed
                                mblnWin2000_SP3orNewer = True
                                mblnServicePack1 = True
                                mblnServicePack2 = True
                                mblnServicePack3 = True
                           Case 4   ' Service pack 4 installed
                                mblnWin2000_SP4orNewer = True
                                mblnServicePack1 = True
                                mblnServicePack2 = True
                                mblnServicePack3 = True
                                mblnServicePack4 = True
                    End Select
                        
                    If mblnWindowsServer Then
                        mblnWin2000Server = True
                    End If

               Case 1
                    mblnWinXP = True   ' Base operating system is Windows XP
                    mblnWinXPorNewer = True
                    
                    ' Windows XP testing to see which
                    ' service packs are installed
                    Select Case mintSvrPackMajor
                           Case 0   ' No service pack installed
                           Case 1   ' Service pack 1 installed
                                mblnWinXP_SP1orNewer = True
                                mblnServicePack1 = True
                           Case 2   ' Service pack 2 installed
                                mblnWinXP_SP2orNewer = True
                                mblnServicePack1 = True
                                mblnServicePack2 = True
                           Case 3   ' Service pack 3 installed
                                mblnWinXP_SP3orNewer = True
                                mblnServicePack1 = True
                                mblnServicePack2 = True
                                mblnServicePack3 = True
                    End Select
                    
                    If mblnWindowsServer Then
                        mblnWinXPServer = True
                    End If

               Case 2
                    mblnWinXP = True   ' Base operating system is Windows XP
                    mblnWinXPorNewer = True
                    
                    ' Windows XP Professional testing to
                    ' see which service packs are installed
                    Select Case mintSvrPackMajor
                           Case 0   ' No service pack installed
                           Case 1   ' Service pack 1 installed
                                mblnWinXP_SP1orNewer = True
                                mblnServicePack1 = True
                           Case 2   ' Service pack 2 installed
                                mblnWinXP_SP2orNewer = True
                                mblnServicePack1 = True
                                mblnServicePack2 = True
                           Case 3   ' Service pack 3 installed
                                mblnWinXP_SP3orNewer = True
                                mblnServicePack1 = True
                                mblnServicePack2 = True
                                mblnServicePack3 = True
                    End Select
                            
                    If mblnWindowsServer Then
                    
                        mblnWinXPServer = True
                        
                        ' Because the version numbers for Windows Server 2003 R2 and
                        ' Windows Server 2003 are identical, you must also check the
                        ' value of the SM_SERVERR2 setting.  If it returns a nonzero
                        ' value, the operating system is Windows Server 2003 R2;
                        ' otherwise, it is Windows Server 2003.
                        If GetSystemMetrics(SM_SERVERR2) <> 0 Then
                            mblnWinServer2003R2 = True
                        Else
                            mblnWinServer2003 = True
                        End If
                    End If
        
               Case Else   ' Unknown version
                    mstrVerName = "<Unknown Windows Version>"
                    Exit Sub
        End Select
    End With

    ' Test for additional enhancements
    RefineNonServerData
    
    ' Only test if this is a server
    If mblnWindowsServer Then
        RefineServerData
    End If
    
    ' Remove unwanted leading-trailing characters
    mstrVerName = TrimStr(mstrVerName) & mstrAppend

End Sub

Private Sub TestMajorVer_6()

    ' Called by GetWindowsVersion()
            
    ' Test for earliest version
    With mtypOSVerInfoW
        Select Case mlngMajorVer
        
               Case 6   ' Test for Vista, 7, 8, 8.1, 10 (Beta)
                    Select Case mlngMinorVer
                           
                           Case 0   ' Test Vista
                                mblnWinVista = True
                                mblnWinVistaOrNewer = True
                                
                                ' Windows Vista testing to see which
                                ' service packs are installed
                                Select Case mintSvrPackMajor
                                       Case 0   ' No service pack installed
                                       Case 1   ' Service pack 1 installed
                                            mblnWinVista_SP1orNewer = True
                                            mblnServicePack1 = True
                                       Case 2   ' Service pack 2 installed
                                            mblnWinVista_SP2orNewer = True
                                            mblnServicePack1 = True
                                            mblnServicePack2 = True
                                End Select
                                        
                                If mblnWindowsServer Then
                                    mblnWinServer2008 = True
                                End If
                    
                           Case 1   ' Test Windows 7
                                mblnWindows7 = True
                                mblnWindows7orNewer = True
                                
                                ' Windows 7 testing with no service packs installed
                                ' Windows Vista testing to see which
                                ' service packs are installed
                                Select Case mintSvrPackMajor
                                       Case 0   ' No service pack installed
                                       Case 1   ' Service pack 1 installed
                                            mblnWindows7_SP1orNewer = True
                                            mblnServicePack1 = True
                                End Select
                                
                                If mblnWindowsServer Then
                                    mblnWinServer2008R2 = True
                                End If
                                
                           Case 2   ' Test Windows 8
                                mblnWindows8 = True
                                mblnWindows8orNewer = True
                                mblnCaptionIsCentered = True
                                
                                If mblnWindowsServer Then
                                    mblnWinServer2012 = True
                                End If
                                
                           Case 3   ' Test Windows 8.1
                                mblnWindows81 = True
                                mblnWindows81orNewer = True
                                mblnCaptionIsCentered = True
                                
                                If mblnWindowsServer Then
                                    mblnWinServer2012R2 = True
                                End If
                                
                           Case 4   ' Test Windows 10 (Beta)
                                mblnWindows10 = True
                                mblnWindows10_Beta = True
                                mblnWindows10orNewer = True
                                            
                                If mblnWindowsServer Then
                                    mblnWinServer10_Beta = True
                                End If
                    End Select
                    
               Case Else   ' Unknown version
                    mstrVerName = "<Unknown Windows Version>"
                    Exit Sub
        End Select
    End With

    ' Test for additional enhancements
    RefineNonServerData
    
    ' Only test if this is a server
    If mblnWindowsServer Then
        RefineServerData
    End If
    
    ' Remove unwanted leading-trailing characters
    mstrVerName = TrimStr(mstrVerName) & mstrAppend

End Sub

Private Sub TestMajorVer_10()

    ' Called by GetWindowsVersion()
            
    ' Test for earliest version
    With mtypOSVerInfoW
        Select Case mlngMajorVer
        
               Case 10   ' Test Windows 10 (Official Release)
                    mblnWindows10 = True
                    mblnWindows10orNewer = True
                                
                    If mblnWindowsServer Then
                        mblnWinServer2016 = True
                    End If
                    
               Case Else   ' Unknown version
                    mstrVerName = "<Unknown Windows Version>"
                    Exit Sub
        End Select
    End With

    ' Test for additional enhancements
    RefineNonServerData
    
    ' Only test if this is a server
    If mblnWindowsServer Then
        RefineServerData
    End If
    
    ' Remove unwanted leading-trailing characters
    mstrVerName = TrimStr(mstrVerName) & mstrAppend

End Sub

Private Sub RefineNonServerData()
    
    ' Called by TestMajorVer_5()
    '           TestMajorVer_6()
    '           TestMajorVer_10()
    
    Dim objItem       As Object
    Dim objOperSys    As Object
    Dim objWMIService As Object
    
    On Error GoTo RefineNonServerData_CleanUp
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!//./root/cimv2")
    Set objOperSys = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    
    For Each objItem In objOperSys
 
        Select Case mlngMajorVer
        
                Case 5    ' Win2000, XP
                    Select Case objItem.OSProductSuite
                           Case VER_SUITE_EMBEDDEDNT         ' Windows XP Embedded
                                mblnXPEmbedded = True
                                mstrAppend = ", Embedded"
                           Case VER_SUITE_SINGLEUSERTS       ' Single user
                                mblnSingleUser = True
                                mstrAppend = ", Single User"
                           Case VER_SUITE_PERSONAL           ' Windows XP Home Edition
                                mblnXPHomeEd = True
                                mstrAppend = ", Home Edition"
                    End Select
               
               Case 6, 10   ' Vista, 7, 8, 8.1, 10 (Beta), 10 (Official release)
                    Select Case objItem.OperatingSystemSKU
                           Case PRODUCT_UNDEFINED                             ' 0-Unknown product
                                mblnUndefinedProduct = True
                                mstrAppend = ", <Unknown Operating System>"
                           Case PRODUCT_ULTIMATE                              ' 1-Ultimate Edition
                                mblnUltimateEd = True
                                mstrAppend = ", Ultimate Ed."
                           Case PRODUCT_HOME_BASIC                            ' 2-Home Basic Edition
                                mblnHomeBasic = True
                                mstrAppend = ", Home Basic Ed."
                           Case PRODUCT_HOME_PREMIUM                          ' 3-Home Premium Edition
                                mblnHomePremium = True
                                mstrAppend = ", Home Premium Ed."
                           Case PRODUCT_ENTERPRISE                            ' 4-Enterprise Edition
                                mblnEnterprise = True
                                mstrAppend = ", Enterprise Ed."
                           Case PRODUCT_HOME_BASIC_N                          ' 5-Home Basic N Edition
                                mblnHomeBasic_N = True
                                mstrAppend = ", Home Basic N Ed."
                           Case PRODUCT_BUSINESS                              ' 6-Business Edition
                                mblnBusinessEd = True
                                mstrAppend = ", Business Ed."
                           Case PRODUCT_STARTER                               ' 11-Starter Edition
                                mblnStarterEd = True
                                mstrAppend = ", Starter Ed."
                           Case PRODUCT_BUSINESS_N                            ' 16-Business N Edition
                                mblnBusinessEd_N = True
                                mstrAppend = ", Business N Ed."
                           Case PRODUCT_HOME_PREMIUM_N                        ' 26-Home Premium N Edition
                                mblnHomePremium_N = True
                                mstrAppend = ", Home Premium N Ed."
                           Case PRODUCT_ENTERPRISE_N                          ' 27-Enterprise N Edition
                                mblnEnterprise_N = True
                                mstrAppend = ", Enterprise N Ed."
                           Case PRODUCT_ULTIMATE_N                            ' 28-Ultimate N Edition
                                mblnUltimateEd_N = True
                                mstrAppend = ", Ultimate N Ed."
                           Case PRODUCT_STARTER_N                             ' 47-Starter N Edition
                                mblnStarterEd_N = True
                                mstrAppend = ", Starter N Ed."
                           Case PRODUCT_PROFESSIONAL                          ' 48-Professional Edition
                                mblnProfessional = True
                                mstrAppend = ", Professional Ed."
                           Case PRODUCT_PROFESSIONAL_N                        ' 49-Professional N Edition
                                mblnProfessional_N = True
                                mstrAppend = ", Professional N Ed."
                           Case PRODUCT_STARTER_E                             ' 66-Starter E Edition
                                mblnStarterEd_E = True
                                mstrAppend = ", Starter E Ed."
                           Case PRODUCT_HOME_BASIC_E                          ' 67-Home Basic E Edition
                                mblnHomeBasic_E = True
                                mstrAppend = ", Home Basic E Ed."
                           Case PRODUCT_HOME_PREMIUM_E                        ' 68-Home Premium E Edition
                                mblnHomePremium_E = True
                                mstrAppend = ", Home Premium E Ed."
                           Case PRODUCT_PROFESSIONAL_E                        ' 69-Professional E Edition
                                mblnProfessional_E = True
                                mstrAppend = ", Professional E Ed."
                           Case PRODUCT_ENTERPRISE_E                          ' 70-Enterprise E Edition
                                mblnEnterprise_E = True
                                mstrAppend = ", Enterprise E Ed."
                           Case PRODUCT_ULTIMATE_E                            ' 71-Ultimate E Edition
                                mblnUltimateEd_E = True
                                mstrAppend = ", Ultimate E Ed."
                           Case PRODUCT_WINDOWS_8_N                           ' 98-Windows 8 N Edition
                                mblnWindows8_N = True
                                mstrAppend = ", Windows 8 N Ed."
                           Case PRODUCT_WINDOWS_8_SNG_LANG                    ' 100-Windows 8 Single Language
                                mblnWindows8SL = True
                                mstrAppend = ", Windows 8 Single Language"
                           Case PRODUCT_WINDOWS_EDITION                       ' 101-Home Edition
                                mstrAppend = " Edition"
                           Case PRODUCT_PROFESSIONAL_WMC                      ' 103-Professional with Media Center
                                mblnProfessional_WMC = True
                                mblnMediaCenter = True
                                mstrAppend = ", Professional with Media Center"
                           Case PRODUCT_UNLICENSED                            ' Unlicensed product
                                mblnUnlicensedProduct = True
                                mstrAppend = ", <Unlicensed Product>"
                    End Select
        End Select
        
    Next objItem
    
RefineNonServerData_CleanUp:
    Set objItem = Nothing         ' Free objects from memory
    Set objOperSys = Nothing
    Set objWMIService = Nothing
    On Error GoTo 0               ' Nullify this error trap
        
End Sub

Private Sub RefineServerData()

    ' Called by TestMajorVer_5()
    '           TestMajorVer_6()
    '           TestMajorVer_10()
    
    Dim objItem       As Object
    Dim objOperSys    As Object
    Dim objWMIService As Object
    
    On Error GoTo RefineServerData_CleanUp
    
    ' Support for Windows Server 2003 ended on July 14, 2015
    ' https://www.microsoft.com/en-us/server-cloud/products/windows-server-2003/
    Select Case mtypOSVerInfoW.wSuiteMask
           
           Case VER_SUITE_SMALLBUSINESS
                mblnWinServer2003 = True
                mblnSmallBusServer = True
                mstrAppend = ", Server 2003, Small Business Edition"
           Case VER_SUITE_ENTERPRISE
                If mblnWin2000Server Then
                    mstrAppend = ",  Advanced Server"
                ElseIf mblnWinXPServer Then
                    mblnWinServer2003 = True
                    mstrAppend = ", Server 2003, Enterprise Edition"
                Else
                    mblnWinServer2008 = True
                    mstrVerName = "Windows Server 2008, Enterprise Edition"
                End If
                mblnEnterpriseServer = True
                mblnEnterpriseServer64 = mblnOperSystem64
           Case VER_SUITE_COMMUNICATION
                mblnWinServer2003 = True
                mblnCommServer = True
                mstrAppend = ", Server 2003, Communication Edition"
           Case VER_SUITE_BACKOFFICE
                mblnWinServer2003 = True
                mblnBackOfficeServer = True
                mstrAppend = ", Server 2003, BackOffice Edition"
           Case VER_SUITE_SMALLBUSINESS_RESTRICTED
                mblnWinServer2003 = True
                mblnSmallBusRestrictedServer = True
                mstrAppend = ", Server 2003, Small Business Edition Restricted"
           Case VER_SUITE_EMBEDDEDNT
                mblnXPEmbedded = True
                mstrAppend = ", Embedded"
           Case VER_SUITE_TERMINAL
                mblnWinServer2003 = True
                mblnTerminalServer = True
                mblnBackOfficeTermSvr = True
                mstrAppend = ", Server 2003, Terminal Services"
           Case VER_SUITE_DATACENTER
                If mblnWin2000Server Then
                   mstrAppend = ", DataCenter Server"
                ElseIf mblnWinXPServer Then
                    mblnWinServer2003 = True
                    mstrAppend = ", Server 2003, DataCenter Edition"
                Else
                    mblnWinServer2008 = True
                    mstrAppend = ", Server 2008, DataCenter Edition"
                End If
                mblnDatacenterServer = True
                mblnDatacenterServer64 = mblnOperSystem64
           Case VER_SUITE_SINGLEUSERTS
                mblnSingleUser = True
                mstrAppend = ", Remote Desktop is supported"
           Case VER_SUITE_PERSONAL
                mblnXPHomeEd = True
                mstrAppend = ", Home Edition"
           Case VER_SUITE_BLADE
                mblnWinServer2003 = True
                mblnWebServer = True
                mstrAppend = ", Server 2003, Web Edition"
           Case VER_SUITE_STORAGE_SERVER
                mblnWinServer2003 = True
                mblnWinServer2003R2 = True
                mblnStorageServer = True
                mstrAppend = ", Server 2003, Storage"
           Case VER_SUITE_COMPUTE_SERVER
                mblnWinServer2003 = True
                mblnClusterServer = True
                mblnClusterServer64 = mblnOperSystem64
                mstrAppend = ", Cluster Edition Server"
           Case VER_SUITE_WH_SERVER
                mblnWinServer2003 = True
                mblnHomeEdServer = True
                mstrAppend = ", Server 2003, Home Edition"
    End Select
        
    ' Other types of servers
    If mlngMajorVer > 5 Then
            
        Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!//./root/cimv2")
        Set objOperSys = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
        
        For Each objItem In objOperSys
 
            Select Case objItem.OperatingSystemSKU
                   
                   Case PRODUCT_STANDARD_SERVER                          ' 7-Server Standard
                        mblnStandardServer = True
                        mblnStandardServer64 = mblnOperSystem64
                        mstrAppend = ", Standard Server Edition"
                   Case PRODUCT_DATACENTER_SERVER                        ' 8-Datacenter Server Edition
                        mblnDatacenterServer = True
                        mblnDatacenterServer64 = mblnOperSystem64
                        mstrAppend = ", Datacenter Server Edition"
                   Case PRODUCT_SMALLBUSINESS_SERVER                     ' 9-Small Business Server Edition
                        mblnSmallBusServer = True
                        mstrAppend = ", Small Business Server Edition"
                   Case PRODUCT_ENTERPRISE_SERVER                        ' 10-Enterprise Server Edition
                        mblnEnterpriseServer = True
                        mblnEnterpriseServer64 = mblnOperSystem64
                        mstrAppend = ", Enterprise Server Edition"
                   Case PRODUCT_DATACENTER_SERVER_CORE                   ' 12-Datacenter Server Core Edition
                        mblnDatacenterServerCore = True
                        mblnDatacenterServer64 = mblnOperSystem64
                        mstrAppend = ", Datacenter Server Core Edition"
                   Case PRODUCT_STANDARD_SERVER_CORE                     ' 13-Standard Server Core Edition
                        mblnStandardServerCore = True
                        mblnStandardServer64 = mblnOperSystem64
                        mstrAppend = ", Standard Server Core Edition"
                   Case PRODUCT_ENTERPRISE_SERVER_CORE                   ' 14-Enterprise Server Core Edition
                        mblnEnterpriseServerCore = True
                        mblnEnterpriseServer64 = mblnOperSystem64
                        mstrAppend = ", Enterprise Server Core Edition"
                   Case PRODUCT_ENTERPRISE_SERVER_ITANIUM                ' 15-Enterprise Server Edition for Itanium-Based Systems
                        mblnEnterpriseItanium64 = True
                        mblnEnterpriseServer64 = mblnOperSystem64
                        mstrAppend = ", Enterprise Server IA64 Edition"
                   Case PRODUCT_WEB_SERVER                               ' 17-Web Server Edition
                        mblnWebServer = True
                        mstrAppend = ", Web Server Edition"
                   Case PRODUCT_CLUSTER_SERVER                           ' 18-Cluster Server Edition
                        mblnClusterServer = True
                        mblnClusterServer64 = mblnOperSystem64
                        mstrAppend = ", Cluster Server Edition"
                   Case PRODUCT_STORAGE_SERVER_2008_R2_ESS               ' 19-Windows Storage Server 2008 R2 Essentials
                        mblnWinServer2008R2 = True
                        mblnStorageServer2008R2 = True
                        mstrAppend = ", Storage Server 2008 R2 Essentials"
                   Case PRODUCT_STORAGE_SERVER_EXPRESS                   ' 20-Storage Express Server Edition
                        mblnStorageExpressServer = True
                        mstrAppend = ", Storage Express Server Edition"
                   Case PRODUCT_STORAGE_SERVER_STANDARD                  ' 21-Storage Standard Server Edition
                        mblnStorageStandardServer = True
                        mstrAppend = ", Storage Standard Server Edition"
                   Case PRODUCT_STORAGE_SERVER_WORKGROUP                 ' 22-Storage Workgroup Server Edition
                        mblnStorageWorkgrpServer = True
                        mstrAppend = ", Storage Workgroup Server Edition"
                   Case PRODUCT_STORAGE_SERVER_ENTERPRISE                ' 23 = Storage Enterprise Server Edition
                        mblnStorageEnterpriseServer = True
                        mstrAppend = ", Storage Enterprise Server Edition"
                   Case PRODUCT_SERVER_2008_ESS_SVCS                     ' 24-Windows Server 2008 for Windows Essential Server Solutions
                        mblnWinServer2008 = True
                        mstrAppend = ", Server 2008, Essential Solutions"
                   Case PRODUCT_SMALL_BUS_SERVER_PREMIUM                 ' 25-Small Business Server Premium Edition
                        mblnSmallBusServerPremium = True
                        mstrAppend = ", Small Business Server Premium Edition"
                   Case PRODUCT_WEB_SERVER_CORE                          ' 29-Web Server (core installation)
                        mblnWebServerCore = True
                        mstrAppend = ", Web Server Core"
                   Case PRODUCT_ESS_BUS_SERVER_MANAGEMENT                ' 30-Windows Essential Business Management Server
                        mblnMedBusServerMgmt = True
                        mstrAppend = ", Windows Essential Business Management Server"
                   Case PRODUCT_ESS_BUS_SERVER_SECURITY                  ' 31-Windows Essential Business Security Server
                        mblnMedBusServerSecurity = True
                        mstrAppend = ", Windows Essential Business Security Server"
                   Case PRODUCT_ESS_BUS_SERVER_MESSAGING                 ' 32-Windows Essential Business Messaging Server
                        mblnMedBusServerMessaging = True
                        mstrAppend = ", Windows Essential Business Messaging Server"
                   Case PRODUCT_SERVER_FOUNDATION                        ' 33-Server Foundation
                        mblnServerFoundation = True
                        mstrAppend = ", Server Foundation"
                   Case PRODUCT_HOME_SERVER_2011                         ' 34-Windows Home Server 2011
                        mblnHomePremiumServer = True
                        mstrAppend = ", Home Premium Server"
                   Case PRODUCT_SERVER_2008_WO_V                         ' 35-Windows Server 2008 without Hyper-V for Windows Essential Server Solutions
                        mblnServerForSmallBus_V = True
                        mstrAppend = ", Server 2008 w/o Hyper-V"
                   Case PRODUCT_STANDARD_SERVER_WO_V                     ' 36-Server Standard without Hyper-V
                        mblnStandardServer_V = True
                        mblnStandardServer64 = mblnOperSystem64
                        mstrAppend = ", Standard Server w/o Hyper-V"
                   Case PRODUCT_DATACENTER_SERVER_WO_V                   ' 37-Server Datacenter without Hyper-V (core installation)
                        mblnDatacenterServer_V = True
                        mblnDatacenterServer64 = mblnOperSystem64
                        mstrAppend = ", Datacenter Server w/o Hyper-V"
                   Case PRODUCT_ENTERPRISE_SVR_WO_V                      ' 38-Server Enterprise without Hyper-V (core installation)
                        mblnEnterpriseServer_V = True
                        mblnEnterpriseServer64 = mblnOperSystem64
                        mstrAppend = ", Enterprise Server w/o Hyper-V"
                   Case PRODUCT_DATACENTER_SVR_CORE_WO_V                 ' 39-Server Datacenter without Hyper-V (core installation)
                        mblnDatacenterServerCore_V = True
                        mblnDatacenterServer64 = mblnOperSystem64
                        mstrAppend = ", Datacenter Server Core w/o Hyper-V"
                   Case PRODUCT_STANDARD_SVR_CORE_WO_V                   ' 40-Server Standard without Hyper-V (core installation)
                        mblnStandardServerCore_V = True
                        mblnStandardServer64 = mblnOperSystem64
                        mstrAppend = ", Standard Server Core w/o Hyper-V"
                   Case PRODUCT_ENTERPRISE_SVR_CORE_WO_V                 ' 41-Server Enterprise without Hyper-V (core installation)
                        mblnEnterpriseServerCore_V = True
                        mblnEnterpriseServer64 = mblnOperSystem64
                        mstrAppend = ", Enterprise Server Core w/o Hyper-V"
                   Case PRODUCT_HYPER_V                                  ' 42-Windows Storage Server 2008 R2 Essentials
                        mblnStorageServer2008R2 = True
                        mstrAppend = ", Storage Server 2008 R2 Essentials"
                   Case PRODUCT_STORAGE_SERVER_EXPRESS_CORE              ' 43-Storage Server Express without Hyper-V, Server Core
                        mblnStorageExpressServerCore = True
                        mstrAppend = ", Storage Express Server Core"
                   Case PRODUCT_STORAGE_SERVER_STANDARD_CORE             ' 44-Storage Server Standard without Hyper-V, Server Core
                        mblnStorageStandardServerCore = True
                        mstrAppend = ", Storage Standard Server Core"
                   Case PRODUCT_STORAGE_SERVER_WORKGROUP_CORE            ' 45-Storage Server Workgroup without Hyper-V, Server Core
                        mblnStorageWorkgrpServerCore = True
                        mstrAppend = ", Storage Workgroup Server Core"
                   Case PRODUCT_STORAGE_SERVER_ENTERPRISE_CORE           ' 46-Storage Server Enterprise without Hyper-V, Server Core
                        mblnStorageEnterpriseServerCore = True
                        mstrAppend = ", Storage Enterprise Server Core"
                   Case PRODUCT_SB_SERVER_2011_ESS                       ' 50-Windows Small Business Solution Server 2011 Essentials
                        mblnSB_SolutionServer2011 = True
                        mstrAppend = ", Small Business Solution Server 2011 Essentials"
                   Case PRODUCT_SERVER_FOR_SB_SOLUTIONS                  ' 51-Server For Small Business Solutions
                        mblnServerForSB_Solution = True
                        mstrAppend = ", Server For Small Business Solutions"
                   Case PRODUCT_SERVER_SOLUTIONS_PREM                    ' 52-Server Solutions Premium
                        mblnStandardServerSolution = True
                        mblnStandardServer64 = mblnOperSystem64
                        mstrAppend = ", Standard Server Solutions"
                   Case PRODUCT_SERVER_SOLUTIONS_PREM_CORE               ' 53-Server Solutions Premium (core installation)
                        mblnStandardSvrSolutionCore = True
                        mstrAppend = ", Server Solutions Premium (core installation)"
                   Case PRODUCT_SERVER_SB_SOLUTIONS_EM_1                 ' 54-Server For Small Business Solutions EM
                        mblnServerForSB_Solution_EM = True
                        mstrAppend = ", Server For Small Business Solutions EM"
                   Case PRODUCT_SERVER_SB_SOLUTIONS_EM_2                 ' 55-Solution Server For Small Business Solutions EM
                        mblnSolutionSvrForSB_Sol_EM = True
                        mstrAppend = ", Solution Server For Small Business Solutions EM"
                   Case PRODUCT_MULTIPOINT_SERVER                        ' 56-Windows MultiPoint Server
                        mblnMultiPointServer = True
                        mstrAppend = ", Multi-Point Server "
                   Case PRODUCT_ESS_SERVER_SOLUTION_MGMT                  ' 59-Windows Essential Server Solution Management
                        mblnEssBusServerSolMgmt = True
                        mstrAppend = ", Essential Server Solution Management"
                   Case PRODUCT_ESS_SERVER_SOLUTION_ADDL                  ' 60-Windows Essential Server Solution Additional
                        mblnEssBusServerSolAddl = True
                        mstrAppend = ", Essential Server Solution Additional"
                   Case PRODUCT_ESS_SERVER_SOLUTION_MGMT_SVC              ' 61-Windows Essential Server Solution Management Services
                        mblnEssBusServerSolMgmtSvc = True
                        mstrAppend = ", Essential Server Solution Management Services"
                   Case PRODUCT_ESS_SERVER_SOLUTION_ADDL_SVC              ' 62-Windows Essential Server Solution Additional Services
                        mblnEssBusServerSolAddlSvc = True
                        mstrAppend = ", Essential Server Solution Additional Services"
                   Case PRODUCT_SMALL_BUS_SERVER_PREMIUM_CORE             ' 63-Small Business Server Premium Core
                        mblnSmallBusServerPremCore = True
                        mstrAppend = ", Small Business Server Premium Core"
                   Case PRODUCT_SERVER_HYPER_V                            ' 64-Cluster Server w/ Hyper V Core
                        mblnClusterServer_V = True
                        mblnClusterServer64 = mblnOperSystem64
                        mstrAppend = ", Cluster Server w/ Hyper V Core"
                   Case PRODUCT_SVR_ENTERPRISE_EVAL                       ' 72-Server Enterprise (evaluation installation)
                        mblnEnterpriseServerEval = True
                        mblnEnterpriseServer64 = mblnOperSystem64
                        mstrAppend = ", Enterprise Server Evaluation"
                   Case PRODUCT_MULTIPOINT_SERVER_STANDARD                ' 76-Windows MultiPoint Server Standard (full installation)
                        mblnMultiPointStandardServer = True
                        mstrAppend = ", MultiPoint Standard Server"
                   Case PRODUCT_MULTIPOINT_SERVER_PREMIUM                 ' 77-Windows MultiPoint Server Premium (full installation)
                        mblnMultiPointPremiumServerEval = True
                        mstrAppend = ", MultiPoint Premium Server Evaluation"
                   Case PRODUCT_SERVER_STANDARD_EVAL                      ' 79-Standard Server (evaluation installation)
                        mblnStandardServerEval = True
                        mblnStandardServer64 = mblnOperSystem64
                        mstrAppend = ", Standard Server Evaluation"
                   Case PRODUCT_SERVER_DATACENTER_EVAL                    ' 80-Datacenter Server (evaluation installation)
                        mblnDatacenterServerEval = True
                        mblnDatacenterServer64 = mblnOperSystem64
                        mstrAppend = ", Datacenter Server Evaluation"
                   Case PRODUCT_ENTERPRISE_N_EVAL                         ' 84-Enterprise N Server (evaluation installation)
                        mblnEnterpriseServerN_Eval = True
                        mblnEnterpriseServer64 = mblnOperSystem64
                        mstrAppend = ", Enterprise N Server Evaluation"
                   Case PRODUCT_STORAGE_SERVER_WORKGROUP_EVAL             ' 95-Storage Server Workgroup (evaluation installation)
                        mblnStorageWorkgroupServerEval = True
                        mstrAppend = ", Storage Workgroup Server Evaluation"
                   Case PRODUCT_STORAGE_SERVER_STANDARD_EVAL              ' 96-Storage Server Standard (evaluation installation)
                        mblnStorageStandardServerEval = True
                        mstrAppend = ", Storage Standard Server Evaluation"
                        
            End Select   ' objItem.OperatingSystemSKU
        Next objItem
        
    End If   ' mlngMajorVer
        
RefineServerData_CleanUp:
    Set objItem = Nothing         ' Free objects from memory
    Set objOperSys = Nothing
    Set objWMIService = Nothing
    On Error GoTo 0               ' Nullify this error trap
    
End Sub

Private Sub ResetVariables()

    ' Called by GetWindowsVersion()
    
    ' Generalized flags
    mblnWindowsNT = False
    mblnCaptionIsCentered = False
    mblnMajorVersion6 = False
    mblnMajorVersion10 = False
    mblnUndefinedProduct = False
    mblnUnlicensedProduct = False
    mblnWin2000orNewer = False
    mblnWin2000_SP1orNewer = False
    mblnWin2000_SP2orNewer = False
    mblnWin2000_SP3orNewer = False
    mblnWin2000_SP4orNewer = False
    mblnWinVistaOrNewer = False
    mblnWinVista_SP1orNewer = False
    mblnWinVista_SP2orNewer = False
    mblnWinXPorNewer = False
    mblnWinXP_SP1orNewer = False
    mblnWinXP_SP2orNewer = False
    mblnWinXP_SP3orNewer = False
    mblnWindows7_SP1orNewer = False
    mblnWindows7orNewer = False
    mblnWindows8orNewer = False
    mblnWindows81orNewer = False
    mblnWindows10orNewer = False

    ' Service packs
    mblnServicePack1 = False
    mblnServicePack2 = False
    mblnServicePack3 = False
    mblnServicePack4 = False

    ' Base operating systems
    mblnWin2000 = False
    mblnWin2000Server = False
    mblnWinVista = False
    mblnWinXP = False
    mblnWinXP_Pro = False
    mblnWinXPServer = False
    mblnWindows7 = False
    mblnWindows8 = False
    mblnWindows81 = False
    mblnWindows10 = False
    mblnWindows10_Beta = False
    
    mblnWin2000_Pro = False
    mblnWinVista_Pro = False
    mblnWindows7_Pro = False
    mblnWindows8_Pro = False
    mblnWindows81_Pro = False
    mblnWindows10_Pro = False
    mblnWindows10_Pro_Beta = False

    ' Workstations
    mblnBusinessEd = False
    mblnXPEmbedded = False
    mblnEnterprise = False
    mblnHomeBasic = False
    mblnHomePremium = False
    mblnMediaCenter = False
    mblnProfessional = False
    mblnStarterEd = False
    mblnUltimateEd = False
    mblnWorkstation = False
    mblnSingleUser = False
    mblnXPHomeEd = False
    mblnWindows8SL = False
    
    mblnUltimateEd_E = False
    mblnHomeBasic_E = False
    mblnHomePremium_E = False
    mblnEnterprise_E = False
    mblnStarterEd_E = False
    mblnProfessional_E = False
  
    mblnUltimateEd_N = False
    mblnHomeBasic_N = False
    mblnHomePremium_N = False
    mblnEnterprise_N = False
    mblnBusinessEd_N = False
    mblnStarterEd_N = False
    mblnProfessional_N = False
    mblnWindows8_N = False
        
    ' Servers
    mblnBackOfficeServer = False
    mblnBackOfficeTermSvr = False
    mblnClusterServer = False
    mblnCommServer = False
    mblnDatacenterServer = False
    mblnDatacenterServerCore = False
    mblnDomainController = False
    mblnEnterpriseServer = False
    mblnEnterpriseServerCore = False
    mblnEssBusServerSolAddl = False
    mblnEssBusServerSolAddlSvc = False
    mblnEssBusServerSolMgmt = False
    mblnEssBusServerSolMgmtSvc = False
    mblnHomePremiumServer = False
    mblnHomeEdServer = False
    mblnMedBusServerMgmt = False
    mblnMedBusServerMessaging = False
    mblnMedBusServerSecurity = False
    mblnMultiPointServer = False
    mblnProfessional_WMC = False
    mblnSB_SolutionServer2011 = False
    mblnServerForSB_Solution = False
    mblnServerForSB_Solution_EM = False
    mblnServerFoundation = False
    mblnSmallBusRestrictedServer = False
    mblnSmallBusServer = False
    mblnSmallBusServerPremCore = False
    mblnSmallBusServerPremium = False
    mblnStandardServer = False
    mblnStandardServerCore = False
    mblnStandardServerSolution = False
    mblnStandardSvrSolutionCore = False
    mblnSolutionSvrForSB_Sol_EM = False
    mblnStorageServer = False
    mblnTerminalServer = False
    mblnWebServer = False
    mblnWebServerCore = False
    mblnWinServer10_Beta = False
    mblnWinServer2003 = False
    mblnWinServer2003R2 = False
    mblnWinServer2008 = False
    mblnWinServer2008R2 = False
    mblnWinServer2012 = False
    mblnWinServer2012R2 = False
    mblnWinServer2016 = False
    mblnWindowsServer = False

    ' Hyper-V based
    mblnClusterServer_V = False
    mblnDatacenterServerCore_V = False
    mblnDatacenterServer_V = False
    mblnEnterpriseServerCore_V = False
    mblnEnterpriseServer_V = False
    mblnServerForSmallBus_V = False
    mblnStandardServerCore_V = False
    mblnStandardServer_V = False
    
    ' Storage servers
    mblnStorageEnterpriseServer = False
    mblnStorageEnterpriseServerCore = False
    mblnStorageExpressServer = False
    mblnStorageExpressServerCore = False
    mblnStorageServer2008R2 = False
    mblnStorageStandardServer = False
    mblnStorageStandardServerCore = False
    mblnStorageWorkgrpServer = False
    mblnStorageWorkgrpServerCore = False
        
    ' Evaluation servers
    mblnDatacenterServerEval = False
    mblnEnterpriseServerEval = False
    mblnEnterpriseServerN_Eval = False
    mblnMultiPointPremiumServerEval = False
    mblnMultiPointStandardServer = False
    mblnStandardServerEval = False
    mblnStorageStandardServerEval = False
    mblnStorageWorkgroupServerEval = False
 
    ' 64-Bit Versions
    mblnClusterServer64 = False
    mblnDatacenterServer64 = False
    mblnEnterpriseItanium64 = False
    mblnEnterpriseServer64 = False
    mblnOperSystem64 = False
    mblnStandardServer64 = False
    mblnWinVista64 = False
    mblnWinXP_Pro64 = False

    EmptyOtherVariables
    
End Sub

Private Sub EmptyOtherVariables()

    ' Called by ResetVariables()
    '           Class_Terminate()
    
    ' Empty string variables
    mstrAppend = vbNullString
    mstrVerName = vbNullString
    mstrProcArch = vbNullString
    mstrBitStruct = vbNullString
    mstrVerNumber = vbNullString

    ' Empty type structure
    ZeroMemory mtypOSVerInfoW, Len(mtypOSVerInfoW)

End Sub

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
    
    GetWindowsVersion   ' Determine current operating system

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.

    EmptyOtherVariables

End Sub


