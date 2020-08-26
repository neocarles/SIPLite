Attribute VB_Name = "ModuleKoneksiMySQL"
Public RSData As ADODB.Recordset
Public strSQL As String

Public Enum MyOption
    'recommended option values for various configurations.
    optVB = 3                      'Microsoft Access, Visual Basic
    optLargeTables = 2049          'Large tables with too many rows
    optSybasePB = 135168           'sysbase powerbuilder
    optLT_nocache = 3145731         'Large tables with no-cache results

    'other Option Flags
    optFieldLength = 1              'FLAG_FIELD_LENGTH       'Don't Optimize Column Width
    optFoundRows = 2                'FLAG_FOUND_ROWS         'Return Matching Rows
    optDebug = 4                    'FLAG_DEBUG              'Trace Driver Calls To myoptbc.log
    optBigPacket = 8                'FLAG_BIG_PACKETS        'Allow Big Results
    optNoPrompt = 16                'FLAG_NO_PROMPT          'Don't Prompt Upon Connect
    optDynamicCursor = 32           'FLAG_DYNAMIC_CURSOR     'Enable Dynamic Cursor
    optNoSchema = 64                'FLAG_NO_SCHEMA          'Ignore # in Table Name
    optNoDefaultCursor = 128        'FLAG_NO_DEFAULT_CURSOR  'User Manager Cursors
    optNoLocale = 256               'FLAG_NO_LOCALE          'Don't Use Set Locale
    optPadSpace = 512               'FLAG_PAD_SPACE          'Pad Char To Full Length
    optFullColumnNames = 1024       'FLAG_FULL_COLUMN_NAMES  'Return Table Names for SQLDescribeCol
    optCompressedProto = 2048       'FLAG_COMPRESSED_PROTO   'Use Compressed Protocol
    optIgnoreSpace = 4096           'FLAG_IGNORE_SPACE       'Ignore Space After Function Names
    optNamedPipe = 8192             'FLAG_NAMED_PIPE         'Force Use of Named Pipes
    optNoBigInt = 16384             'FLAG_NO_BIGINT          'Change BIGINT Columns to Int
    optNoCatalog = 32768            'FLAG_NO_CATALOG         'No Catalog
    optUseMyCnf = 65536             'FLAG_USE_MYCNF          'Read Options From my.cnf
    optSafe = 131072                'FLAG_SAFE               'Safe
    optNoTransactions = 262144      'FLAG_NO_TRANSACTIONS    'Disable transactions
    optLogQuery = 524288            'FLAG_LOG_QUERY          'Save queries to myodbc.sql
    optNoCache = 1048576            'FLAG_NO_CACHE           'Don't Cache Result (forward only cursors)
    optForwardCursor = 2097152      'FLAG_FORWARD_CURSOR     'Force Use Of Forward Only Cursors
    optAutoReconnect = 4194304      'FLAG_AUTO_RECONNECT     'Enable auto-reconnect.
    optAutoIsNull = 8388608         'FLAG_AUTO_IS_NULL       'Flag Auto Is Null
    optZeroDateToMin = 16777216     'FLAG_ZERO_DATE_TO_MIN   'Flag Zero Date to Min
    optMinDateToZero = 33554432     'FLAG_MIN_DATE_TO_ZERO   'Flag Min Date to Zero
    optMultiStatements = 67108864   'FLAG_MULTI_STATEMENTS   'Allow multiple statements
    optColumnSizeS32 = 134217728    'FLAG_COLUMN_SIZE_S32    'Limit column size to 32-bit value
End Enum

Public Conn As New ADODB.Connection
Public DriverConn, DBase, Ports, Servers, Username, Password As String
Public Konfigurasi As String

Public Function BukaMySQL() As Boolean
    On Error GoTo HellHandle
    Dim MyOdbcOption
    Set Conn = New ADODB.Connection
    MyOdbcOption = optFoundRows + optVB + optBigPacket + optDynamicCursor + _
                   optCompressedProto + optNoBigInt + optAutoReconnect

    Konfigurasi = App.Path & "\Config\Koneksi.ini"
    DriverConn = ReadINI("ConnMySQL", "Driver", Konfigurasi)
    DBase = Decrypt(ReadINI("ConnMySQL", "DBase", Konfigurasi))
    Ports = Decrypt(ReadINI("ConnMySQL", "Port", Konfigurasi))
    Servers = Decrypt(ReadINI("ConnMySQL", "Server", Konfigurasi))
    Username = Decrypt(ReadINI("ConnMySQL", "Username", Konfigurasi))
    Password = Decrypt(ReadINI("ConnMySQL", "Password", Konfigurasi))
    Driver = Decrypt(ReadINI("ConnMySQL", "Driver", Konfigurasi))

    With Conn
        .CursorLocation = adUseClient
        .ConnectionString = "DRIVER=" & Driver & ";SERVER=" & Servers & ";DATABASE=" & DBase & ";UID=" & Username & ";PWD=" & Password & ";PORT=" & Ports & ";OPTION=" & MyOdbcOption
        .Open
    End With

    BukaMySQL = True

    Exit Function

HellHandle:
    BukaMySQL = False
'    FormKoneksi.Show
End Function


