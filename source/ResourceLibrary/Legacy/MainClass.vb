'Option Explicit On
'Option Strict On

'Imports Contensive.BaseClasses

'Namespace Contensive.VbConversion
'    Public Class MainClass
'        Private cp As CPBaseClass
'        Public Sub New(cp As CPBaseClass)
'            Me.cp = cp
'        End Sub
'        '
'        ' a way to fake cs integers. an array of csController objects
'        ' 
'        Private csArray(100) As CPCSBaseClass
'        ''
'        ''
'        ''
'        'Type InStreamType
'        '    Name As String
'        '    Value As String
'        '    NameValue As String
'        '    IsForm As Boolean
'        '    IsFile As Boolean
'        '    fileContent() As Byte
'        'End Type

'        ''
'        ''
'        ''
'        'Type InstType
'        '    Type As Long
'        '    Caption As String
'        '    REquired As Boolean
'        '    PeopleField As String
'        '    GroupName As String
'        'End Type

'        ''
'        ''
'        ''
'        'Type FormPageType
'        '    PreRepeat As String
'        '    PostRepeat As String
'        '    RepeatCell As String
'        '    AddGroupNameList As String
'        '    AuthenticateOnFormProcess As Boolean
'        '    Inst() As InstType
'        'End Type

'        ''
'        ''
'        ''
'        'Type InputSelectCacheType
'        '    SelectRaw As String
'        '    ContentName As String
'        '    Criteria As String
'        '    CurrentValue As String
'        'End Type

'        ''
'        ''
'        ''
'        'Enum AddonContextEnum
'        '        ' these should have been addonContextPage, etc.
'        '        ContextPage = 1
'        '        ContextAdmin = 2
'        '        ContextTemplate = 3
'        '        contextEmail = 4
'        '        ContextRemoteMethod = 5
'        '        ContextOnNewVisit = 6
'        '        ContextOnPageEnd = 7
'        '        ContextOnPageStart = 8
'        '        ContextEditor = 9
'        '        ContextHelpUser = 10
'        '        ContextHelpAdmin = 11
'        '        ContextHelpDeveloper = 12
'        '        ContextOnContentChange = 13
'        '        ContextFilter = 14
'        '        ContextSimple = 15
'        '        ContextOnBodyStart = 16
'        '        ContextOnBodyEnd = 17
'        '    End Enum

'        '    '
'        '    '
'        '    '
'        '    Enum EditorUserScopeEnum
'        '        ' should have been userTypeAdministrator, etc
'        '        Administrator = 1
'        '        ContentManager = 2
'        '        PublicUser = 3
'        '    End Enum

'        '    '
'        '    '
'        '    '
'        '    Enum EditorContentScopeEnum
'        '        ' should have been contentTypePage
'        '        Page = 1
'        '        PageTemplate = 2
'        '        email = 3
'        '        EmailTemplate = 4
'        '    End Enum
'        '    '
'        '    '   get cmc object
'        '    '       used only for proxy classes covering older compatility calls
'        '    '       for example, AdminUi has methods that pass in mainClass. cpCom class handles to work, but requires cmc not Main
'        '    '
'        '    Public Function getCmc() As comMainCsvClass
'        '    Set getCmc = cmc
'        'End Function

'        '    '
'        '    ' EditWrapperCnt
'        '    '
'        '    Public Property Let EditWrapperCnt(ByVal vNewValue As Long)
'        '    cmc.main_EditWrapperCnt = vNewValue
'        'End Property
'        '    '
'        '    ' EditWrapperCnt
'        '    '
'        '    Public Property Get EditWrapperCnt() As Long
'        '    EditWrapperCnt = cmc.main_EditWrapperCnt
'        'End Property
'        '    '
'        '    ' BrowserLanguage
'        '    '
'        '    Public Property Let BrowserLanguage(ByVal vNewValue As String)
'        '    cmc.main_BrowserLanguage = vNewValue
'        'End Property
'        '    '
'        '    ' BrowserLanguage
'        '    '
'        '    Public Property Get BrowserLanguage() As String
'        '    BrowserLanguage = cmc.main_BrowserLanguage
'        'End Property
'        '    '
'        '    ' HTTP_Accept
'        '    '
'        '    Public Property Let HTTP_Accept(ByVal vNewValue As String)
'        '    cmc.main_HTTP_Accept = vNewValue
'        'End Property
'        '    '
'        '    ' HTTP_Accept
'        '    '
'        '    Public Property Get HTTP_Accept() As String
'        '    HTTP_Accept = cmc.main_HTTP_Accept
'        'End Property
'        '    '
'        '    ' HTTP_Accept_charset
'        '    '
'        '    Public Property Let HTTP_Accept_charset(ByVal vNewValue As String)
'        '    cmc.main_HTTP_Accept_charset = vNewValue
'        'End Property
'        '    '
'        '    ' HTTP_Accept_charset
'        '    '
'        '    Public Property Get HTTP_Accept_charset() As String
'        '    HTTP_Accept_charset = cmc.main_HTTP_Accept_charset
'        'End Property
'        '    '
'        '    ' HTTP_Profile
'        '    '
'        '    Public Property Let HTTP_Profile(ByVal vNewValue As String)
'        '    cmc.main_HTTP_Profile = vNewValue
'        'End Property
'        '    '
'        '    ' HTTP_Profile
'        '    '
'        '    Public Property Get HTTP_Profile() As String
'        '    HTTP_Profile = cmc.main_HTTP_Profile
'        'End Property
'        '    '
'        '    ' HTTP_X_Wap_Profile
'        '    '
'        '    Public Property Let HTTP_X_Wap_Profile(ByVal vNewValue As String)
'        '    cmc.main_HTTP_X_Wap_Profile = vNewValue
'        'End Property
'        '    '
'        '    ' HTTP_X_Wap_Profile
'        '    '
'        '    Public Property Get HTTP_X_Wap_Profile() As String
'        '    HTTP_X_Wap_Profile = cmc.main_HTTP_X_Wap_Profile
'        'End Property
'        '    '
'        '    ' HTTP_Via
'        '    '
'        '    Public Property Let HTTP_Via(ByVal vNewValue As String)
'        '    cmc.main_HTTP_Via = vNewValue
'        'End Property
'        '    '
'        '    ' HTTP_Via
'        '    '
'        '    Public Property Get HTTP_Via() As String
'        '    HTTP_Via = cmc.main_HTTP_Via
'        'End Property
'        '    '
'        '    ' HTTP_From
'        '    '
'        '    Public Property Let HTTP_From(ByVal vNewValue As String)
'        '    cmc.main_HTTP_From = vNewValue
'        'End Property
'        '    '
'        '    ' HTTP_From
'        '    '
'        '    Public Property Get HTTP_From() As String
'        '    HTTP_From = cmc.main_HTTP_From
'        'End Property
'        '    '
'        '    ' ServerPathPage
'        '    '
'        '    Public Property Let ServerPathPage(ByVal vNewValue As String)
'        '    cmc.main_ServerPathPage = vNewValue
'        'End Property
'        '    '
'        '    ' ServerPathPage
'        '    '
'        '    Public Property Get ServerPathPage() As String
'        '    ServerPathPage = cmc.main_ServerPathPage
'        'End Property
'        '    '
'        '    ' ServerReferrer
'        '    '
'        '    Public Property Let ServerReferrer(ByVal vNewValue As String)
'        '    cmc.main_ServerReferrer = vNewValue
'        'End Property
'        '    '
'        '    ' ServerReferrer
'        '    '
'        '    Public Property Get ServerReferrer() As String
'        '    ServerReferrer = cmc.main_ServerReferrer
'        'End Property
'        '    '
'        '    ' ServerHost
'        '    '
'        '    Public Property Let ServerHost(ByVal vNewValue As String)
'        '    cmc.main_ServerHost = vNewValue
'        'End Property
'        '    '
'        '    ' ServerHost
'        '    '
'        '    Public Property Get ServerHost() As String
'        '    ServerHost = cmc.main_ServerHost
'        'End Property
'        '    '
'        '    ' ServerPageSecure
'        '    '
'        '    Public Property Let ServerPageSecure(ByVal vNewValue As Boolean)
'        '    cmc.main_ServerPageSecure = vNewValue
'        'End Property
'        '    '
'        '    ' ServerPageSecure
'        '    '
'        '    Public Property Get ServerPageSecure() As Boolean
'        '    ServerPageSecure = cmc.main_ServerPageSecure
'        'End Property
'        '    '
'        '    ' PhysicalWWWPath
'        '    '
'        '    Public Property Let PhysicalWWWPath(ByVal vNewValue As String)
'        '    cmc.main_PhysicalWWWPath = vNewValue
'        'End Property
'        '    '
'        '    ' PhysicalWWWPath
'        '    '
'        '    Public Property Get PhysicalWWWPath() As String
'        '    PhysicalWWWPath = cmc.main_PhysicalWWWPath
'        'End Property
'        '    '
'        '    ' PhysicalccLibPath
'        '    '
'        '    Public Property Let PhysicalccLibPath(ByVal vNewValue As String)
'        '    cmc.main_PhysicalccLibPath = vNewValue
'        'End Property
'        '    '
'        '    ' PhysicalccLibPath
'        '    '
'        '    Public Property Get PhysicalccLibPath() As String
'        '    PhysicalccLibPath = cmc.main_PhysicalccLibPath
'        'End Property
'        '    '
'        '    ' VisitRemoteIP
'        '    '
'        '    Public Property Let VisitRemoteIP(ByVal vNewValue As String)
'        '    cmc.main_VisitRemoteIP = vNewValue
'        'End Property
'        '    '
'        '    ' VisitRemoteIP
'        '    '
'        '    Public Property Get VisitRemoteIP() As String
'        '    VisitRemoteIP = cmc.main_VisitRemoteIP
'        'End Property
'        '    '
'        '    ' VisitBrowser
'        '    '
'        '    Public Property Let VisitBrowser(ByVal vNewValue As String)
'        '    cmc.main_VisitBrowser = vNewValue
'        'End Property
'        '    '
'        '    ' VisitBrowser
'        '    '
'        '    Public Property Get VisitBrowser() As String
'        '    VisitBrowser = cmc.main_VisitBrowser
'        'End Property
'        '    '
'        '    ' ServerQueryString
'        '    '
'        '    Public Property Let ServerQueryString(ByVal vNewValue As String)
'        '    cmc.main_ServerQueryString = vNewValue
'        'End Property
'        '    '
'        '    ' ServerQueryString
'        '    '
'        '    Public Property Get ServerQueryString() As String
'        '    ServerQueryString = cmc.main_ServerQueryString
'        'End Property
'        '    '
'        '    ' ReadStreamBinaryRead
'        '    '
'        '    Public Property Let ReadStreamBinaryRead(ByVal vNewValue As Boolean)
'        '    cmc.main_ReadStreamBinaryRead = vNewValue
'        'End Property
'        '    '
'        '    ' ReadStreamBinaryRead
'        '    '
'        '    Public Property Get ReadStreamBinaryRead() As Boolean
'        '    ReadStreamBinaryRead = cmc.main_ReadStreamBinaryRead
'        'End Property
'        '    '
'        '    ' ServerBinaryHeader
'        '    '
'        '    Public Property Set ServerBinaryHeader(vNewValue As Variant)
'        '    cmc.main_ServerBinaryHeader = vNewValue
'        'End Property
'        '    '
'        '    ' ServerBinaryHeader
'        '    '
'        '    Public Property Let ServerBinaryHeader(vNewValue As Variant)
'        '    cmc.main_ServerBinaryHeader = vNewValue
'        'End Property
'        '    '
'        '    ' ServerBinaryHeader
'        '    '
'        '    Public Property Get ServerBinaryHeader() As Variant
'        '    ServerBinaryHeader = cmc.main_ServerBinaryHeader
'        'End Property
'        '    '
'        '    ' ServerForm
'        '    '
'        '    Public Property Let ServerForm(ByVal vNewValue As String)
'        '    cmc.main_ServerForm = vNewValue
'        'End Property
'        '    '
'        '    ' ServerForm
'        '    '
'        '    Public Property Get ServerForm() As String
'        '    ServerForm = cmc.main_ServerForm
'        'End Property
'        '    '
'        '    ' ServerFormFiles
'        '    '
'        '    Public Property Let ServerFormFiles(ByVal vNewValue As String)
'        '    cmc.main_ServerFormFiles = vNewValue
'        'End Property
'        '    '
'        '    ' ServerFormFiles
'        '    '
'        '    Public Property Get ServerFormFiles() As String
'        '    ServerFormFiles = cmc.main_ServerFormFiles
'        'End Property
'        '    '
'        '    ' ServerCookies
'        '    '
'        '    Public Property Let ServerCookies(ByVal vNewValue As String)
'        '    cmc.main_ServerCookies = vNewValue
'        'End Property
'        '    '
'        '    ' ServerCookies
'        '    '
'        '    Public Property Get ServerCookies() As String
'        '    ServerCookies = cmc.main_ServerCookies
'        'End Property
'        '    '
'        '    ' InStreamSpaceAsUnderscore
'        '    '
'        '    Public Property Let InStreamSpaceAsUnderscore(ByVal vNewValue As Boolean)
'        '    cmc.main_InStreamSpaceAsUnderscore = vNewValue
'        'End Property
'        '    '
'        '    ' InStreamSpaceAsUnderscore
'        '    '
'        '    Public Property Get InStreamSpaceAsUnderscore() As Boolean
'        '    InStreamSpaceAsUnderscore = cmc.main_InStreamSpaceAsUnderscore
'        'End Property
'        '    '
'        '    ' InStreamDotAsUnderscore
'        '    '
'        '    Public Property Let InStreamDotAsUnderscore(ByVal vNewValue As Boolean)
'        '    cmc.main_InStreamDotAsUnderscore = vNewValue
'        'End Property
'        '    '
'        '    ' InStreamDotAsUnderscore
'        '    '
'        '    Public Property Get InStreamDotAsUnderscore() As Boolean
'        '    InStreamDotAsUnderscore = cmc.main_InStreamDotAsUnderscore
'        'End Property
'        '    '
'        '    ' UseASPObjects
'        '    '
'        '    Public Property Let UseASPObjects(ByVal vNewValue As Boolean)
'        '        'cmc.main_UseASPObjects = vNewValue
'        '    End Property
'        '    '
'        '    ' UseASPObjects
'        '    '
'        '    Public Property Get UseASPObjects() As Boolean
'        '        'UseASPObjects = cmc.main_UseASPObjects
'        '    End Property
'        '    '
'        '    ' OptionString
'        '    '
'        '    Public Property Let OptionString(ByVal vNewValue As String)
'        '    cmc.main_optionString = vNewValue
'        'End Property
'        '    '
'        '    ' OptionString
'        '    '
'        '    Public Property Get OptionString() As String
'        '    OptionString = cmc.main_optionString
'        'End Property
'        '    '
'        '    ' FilterInput
'        '    '
'        '    Public Property Let FilterInput(ByVal vNewValue As String)
'        '    cmc.main_FilterInput = vNewValue
'        'End Property
'        '    '
'        '    ' FilterInput
'        '    '
'        '    Public Property Get FilterInput() As String
'        '    FilterInput = cmc.main_FilterInput
'        'End Property
'        '    '
'        '    ' ServerConnectionHandle
'        '    '
'        '    Public Property Let ServerConnectionHandle(ByVal vNewValue As Long)
'        '    cmc.main_ServerConnectionHandle = vNewValue
'        'End Property
'        '    '
'        '    ' ServerConnectionHandle
'        '    '
'        '    Public Property Get ServerConnectionHandle() As Long
'        '    ServerConnectionHandle = cmc.main_ServerConnectionHandle
'        'End Property
'        '    '
'        '    ' adminMessage
'        '    '
'        '    Public Property Let adminMessage(ByVal vNewValue As String)
'        '    cmc.main_AdminMessage = vNewValue
'        'End Property
'        '    '
'        '    ' adminMessage
'        '    '
'        '    Public Property Get adminMessage() As String
'        '    adminMessage = cmc.main_AdminMessage
'        'End Property
'        '    '
'        '    ' WebClientVersion
'        '    '
'        '    Public Property Let WebClientVersion(ByVal vNewValue As String)
'        '    cmc.main_WebClientVersion = vNewValue
'        'End Property
'        '    '
'        '    ' WebClientVersion
'        '    '
'        '    Public Property Get WebClientVersion() As String
'        '    WebClientVersion = cmc.main_WebClientVersion
'        'End Property
'        '    '
'        '    ' ContentServerVersion
'        '    '
'        '    Public Property Let ContentServerVersion(ByVal vNewValue As String)
'        '    cmc.main_ContentServerVersion = vNewValue
'        'End Property
'        '    '
'        '    ' ContentServerVersion
'        '    '
'        '    Public Property Get ContentServerVersion() As String
'        '    ContentServerVersion = cmc.main_ContentServerVersion
'        'End Property
'        '    '
'        '    ' VisitID
'        '    '
'        '    Public Property Let VisitID(ByVal vNewValue As Long)
'        '    cmc.main_VisitId = vNewValue
'        'End Property
'        '    '
'        '    ' VisitID
'        '    '
'        '    Public Property Get VisitID() As Long
'        '    VisitID = cmc.main_VisitId
'        'End Property
'        '    '
'        '    ' VisitName
'        '    '
'        '    Public Property Let VisitName(ByVal vNewValue As String)
'        '    cmc.main_VisitName = vNewValue
'        'End Property
'        '    '
'        '    ' VisitName
'        '    '
'        '    Public Property Get VisitName() As String
'        '    VisitName = cmc.main_VisitName
'        'End Property
'        '    '
'        '    ' VisitStartDateValue
'        '    '
'        '    Public Property Let VisitStartDateValue(ByVal vNewValue As Long)
'        '    cmc.main_VisitStartDateValue = vNewValue
'        'End Property
'        '    '
'        '    ' VisitStartDateValue
'        '    '
'        '    Public Property Get VisitStartDateValue() As Long
'        '    VisitStartDateValue = cmc.main_VisitStartDateValue
'        'End Property
'        '    '
'        '    ' VisitStartTime
'        '    '
'        '    Public Property Let VisitStartTime(ByVal vNewValue As Date)
'        '    cmc.main_VisitStartTime = vNewValue
'        'End Property
'        '    '
'        '    ' VisitStartTime
'        '    '
'        '    Public Property Get VisitStartTime() As Date
'        '    VisitStartTime = cmc.main_VisitStartTime
'        'End Property
'        '    '
'        '    ' VisitLastTime
'        '    '
'        '    Public Property Let VisitLastTime(ByVal vNewValue As Date)
'        '    cmc.main_VisitLastTime = vNewValue
'        'End Property
'        '    '
'        '    ' VisitLastTime
'        '    '
'        '    Public Property Get VisitLastTime() As Date
'        '    VisitLastTime = cmc.main_VisitLastTime
'        'End Property
'        '    '
'        '    ' VisitCookieSupport
'        '    '
'        '    Public Property Let VisitCookieSupport(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitCookieSupport = vNewValue
'        'End Property
'        '    '
'        '    ' VisitCookieSupport
'        '    '
'        '    Public Property Get VisitCookieSupport() As Boolean
'        '    VisitCookieSupport = cmc.main_VisitCookieSupport
'        'End Property
'        '    '
'        '    ' VisitPages
'        '    '
'        '    Public Property Let VisitPages(ByVal vNewValue As Long)
'        '    cmc.main_VisitPages = vNewValue
'        'End Property
'        '    '
'        '    ' VisitPages
'        '    '
'        '    Public Property Get VisitPages() As Long
'        '    VisitPages = cmc.main_VisitPages
'        'End Property
'        '    '
'        '    ' VisitReferer
'        '    '
'        '    Public Property Let VisitReferer(ByVal vNewValue As String)
'        '    cmc.main_VisitReferer = vNewValue
'        'End Property
'        '    '
'        '    ' VisitReferer
'        '    '
'        '    Public Property Get VisitReferer() As String
'        '    VisitReferer = cmc.main_VisitReferer
'        'End Property
'        '    '
'        '    ' VisitRefererHost
'        '    '
'        '    Public Property Let VisitRefererHost(ByVal vNewValue As String)
'        '    cmc.main_VisitRefererHost = vNewValue
'        'End Property
'        '    '
'        '    ' VisitRefererHost
'        '    '
'        '    Public Property Get VisitRefererHost() As String
'        '    VisitRefererHost = cmc.main_VisitRefererHost
'        'End Property
'        '    '
'        '    ' VisitRefererPathPage
'        '    '
'        '    Public Property Let VisitRefererPathPage(ByVal vNewValue As String)
'        '    cmc.main_VisitRefererPathPage = vNewValue
'        'End Property
'        '    '
'        '    ' VisitRefererPathPage
'        '    '
'        '    Public Property Get VisitRefererPathPage() As String
'        '    VisitRefererPathPage = cmc.main_VisitRefererPathPage
'        'End Property
'        '    '
'        '    ' VisitLoginAttempts
'        '    '
'        '    Public Property Let VisitLoginAttempts(ByVal vNewValue As Long)
'        '    cmc.main_VisitLoginAttempts = vNewValue
'        'End Property
'        '    '
'        '    ' VisitLoginAttempts
'        '    '
'        '    Public Property Get VisitLoginAttempts() As Long
'        '    VisitLoginAttempts = cmc.main_VisitLoginAttempts
'        'End Property
'        '    '
'        '    ' VisitAuthenticated
'        '    '
'        '    Public Property Let VisitAuthenticated(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitAuthenticated = vNewValue
'        'End Property
'        '    '
'        '    ' VisitAuthenticated
'        '    '
'        '    Public Property Get VisitAuthenticated() As Boolean
'        '    VisitAuthenticated = cmc.main_VisitAuthenticated
'        'End Property
'        '    '
'        '    ' VisitBrowserIsIE
'        '    '
'        '    Public Property Let VisitBrowserIsIE(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitBrowserIsIE = vNewValue
'        'End Property
'        '    '
'        '    ' VisitBrowserIsIE
'        '    '
'        '    Public Property Get VisitBrowserIsIE() As Boolean
'        '    VisitBrowserIsIE = cmc.main_VisitBrowserIsIE
'        'End Property
'        '    '
'        '    ' VisitBrowserIsNS
'        '    '
'        '    Public Property Let VisitBrowserIsNS(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitBrowserIsNS = vNewValue
'        'End Property
'        '    '
'        '    ' VisitBrowserIsNS
'        '    '
'        '    Public Property Get VisitBrowserIsNS() As Boolean
'        '    VisitBrowserIsNS = cmc.main_VisitBrowserIsNS
'        'End Property
'        '    '
'        '    ' VisitBrowserVersion
'        '    '
'        '    Public Property Let VisitBrowserVersion(ByVal vNewValue As String)
'        '    cmc.main_VisitBrowserVersion = vNewValue
'        'End Property
'        '    '
'        '    ' VisitBrowserVersion
'        '    '
'        '    Public Property Get VisitBrowserVersion() As String
'        '    VisitBrowserVersion = cmc.main_VisitBrowserVersion
'        'End Property
'        '    '
'        '    ' VisitBrowserIsWindows
'        '    '
'        '    Public Property Let VisitBrowserIsWindows(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitBrowserIsWindows = vNewValue
'        'End Property
'        '    '
'        '    ' VisitBrowserIsWindows
'        '    '
'        '    Public Property Get VisitBrowserIsWindows() As Boolean
'        '    VisitBrowserIsWindows = cmc.main_VisitBrowserIsWindows
'        'End Property
'        '    '
'        '    ' VisitBrowserIsMac
'        '    '
'        '    Public Property Let VisitBrowserIsMac(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitBrowserIsMac = vNewValue
'        'End Property
'        '    '
'        '    ' VisitBrowserIsMac
'        '    '
'        '    Public Property Get VisitBrowserIsMac() As Boolean
'        '    VisitBrowserIsMac = cmc.main_VisitBrowserIsMac
'        'End Property
'        '    '
'        '    ' VisitBrowserIsLinux
'        '    '
'        '    Public Property Let VisitBrowserIsLinux(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitBrowserIsLinux = vNewValue
'        'End Property
'        '    '
'        '    ' VisitBrowserIsLinux
'        '    '
'        '    Public Property Get VisitBrowserIsLinux() As Boolean
'        '    VisitBrowserIsLinux = cmc.main_VisitBrowserIsLinux
'        'End Property
'        '    '
'        '    ' VisitBrowserIsMobile
'        '    '
'        '    Public Property Let VisitBrowserIsMobile(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitBrowserIsMobile = vNewValue
'        'End Property
'        '    '
'        '    ' VisitBrowserIsMobile
'        '    '
'        '    Public Property Get VisitBrowserIsMobile() As Boolean
'        '    VisitBrowserIsMobile = cmc.main_VisitBrowserIsMobile
'        'End Property
'        '    '
'        '    ' VisitExcludeFromAnalytics
'        '    '
'        '    Public Property Let VisitExcludeFromAnalytics(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitExcludeFromAnalytics = vNewValue
'        'End Property
'        '    '
'        '    ' VisitExcludeFromAnalytics
'        '    '
'        '    Public Property Get VisitExcludeFromAnalytics() As Boolean
'        '    VisitExcludeFromAnalytics = cmc.main_VisitExcludeFromAnalytics
'        'End Property
'        '    '
'        '    ' VisitIsBot
'        '    '
'        '    Public Property Let VisitIsBot(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitIsBot = vNewValue
'        'End Property
'        '    '
'        '    ' VisitIsBot
'        '    '
'        '    Public Property Get VisitIsBot() As Boolean
'        '    VisitIsBot = cmc.main_VisitIsBot
'        'End Property
'        '    '
'        '    ' VisitIsBadBot
'        '    '
'        '    Public Property Let VisitIsBadBot(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitIsBadBot = vNewValue
'        'End Property
'        '    '
'        '    ' VisitIsBadBot
'        '    '
'        '    Public Property Get VisitIsBadBot() As Boolean
'        '    VisitIsBadBot = cmc.main_VisitIsBadBot
'        'End Property
'        '    '
'        '    ' VisitorID
'        '    '
'        '    Public Property Let VisitorID(ByVal vNewValue As Long)
'        '    cmc.main_VisitorID = vNewValue
'        'End Property
'        '    '
'        '    ' VisitorID
'        '    '
'        '    Public Property Get VisitorID() As Long
'        '    VisitorID = cmc.main_VisitorID
'        'End Property
'        '    '
'        '    ' VisitorName
'        '    '
'        '    Public Property Let VisitorName(ByVal vNewValue As String)
'        '    cmc.main_VisitorName = vNewValue
'        'End Property
'        '    '
'        '    ' VisitorName
'        '    '
'        '    Public Property Get VisitorName() As String
'        '    VisitorName = cmc.main_VisitorName
'        'End Property
'        '    '
'        '    ' VisitorMemberID
'        '    '
'        '    Public Property Let VisitorMemberID(ByVal vNewValue As Long)
'        '    cmc.main_VisitorMemberID = vNewValue
'        'End Property
'        '    '
'        '    ' VisitorMemberID
'        '    '
'        '    Public Property Get VisitorMemberID() As Long
'        '    VisitorMemberID = cmc.main_VisitorMemberID
'        'End Property
'        '    '
'        '    ' VisitorOrderID
'        '    '
'        '    Public Property Let VisitorOrderID(ByVal vNewValue As Long)
'        '    cmc.main_VisitorOrderID = vNewValue
'        'End Property
'        '    '
'        '    ' VisitorOrderID
'        '    '
'        '    Public Property Get VisitorOrderID() As Long
'        '    VisitorOrderID = cmc.main_VisitorOrderID
'        'End Property
'        '    '
'        '    ' VisitorNew
'        '    '
'        '    Public Property Let VisitorNew(ByVal vNewValue As Boolean)
'        '    cmc.main_VisitorNew = vNewValue
'        'End Property
'        '    '
'        '    ' VisitorNew
'        '    '
'        '    Public Property Get VisitorNew() As Boolean
'        '    VisitorNew = cmc.main_VisitorNew
'        'End Property
'        '    '
'        '    ' VisitorForceBrowserMobile
'        '    '
'        '    Public Property Set VisitorForceBrowserMobile(vNewValue As Variant)
'        '    cmc.main_VisitorForceBrowserMobile = vNewValue
'        'End Property
'        '    '
'        '    ' VisitorForceBrowserMobile
'        '    '
'        '    Public Property Let VisitorForceBrowserMobile(vNewValue As Variant)
'        '    cmc.main_VisitorForceBrowserMobile = vNewValue
'        'End Property
'        '    '
'        '    ' VisitorForceBrowserMobile
'        '    '
'        '    Public Property Get VisitorForceBrowserMobile() As Variant
'        '    VisitorForceBrowserMobile = cmc.main_VisitorForceBrowserMobile
'        'End Property
'        '    '
'        '    ' memberID
'        '    '
'        '    Public Property Let memberID(ByVal vNewValue As Long)
'        '    cmc.main_memberID = vNewValue
'        'End Property
'        '    '
'        '    ' memberID
'        '    '
'        '    Public Property Get memberID() As Long
'        '    memberID = cmc.main_memberID
'        'End Property
'        '    '
'        '    ' MemberName
'        '    '
'        '    Public Property Let MemberName(ByVal vNewValue As String)
'        '    cmc.main_MemberName = vNewValue
'        'End Property
'        '    '
'        '    ' MemberName
'        '    '
'        '    Public Property Get MemberName() As String
'        '    MemberName = cmc.main_MemberName
'        'End Property
'        '    '
'        '    ' MemberAdmin
'        '    '
'        '    Public Property Let MemberAdmin(ByVal vNewValue As Boolean)
'        '    cmc.main_MemberAdmin = vNewValue
'        'End Property
'        '    '
'        '    ' MemberAdmin
'        '    '
'        '    Public Property Get MemberAdmin() As Boolean
'        '    MemberAdmin = cmc.main_MemberAdmin
'        'End Property
'        '    '
'        '    ' MemberDeveloper
'        '    '
'        '    Public Property Let MemberDeveloper(ByVal vNewValue As Boolean)
'        '    cmc.main_MemberDeveloper = vNewValue
'        'End Property
'        '    '
'        '    ' MemberDeveloper
'        '    '
'        '    Public Property Get MemberDeveloper() As Boolean
'        '    MemberDeveloper = cmc.main_MemberDeveloper
'        'End Property
'        '    '
'        '    ' MemberOrganizationID
'        '    '
'        '    Public Property Let MemberOrganizationID(ByVal vNewValue As Long)
'        '    cmc.main_MemberOrganizationID = vNewValue
'        'End Property
'        '    '
'        '    ' MemberOrganizationID
'        '    '
'        '    Public Property Get MemberOrganizationID() As Long
'        '    MemberOrganizationID = cmc.main_MemberOrganizationID
'        'End Property
'        '    '
'        '    ' MemberLanguageID
'        '    '
'        '    Public Property Let MemberLanguageID(ByVal vNewValue As Long)
'        '    cmc.main_MemberLanguageID = vNewValue
'        'End Property
'        '    '
'        '    ' MemberLanguageID
'        '    '
'        '    Public Property Get MemberLanguageID() As Long
'        '    MemberLanguageID = cmc.main_MemberLanguageID
'        'End Property
'        '    '
'        '    ' MemberLanguage
'        '    '
'        '    Public Property Let MemberLanguage(ByVal vNewValue As String)
'        '    cmc.main_MemberLanguage = vNewValue
'        'End Property
'        '    '
'        '    ' MemberLanguage
'        '    '
'        '    Public Property Get MemberLanguage() As String
'        '    MemberLanguage = cmc.main_MemberLanguage
'        'End Property
'        '    '
'        '    ' MemberNew
'        '    '
'        '    Public Property Let MemberNew(ByVal vNewValue As Boolean)
'        '    cmc.main_MemberNew = vNewValue
'        'End Property
'        '    '
'        '    ' MemberNew
'        '    '
'        '    Public Property Get MemberNew() As Boolean
'        '    MemberNew = cmc.main_MemberNew
'        'End Property
'        '    '
'        '    ' MemberEmail
'        '    '
'        '    Public Property Let MemberEmail(ByVal vNewValue As String)
'        '    cmc.main_MemberEmail = vNewValue
'        'End Property
'        '    '
'        '    ' MemberEmail
'        '    '
'        '    Public Property Get MemberEmail() As String
'        '    MemberEmail = cmc.main_MemberEmail
'        'End Property
'        '    '
'        '    ' MemberAllowBulkEmail
'        '    '
'        '    Public Property Let MemberAllowBulkEmail(ByVal vNewValue As Boolean)
'        '    cmc.main_MemberAllowBulkEmail = vNewValue
'        'End Property
'        '    '
'        '    ' MemberAllowBulkEmail
'        '    '
'        '    Public Property Get MemberAllowBulkEmail() As Boolean
'        '    MemberAllowBulkEmail = cmc.main_MemberAllowBulkEmail
'        'End Property
'        '    '
'        '    ' MemberAllowToolsPanel
'        '    '
'        '    Public Property Let MemberAllowToolsPanel(ByVal vNewValue As Boolean)
'        '    cmc.main_MemberAllowToolsPanel = vNewValue
'        'End Property
'        '    '
'        '    ' MemberAllowToolsPanel
'        '    '
'        '    Public Property Get MemberAllowToolsPanel() As Boolean
'        '    MemberAllowToolsPanel = cmc.main_MemberAllowToolsPanel
'        'End Property
'        '    '
'        '    ' MemberAutoLogin
'        '    '
'        '    Public Property Let MemberAutoLogin(ByVal vNewValue As Boolean)
'        '    cmc.main_MemberAutoLogin = vNewValue
'        'End Property
'        '    '
'        '    ' MemberAutoLogin
'        '    '
'        '    Public Property Get MemberAutoLogin() As Boolean
'        '    MemberAutoLogin = cmc.main_MemberAutoLogin
'        'End Property
'        '    '
'        '    ' MemberSendNotes
'        '    '
'        '    Public Property Let MemberSendNotes(ByVal vNewValue As Boolean)
'        '        'cmc.main_MemberSendNotes = vNewValue
'        '    End Property
'        '    '
'        '    ' MemberSendNotes
'        '    '
'        '    Public Property Get MemberSendNotes() As Boolean
'        '        'MemberSendNotes = cmc.main_MemberSendNotes
'        '    End Property
'        '    '
'        '    ' MemberAdminMenuModeID
'        '    '
'        '    Public Property Let MemberAdminMenuModeID(ByVal vNewValue As Long)
'        '    cmc.main_MemberAdminMenuModeID = vNewValue
'        'End Property
'        '    '
'        '    ' MemberAdminMenuModeID
'        '    '
'        '    Public Property Get MemberAdminMenuModeID() As Long
'        '    MemberAdminMenuModeID = cmc.main_MemberAdminMenuModeID
'        'End Property
'        '    '
'        '    ' MemberLoginUsername
'        '    '
'        '    Public Property Let MemberLoginUsername(ByVal vNewValue As String)
'        '    cmc.main_MemberLoginUsername = vNewValue
'        'End Property
'        '    '
'        '    ' MemberLoginUsername
'        '    '
'        '    Public Property Get MemberLoginUsername() As String
'        '    MemberLoginUsername = cmc.main_MemberLoginUsername
'        'End Property
'        '    '
'        '    ' MemberLoginPassword
'        '    '
'        '    Public Property Let MemberLoginPassword(ByVal vNewValue As String)
'        '    cmc.main_MemberLoginPassword = vNewValue
'        'End Property
'        '    '
'        '    ' MemberLoginPassword
'        '    '
'        '    Public Property Get MemberLoginPassword() As String
'        '    MemberLoginPassword = cmc.main_MemberLoginPassword
'        'End Property
'        '    '
'        '    ' MemberLoginEmail
'        '    '
'        '    Public Property Let MemberLoginEmail(ByVal vNewValue As String)
'        '    cmc.main_MemberLoginEmail = vNewValue
'        'End Property
'        '    '
'        '    ' MemberLoginEmail
'        '    '
'        '    Public Property Get MemberLoginEmail() As String
'        '    MemberLoginEmail = cmc.main_MemberLoginEmail
'        'End Property
'        '    '
'        '    ' MemberLoginAutoLogin
'        '    '
'        '    Public Property Let MemberLoginAutoLogin(ByVal vNewValue As Boolean)
'        '    cmc.main_MemberLoginAutoLogin = vNewValue
'        'End Property
'        '    '
'        '    ' MemberLoginAutoLogin
'        '    '
'        '    Public Property Get MemberLoginAutoLogin() As Boolean
'        '    MemberLoginAutoLogin = cmc.main_MemberLoginAutoLogin
'        'End Property
'        '    '
'        '    ' MemberAdded
'        '    '
'        '    Public Property Let MemberAdded(ByVal vNewValue As Boolean)
'        '    cmc.main_MemberAdded = vNewValue
'        'End Property
'        '    '
'        '    ' MemberAdded
'        '    '
'        '    Public Property Get MemberAdded() As Boolean
'        '    MemberAdded = cmc.main_MemberAdded
'        'End Property
'        '    '
'        '    ' MemberUsername
'        '    '
'        '    Public Property Let MemberUsername(ByVal vNewValue As String)
'        '    cmc.main_MemberUsername = vNewValue
'        'End Property
'        '    '
'        '    ' MemberUsername
'        '    '
'        '    Public Property Get MemberUsername() As String
'        '    MemberUsername = cmc.main_MemberUsername
'        'End Property
'        '    '
'        '    ' MemberPassword
'        '    '
'        '    Public Property Let MemberPassword(ByVal vNewValue As String)
'        '    cmc.main_MemberPassword = vNewValue
'        'End Property
'        '    '
'        '    ' MemberPassword
'        '    '
'        '    Public Property Get MemberPassword() As String
'        '    MemberPassword = cmc.main_MemberPassword
'        'End Property
'        '    '
'        '    ' MemberContentControlID
'        '    '
'        '    Public Property Let MemberContentControlID(ByVal vNewValue As Long)
'        '    cmc.main_MemberContentControlID = vNewValue
'        'End Property
'        '    '
'        '    ' MemberContentControlID
'        '    '
'        '    Public Property Get MemberContentControlID() As Long
'        '    MemberContentControlID = cmc.main_MemberContentControlID
'        'End Property
'        '    '
'        '    ' SQLTablePeople
'        '    '
'        '    Public Property Let SQLTablePeople(ByVal vNewValue As String)
'        '        'cmc.main_SQLTablePeople = vNewValue
'        '    End Property
'        '    '
'        '    ' SQLTablePeople
'        '    '
'        '    Public Property Get SQLTablePeople() As String
'        '    SQLTablePeople = "ccmembers"
'        'End Property
'        '    '
'        '    ' SQLTableMemberRules
'        '    '
'        '    Public Property Let SQLTableMemberRules(ByVal vNewValue As String)
'        '        'cmc.main_SQLTableMemberRules = vNewValue
'        '    End Property
'        '    '
'        '    ' SQLTableMemberRules
'        '    '
'        '    Public Property Get SQLTableMemberRules() As String
'        '    SQLTableMemberRules = "ccMemberRules"
'        'End Property
'        '    '
'        '    ' SQLTableOrganizations
'        '    '
'        '    Public Property Let SQLTableOrganizations(ByVal vNewValue As String)
'        '        'cmc.main_SQLTableOrganizations = vNewValue
'        '    End Property
'        '    '
'        '    ' SQLTableOrganizations
'        '    '
'        '    Public Property Get SQLTableOrganizations() As String
'        '    SQLTableOrganizations = "organizations"
'        'End Property
'        '    '
'        '    ' SQLTableGroups
'        '    '
'        '    Public Property Let SQLTableGroups(ByVal vNewValue As String)
'        '        'cmc.main_SQLTableGroups = vNewValue
'        '    End Property
'        '    '
'        '    ' SQLTableGroups
'        '    '
'        '    Public Property Get SQLTableGroups() As String
'        '    SQLTableGroups = "ccgroups"
'        'End Property
'        '    '
'        '    ' MemberAction
'        '    '
'        '    Public Property Let MemberAction(ByVal vNewValue As Long)
'        '    cmc.main_MemberAction = vNewValue
'        'End Property
'        '    '
'        '    ' MemberAction
'        '    '
'        '    Public Property Get MemberAction() As Long
'        '    MemberAction = cmc.main_MemberAction
'        'End Property
'        '    '
'        '    ' MemberButton
'        '    '
'        '    Public Property Let MemberButton(ByVal vNewValue As String)
'        '    cmc.main_MemberButton = vNewValue
'        'End Property
'        '    '
'        '    ' MemberButton
'        '    '
'        '    Public Property Get MemberButton() As String
'        '    MemberButton = cmc.main_MemberButton
'        'End Property
'        '    '
'        '    ' MemberBillEmail
'        '    '
'        '    Public Property Let MemberBillEmail(ByVal vNewValue As String)
'        '    cmc.main_MemberBillEmail = vNewValue
'        'End Property
'        '    '
'        '    ' MemberBillEmail
'        '    '
'        '    Public Property Get MemberBillEmail() As String
'        '    MemberBillEmail = cmc.main_MemberBillEmail
'        'End Property
'        '    '
'        '    ' MemberBillPhone
'        '    '
'        '    Public Property Let MemberBillPhone(ByVal vNewValue As String)
'        '    cmc.main_MemberBillPhone = vNewValue
'        'End Property
'        '    '
'        '    ' MemberBillPhone
'        '    '
'        '    Public Property Get MemberBillPhone() As String
'        '    MemberBillPhone = cmc.main_MemberBillPhone
'        'End Property
'        '    '
'        '    ' MemberBillFax
'        '    '
'        '    Public Property Let MemberBillFax(ByVal vNewValue As String)
'        '    cmc.main_MemberBillFax = vNewValue
'        'End Property
'        '    '
'        '    ' MemberBillFax
'        '    '
'        '    Public Property Get MemberBillFax() As String
'        '    MemberBillFax = cmc.main_MemberBillFax
'        'End Property
'        '    '
'        '    ' MemberBillCompany
'        '    '
'        '    Public Property Let MemberBillCompany(ByVal vNewValue As String)
'        '    cmc.main_MemberBillCompany = vNewValue
'        'End Property
'        '    '
'        '    ' MemberBillCompany
'        '    '
'        '    Public Property Get MemberBillCompany() As String
'        '    MemberBillCompany = cmc.main_MemberBillCompany
'        'End Property
'        '    '
'        '    ' MemberBillAddress
'        '    '
'        '    Public Property Let MemberBillAddress(ByVal vNewValue As String)
'        '    cmc.main_MemberBillAddress = vNewValue
'        'End Property
'        '    '
'        '    ' MemberBillAddress
'        '    '
'        '    Public Property Get MemberBillAddress() As String
'        '    MemberBillAddress = cmc.main_MemberBillAddress
'        'End Property
'        '    '
'        '    ' MemberBillCity
'        '    '
'        '    Public Property Let MemberBillCity(ByVal vNewValue As String)
'        '    cmc.main_MemberBillCity = vNewValue
'        'End Property
'        '    '
'        '    ' MemberBillCity
'        '    '
'        '    Public Property Get MemberBillCity() As String
'        '    MemberBillCity = cmc.main_MemberBillCity
'        'End Property
'        '    '
'        '    ' MemberBillState
'        '    '
'        '    Public Property Let MemberBillState(ByVal vNewValue As String)
'        '    cmc.main_MemberBillState = vNewValue
'        'End Property
'        '    '
'        '    ' MemberBillState
'        '    '
'        '    Public Property Get MemberBillState() As String
'        '    MemberBillState = cmc.main_MemberBillState
'        'End Property
'        '    '
'        '    ' MemberBillZip
'        '    '
'        '    Public Property Let MemberBillZip(ByVal vNewValue As String)
'        '    cmc.main_MemberBillZip = vNewValue
'        'End Property
'        '    '
'        '    ' MemberBillZip
'        '    '
'        '    Public Property Get MemberBillZip() As String
'        '    MemberBillZip = cmc.main_MemberBillZip
'        'End Property
'        '    '
'        '    ' MemberBillCountry
'        '    '
'        '    Public Property Let MemberBillCountry(ByVal vNewValue As String)
'        '    cmc.main_MemberBillCountry = vNewValue
'        'End Property
'        '    '
'        '    ' MemberBillCountry
'        '    '
'        '    Public Property Get MemberBillCountry() As String
'        '    MemberBillCountry = cmc.main_MemberBillCountry
'        'End Property
'        '    '
'        '    ' MemberShipName
'        '    '
'        '    Public Property Let MemberShipName(ByVal vNewValue As String)
'        '    cmc.main_MemberShipName = vNewValue
'        'End Property
'        '    '
'        '    ' MemberShipName
'        '    '
'        '    Public Property Get MemberShipName() As String
'        '    MemberShipName = cmc.main_MemberShipName
'        'End Property
'        '    '
'        '    ' MemberShipCompany
'        '    '
'        '    Public Property Let MemberShipCompany(ByVal vNewValue As String)
'        '    cmc.main_MemberShipCompany = vNewValue
'        'End Property
'        '    '
'        '    ' MemberShipCompany
'        '    '
'        '    Public Property Get MemberShipCompany() As String
'        '    MemberShipCompany = cmc.main_MemberShipCompany
'        'End Property
'        '    '
'        '    ' MemberShipAddress
'        '    '
'        '    Public Property Let MemberShipAddress(ByVal vNewValue As String)
'        '    cmc.main_MemberShipAddress = vNewValue
'        'End Property
'        '    '
'        '    ' MemberShipAddress
'        '    '
'        '    Public Property Get MemberShipAddress() As String
'        '    MemberShipAddress = cmc.main_MemberShipAddress
'        'End Property
'        '    '
'        '    ' MemberShipCity
'        '    '
'        '    Public Property Let MemberShipCity(ByVal vNewValue As String)
'        '    cmc.main_MemberShipCity = vNewValue
'        'End Property
'        '    '
'        '    ' MemberShipCity
'        '    '
'        '    Public Property Get MemberShipCity() As String
'        '    MemberShipCity = cmc.main_MemberShipCity
'        'End Property
'        '    '
'        '    ' MemberShipState
'        '    '
'        '    Public Property Let MemberShipState(ByVal vNewValue As String)
'        '    cmc.main_MemberShipState = vNewValue
'        'End Property
'        '    '
'        '    ' MemberShipState
'        '    '
'        '    Public Property Get MemberShipState() As String
'        '    MemberShipState = cmc.main_MemberShipState
'        'End Property
'        '    '
'        '    ' MemberShipZip
'        '    '
'        '    Public Property Let MemberShipZip(ByVal vNewValue As String)
'        '    cmc.main_MemberShipZip = vNewValue
'        'End Property
'        '    '
'        '    ' MemberShipZip
'        '    '
'        '    Public Property Get MemberShipZip() As String
'        '    MemberShipZip = cmc.main_MemberShipZip
'        'End Property
'        '    '
'        '    ' MemberShipCountry
'        '    '
'        '    Public Property Let MemberShipCountry(ByVal vNewValue As String)
'        '    cmc.main_MemberShipCountry = vNewValue
'        'End Property
'        '    '
'        '    ' MemberShipCountry
'        '    '
'        '    Public Property Get MemberShipCountry() As String
'        '    MemberShipCountry = cmc.main_MemberShipCountry
'        'End Property
'        '    '
'        '    ' MemberShipPhone
'        '    '
'        '    Public Property Let MemberShipPhone(ByVal vNewValue As String)
'        '    cmc.main_MemberShipPhone = vNewValue
'        'End Property
'        '    '
'        '    ' MemberShipPhone
'        '    '
'        '    Public Property Get MemberShipPhone() As String
'        '    MemberShipPhone = cmc.main_MemberShipPhone
'        'End Property
'        '    '
'        '    ' allowDebugLog
'        '    '
'        '    Public Property Let allowDebugLog(ByVal vNewValue As Boolean)
'        '    cmc.main_allowDebugLog = vNewValue
'        'End Property
'        '    '
'        '    ' allowDebugLog
'        '    '
'        '    Public Property Get allowDebugLog() As Boolean
'        '    allowDebugLog = cmc.main_allowDebugLog
'        'End Property
'        '    '
'        '    ' BlockNotAvailableMessage
'        '    '
'        '    Public Property Let BlockNotAvailableMessage(ByVal vNewValue As Boolean)
'        '    cmc.main_BlockNotAvailableMessage = vNewValue
'        'End Property
'        '    '
'        '    ' BlockNotAvailableMessage
'        '    '
'        '    Public Property Get BlockNotAvailableMessage() As Boolean
'        '    BlockNotAvailableMessage = cmc.main_BlockNotAvailableMessage
'        'End Property
'        '    '
'        '    ' PageStartTime
'        '    '
'        '    Public Property Let PageStartTime(ByVal vNewValue As Date)
'        '    cmc.main_PageStartTime = vNewValue
'        'End Property
'        '    '
'        '    ' PageStartTime
'        '    '
'        '    Public Property Get PageStartTime() As Date
'        '    PageStartTime = cmc.main_PageStartTime
'        'End Property
'        '    '
'        '    ' BlockHandleErrorReporting
'        '    '
'        '    Public Property Let BlockHandleErrorReporting(ByVal vNewValue As Boolean)
'        '    cmc.main_BlockHandleErrorReporting = vNewValue
'        'End Property
'        '    '
'        '    ' BlockHandleErrorReporting
'        '    '
'        '    Public Property Get BlockHandleErrorReporting() As Boolean
'        '    BlockHandleErrorReporting = cmc.main_BlockHandleErrorReporting
'        'End Property
'        '    '
'        '    ' LoadFault
'        '    '
'        '    Public Property Let LoadFault(ByVal vNewValue As Boolean)
'        '    cmc.main_LoadFault = vNewValue
'        'End Property
'        '    '
'        '    ' LoadFault
'        '    '
'        '    Public Property Get LoadFault() As Boolean
'        '    LoadFault = cmc.main_LoadFault
'        'End Property
'        '    '
'        '    ' ForceUpgrade
'        '    '
'        '    Public Property Let ForceUpgrade(ByVal vNewValue As Boolean)
'        '    cmc.main_ForceUpgrade = vNewValue
'        'End Property
'        '    '
'        '    ' ForceUpgrade
'        '    '
'        '    Public Property Get ForceUpgrade() As Boolean
'        '    ForceUpgrade = cmc.main_ForceUpgrade
'        'End Property
'        '    '
'        '    ' ForceTrap
'        '    '
'        '    Public Property Let ForceTrap(ByVal vNewValue As Boolean)
'        '    cmc.main_ForceTrap = vNewValue
'        'End Property
'        '    '
'        '    ' ForceTrap
'        '    '
'        '    Public Property Get ForceTrap() As Boolean
'        '    ForceTrap = cmc.main_ForceTrap
'        'End Property
'        '    '
'        '    ' PageErrorCount
'        '    '
'        '    Public Property Let PageErrorCount(ByVal vNewValue As Long)
'        '    cmc.main_PageErrorCount = vNewValue
'        'End Property
'        '    '
'        '    ' PageErrorCount
'        '    '
'        '    Public Property Get PageErrorCount() As Long
'        '    PageErrorCount = cmc.main_PageErrorCount
'        'End Property
'        '    '
'        '    ' ErrNumber
'        '    '
'        '    Public Property Let ErrNumber(ByVal vNewValue As Long)
'        '        'cmc.main_ErrNumber = vNewValue
'        '    End Property
'        '    '
'        '    ' ErrNumber
'        '    '
'        '    Public Property Get ErrNumber() As Long
'        '        'ErrNumber = cmc.main_ErrNumber
'        '    End Property
'        '    '
'        '    ' ErrDescription
'        '    '
'        '    Public Property Let ErrDescription(ByVal vNewValue As String)
'        '        'cmc.main_ErrDescription = vNewValue
'        '    End Property
'        '    '
'        '    ' ErrDescription
'        '    '
'        '    Public Property Get ErrDescription() As String
'        '        'ErrDescription = cmc.main_ErrDescription
'        '    End Property
'        '    '
'        '    ' ErrSource
'        '    '
'        '    Public Property Let ErrSource(ByVal vNewValue As String)
'        '        'cmc.main_ErrSource = vNewValue
'        '    End Property
'        '    '
'        '    ' ErrSource
'        '    '
'        '    Public Property Get ErrSource() As String
'        '        'ErrSource = cmc.main_ErrSource
'        '    End Property
'        '    '
'        '    ' TimerTicks
'        '    '
'        '    Public Property Let TimerTicks(ByVal vNewValue As Long)
'        '        'cmc.main_TimerTicks = vNewValue
'        '    End Property
'        '    '
'        '    ' TimerTicks
'        '    '
'        '    Public Property Get TimerTicks() As Long
'        '        'TimerTicks = cmc.main_TimerTicks
'        '    End Property
'        '    '
'        '    ' DebugMode
'        '    '
'        '    Public Property Let DebugMode(ByVal vNewValue As Boolean)
'        '    cmc.main_DebugMode = vNewValue
'        'End Property
'        '    '
'        '    ' DebugMode
'        '    '
'        '    Public Property Get DebugMode() As Boolean
'        '    DebugMode = cmc.main_DebugMode
'        'End Property
'        '    '
'        '    ' BlockClosePageCopyright
'        '    '
'        '    Public Property Let BlockClosePageCopyright(ByVal vNewValue As Boolean)
'        '    cmc.main_BlockClosePageCopyright = vNewValue
'        'End Property
'        '    '
'        '    ' BlockClosePageCopyright
'        '    '
'        '    Public Property Get BlockClosePageCopyright() As Boolean
'        '    BlockClosePageCopyright = cmc.main_BlockClosePageCopyright
'        'End Property
'        '    '
'        '    ' BlockClosePageLink
'        '    '
'        '    Public Property Let BlockClosePageLink(ByVal vNewValue As Boolean)
'        '    cmc.main_BlockClosePageLink = vNewValue
'        'End Property
'        '    '
'        '    ' BlockClosePageLink
'        '    '
'        '    Public Property Get BlockClosePageLink() As Boolean
'        '    BlockClosePageLink = cmc.main_BlockClosePageLink
'        'End Property
'        '    '
'        '    ' PageTestPointPrinting
'        '    '
'        '    Public Property Let PageTestPointPrinting(ByVal vNewValue As Boolean)
'        '    cmc.main_PageTestPointPrinting = vNewValue
'        'End Property
'        '    '
'        '    ' PageTestPointPrinting
'        '    '
'        '    Public Property Get PageTestPointPrinting() As Boolean
'        '    PageTestPointPrinting = cmc.main_PageTestPointPrinting
'        'End Property
'        '    '
'        '    ' ClosePageHTML
'        '    '
'        '    Public Property Let ClosePageHTML(ByVal vNewValue As String)
'        '    cmc.main_ClosePageHTML = vNewValue
'        'End Property
'        '    '
'        '    ' ClosePageHTML
'        '    '
'        '    Public Property Get ClosePageHTML() As String
'        '    ClosePageHTML = cmc.main_ClosePageHTML
'        'End Property
'        '    '
'        '    ' OrderItemCount
'        '    '
'        '    Public Property Let OrderItemCount(ByVal vNewValue As Long)
'        '        ''cmc.main_OrderItemCount = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderItemCount
'        '    '
'        '    Public Property Get OrderItemCount() As Long
'        '        'OrderItemCount = 'cmc.main_OrderItemCount
'        '    End Property
'        '    '
'        '    ' OrderAuthorize
'        '    '
'        '    Public Property Let OrderAuthorize(ByVal vNewValue As Boolean)
'        '        ''cmc.main_OrderAuthorize = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAuthorize
'        '    '
'        '    Public Property Get OrderAuthorize() As Boolean
'        '        'OrderAuthorize = 'cmc.main_OrderAuthorize
'        '    End Property
'        '    '
'        '    ' OrderAuthorized
'        '    '
'        '    Public Property Let OrderAuthorized(ByVal vNewValue As Boolean)
'        '        ''cmc.main_OrderAuthorized = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAuthorized
'        '    '
'        '    Public Property Get OrderAuthorized() As Boolean
'        '        ' = 'cmc.main_OrderAuthorized
'        '    End Property
'        '    '
'        '    ' OrderAuthorizeResponse
'        '    '
'        '    Public Property Let OrderAuthorizeResponse(ByVal vNewValue As String)
'        '        ''cmc.main_OrderAuthorizeResponse = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAuthorizeResponse
'        '    '
'        '    Public Property Get OrderAuthorizeResponse() As String
'        '        'OrderAuthorizeResponse = 'cmc.main_OrderAuthorizeResponse
'        '    End Property
'        '    '
'        '    ' ServerLink
'        '    '
'        '    Public Property Let ServerLink(ByVal vNewValue As String)
'        '    cmc.main_ServerLink = vNewValue
'        'End Property
'        '    '
'        '    ' ServerLink
'        '    '
'        '    Public Property Get ServerLink() As String
'        '    ServerLink = cmc.main_ServerLink
'        'End Property
'        '    '
'        '    ' ServerLinkSource
'        '    '
'        '    Public Property Let ServerLinkSource(ByVal vNewValue As String)
'        '    cmc.main_ServerLinkSource = vNewValue
'        'End Property
'        '    '
'        '    ' ServerLinkSource
'        '    '
'        '    Public Property Get ServerLinkSource() As String
'        '    ServerLinkSource = cmc.main_ServerLinkSource
'        'End Property
'        '    '
'        '    ' ServerVirtualPath
'        '    '
'        '    Public Property Let ServerVirtualPath(ByVal vNewValue As String)
'        '    cmc.main_ServerVirtualPath = vNewValue
'        'End Property
'        '    '
'        '    ' ServerVirtualPath
'        '    '
'        '    Public Property Get ServerVirtualPath() As String
'        '    ServerVirtualPath = cmc.main_ServerVirtualPath
'        'End Property
'        '    '
'        '    ' ServerDomainPrimary
'        '    '
'        '    Public Property Let ServerDomainPrimary(ByVal vNewValue As String)
'        '    cmc.main_ServerDomainPrimary = vNewValue
'        'End Property
'        '    '
'        '    ' ServerDomainPrimary
'        '    '
'        '    Public Property Get ServerDomainPrimary() As String
'        '    ServerDomainPrimary = cmc.main_ServerDomainPrimary
'        'End Property
'        '    '
'        '    ' ServerDomain
'        '    '
'        '    Public Property Let ServerDomain(ByVal vNewValue As String)
'        '    cmc.main_ServerDomain = vNewValue
'        'End Property
'        '    '
'        '    ' ServerDomain
'        '    '
'        '    Public Property Get ServerDomain() As String
'        '    ServerDomain = cmc.main_ServerDomain
'        'End Property
'        '    '
'        '    ' ServerPath
'        '    '
'        '    Public Property Let ServerPath(ByVal vNewValue As String)
'        '    cmc.main_ServerPath = vNewValue
'        'End Property
'        '    '
'        '    ' ServerPath
'        '    '
'        '    Public Property Get ServerPath() As String
'        '    ServerPath = cmc.main_ServerPath
'        'End Property
'        '    '
'        '    ' ServerPage
'        '    '
'        '    Public Property Let ServerPage(ByVal vNewValue As String)
'        '    cmc.main_ServerPage = vNewValue
'        'End Property
'        '    '
'        '    ' ServerPage
'        '    '
'        '    Public Property Get ServerPage() As String
'        '    ServerPage = cmc.main_ServerPage
'        'End Property
'        '    '
'        '    ' ServerAppRootPath
'        '    '
'        '    Public Property Let ServerAppRootPath(ByVal vNewValue As String)
'        '    cmc.main_ServerAppRootPath = vNewValue
'        'End Property
'        '    '
'        '    ' ServerAppRootPath
'        '    '
'        '    Public Property Get ServerAppRootPath() As String
'        '    ServerAppRootPath = cmc.main_ServerAppRootPath
'        'End Property
'        '    '
'        '    ' ServerAppPath
'        '    '
'        '    Public Property Let ServerAppPath(ByVal vNewValue As String)
'        '    cmc.main_ServerAppPath = vNewValue
'        'End Property
'        '    '
'        '    ' ServerAppPath
'        '    '
'        '    Public Property Get ServerAppPath() As String
'        '    ServerAppPath = cmc.main_ServerAppPath
'        'End Property
'        '    '
'        '    ' serverFilePath
'        '    '
'        '    Public Property Let serverFilePath(ByVal vNewValue As String)
'        '    cmc.main_serverFilePath = vNewValue
'        'End Property
'        '    '
'        '    ' serverFilePath
'        '    '
'        '    Public Property Get serverFilePath() As String
'        '    serverFilePath = cmc.main_serverFilePath
'        'End Property
'        '    '
'        '    ' ServerSecureURLRoot
'        '    '
'        '    Public Property Let ServerSecureURLRoot(ByVal vNewValue As String)
'        '    cmc.main_ServerSecureURLRoot = vNewValue
'        'End Property
'        '    '
'        '    ' ServerSecureURLRoot
'        '    '
'        '    Public Property Get ServerSecureURLRoot() As String
'        '    ServerSecureURLRoot = cmc.main_ServerSecureURLRoot
'        'End Property
'        '    '
'        '    ' ServerFormActionURL
'        '    '
'        '    Public Property Let ServerFormActionURL(ByVal vNewValue As String)
'        '    cmc.main_ServerFormActionURL = vNewValue
'        'End Property
'        '    '
'        '    ' ServerFormActionURL
'        '    '
'        '    Public Property Get ServerFormActionURL() As String
'        '    ServerFormActionURL = cmc.main_ServerFormActionURL
'        'End Property
'        '    '
'        '    ' ServerPageDefault
'        '    '
'        '    Public Property Let ServerPageDefault(ByVal vNewValue As String)
'        '    cmc.main_ServerPageDefault = vNewValue
'        'End Property
'        '    '
'        '    ' ServerPageDefault
'        '    '
'        '    Public Property Get ServerPageDefault() As String
'        '    ServerPageDefault = cmc.main_ServerPageDefault
'        'End Property
'        '    '
'        '    ' ServerPagePrintVersion
'        '    '
'        '    Public Property Let ServerPagePrintVersion(ByVal vNewValue As Boolean)
'        '    cmc.main_ServerPagePrintVersion = vNewValue
'        'End Property
'        '    '
'        '    ' ServerPagePrintVersion
'        '    '
'        '    Public Property Get ServerPagePrintVersion() As Boolean
'        '    ServerPagePrintVersion = cmc.main_ServerPagePrintVersion
'        'End Property
'        '    '
'        '    ' ServerContentWatchPrefix
'        '    '
'        '    Public Property Let ServerContentWatchPrefix(ByVal vNewValue As String)
'        '    cmc.main_ServerContentWatchPrefix = vNewValue
'        'End Property
'        '    '
'        '    ' ServerContentWatchPrefix
'        '    '
'        '    Public Property Get ServerContentWatchPrefix() As String
'        '    ServerContentWatchPrefix = cmc.main_ServerContentWatchPrefix
'        'End Property
'        '    '
'        '    ' ServerProtocol
'        '    '
'        '    Public Property Let ServerProtocol(ByVal vNewValue As String)
'        '    cmc.main_ServerProtocol = vNewValue
'        'End Property
'        '    '
'        '    ' ServerProtocol
'        '    '
'        '    Public Property Get ServerProtocol() As String
'        '    ServerProtocol = cmc.main_ServerProtocol
'        'End Property
'        '    '
'        '    ' ServerMultiDomainMode
'        '    '
'        '    Public Property Let ServerMultiDomainMode(ByVal vNewValue As Boolean)
'        '    cmc.main_ServerMultiDomainMode = vNewValue
'        'End Property
'        '    '
'        '    ' ServerMultiDomainMode
'        '    '
'        '    Public Property Get ServerMultiDomainMode() As Boolean
'        '    ServerMultiDomainMode = cmc.main_ServerMultiDomainMode
'        'End Property
'        '    '
'        '    ' PhysicalFilePath
'        '    '
'        '    Public Property Let PhysicalFilePath(ByVal vNewValue As String)
'        '    cmc.main_PhysicalFilePath = vNewValue
'        'End Property
'        '    '
'        '    ' PhysicalFilePath
'        '    '
'        '    Public Property Get PhysicalFilePath() As String
'        '    PhysicalFilePath = cmc.main_PhysicalFilePath
'        'End Property
'        '    '
'        '    ' AppPath
'        '    '
'        '    Public Property Let AppPath(ByVal vNewValue As String)
'        '    cmc.main_AppPath = vNewValue
'        'End Property
'        '    '
'        '    ' AppPath
'        '    '
'        '    Public Property Get AppPath() As String
'        '    AppPath = cmc.main_AppPath
'        'End Property
'        '    '
'        '    ' LinkForwardSource
'        '    '
'        '    Public Property Let LinkForwardSource(ByVal vNewValue As String)
'        '    cmc.main_LinkForwardSource = vNewValue
'        'End Property
'        '    '
'        '    ' LinkForwardSource
'        '    '
'        '    Public Property Get LinkForwardSource() As String
'        '    LinkForwardSource = cmc.main_LinkForwardSource
'        'End Property
'        '    '
'        '    ' LinkForwardError
'        '    '
'        '    Public Property Let LinkForwardError(ByVal vNewValue As String)
'        '    cmc.main_LinkForwardError = vNewValue
'        'End Property
'        '    '
'        '    ' LinkForwardError
'        '    '
'        '    Public Property Get LinkForwardError() As String
'        '    LinkForwardError = cmc.main_LinkForwardError
'        'End Property
'        '    '
'        '    ' ApplicationName
'        '    '
'        '    Public Property Let ApplicationName(ByVal vNewValue As String)
'        '    cmc.main_ApplicationName = vNewValue
'        'End Property
'        '    '
'        '    ' ApplicationName
'        '    '
'        '    Public Property Get ApplicationName() As String
'        '    ApplicationName = cmc.main_ApplicationName
'        'End Property
'        '    '
'        '    ' ImportXMLFile
'        '    '
'        '    Public Property Let ImportXMLFile(ByVal vNewValue As String)
'        '    cmc.main_ImportXMLFile = vNewValue
'        'End Property
'        '    '
'        '    ' ImportXMLFile
'        '    '
'        '    Public Property Get ImportXMLFile() As String
'        '    ImportXMLFile = cmc.main_ImportXMLFile
'        'End Property
'        '    '
'        '    ' ExportXMLFile
'        '    '
'        '    Public Property Let ExportXMLFile(ByVal vNewValue As String)
'        '    cmc.main_ExportXMLFile = vNewValue
'        'End Property
'        '    '
'        '    ' ExportXMLFile
'        '    '
'        '    Public Property Get ExportXMLFile() As String
'        '    ExportXMLFile = cmc.main_ExportXMLFile
'        'End Property
'        '    '
'        '    ' PageReferer
'        '    '
'        '    Public Property Let PageReferer(ByVal vNewValue As String)
'        '    cmc.main_PageReferer = vNewValue
'        'End Property
'        '    '
'        '    ' PageReferer
'        '    '
'        '    Public Property Get PageReferer() As String
'        '    PageReferer = cmc.main_PageReferer
'        'End Property
'        '    '
'        '    ' ServerReferer
'        '    '
'        '    Public Property Let ServerReferer(ByVal vNewValue As String)
'        '    cmc.main_ServerReferer = vNewValue
'        'End Property
'        '    '
'        '    ' ServerReferer
'        '    '
'        '    Public Property Get ServerReferer() As String
'        '    ServerReferer = cmc.main_ServerReferer
'        'End Property
'        '    '
'        '    ' PageHandle
'        '    '
'        '    Public Property Let PageHandle(ByVal vNewValue As Long)
'        '    cmc.main_PageHandle = vNewValue
'        'End Property
'        '    '
'        '    ' PageHandle
'        '    '
'        '    Public Property Get PageHandle() As Long
'        '    PageHandle = cmc.main_PageHandle
'        'End Property
'        '    '
'        '    ' StreamOpen
'        '    '
'        '    Public Property Let StreamOpen(ByVal vNewValue As Boolean)
'        '    cmc.main_StreamOpen = vNewValue
'        'End Property
'        '    '
'        '    ' StreamOpen
'        '    '
'        '    Public Property Get StreamOpen() As Boolean
'        '    StreamOpen = cmc.main_StreamOpen
'        'End Property
'        '    '
'        '    ' StreamBuffered
'        '    '
'        '    Public Property Let StreamBuffered(ByVal vNewValue As Boolean)
'        '    cmc.main_StreamBuffered = vNewValue
'        'End Property
'        '    '
'        '    ' StreamBuffered
'        '    '
'        '    Public Property Get StreamBuffered() As Boolean
'        '    StreamBuffered = cmc.main_StreamBuffered
'        'End Property
'        '    '
'        '    ' LoginIconFilename
'        '    '
'        '    Public Property Let LoginIconFilename(ByVal vNewValue As String)
'        '    cmc.main_LoginIconFilename = vNewValue
'        'End Property
'        '    '
'        '    ' LoginIconFilename
'        '    '
'        '    Public Property Get LoginIconFilename() As String
'        '    LoginIconFilename = cmc.main_LoginIconFilename
'        'End Property
'        '    '
'        '    ' ReadStreamFormBlock
'        '    '
'        '    Public Property Let ReadStreamFormBlock(ByVal vNewValue As Boolean)
'        '    cmc.main_ReadStreamFormBlock = vNewValue
'        'End Property
'        '    '
'        '    ' ReadStreamFormBlock
'        '    '
'        '    Public Property Get ReadStreamFormBlock() As Boolean
'        '    ReadStreamFormBlock = cmc.main_ReadStreamFormBlock
'        'End Property
'        '    '
'        '    ' ReadStreamJSProcess
'        '    '
'        '    Public Property Let ReadStreamJSProcess(ByVal vNewValue As Boolean)
'        '    cmc.main_ReadStreamJSProcess = vNewValue
'        'End Property
'        '    '
'        '    ' ReadStreamJSProcess
'        '    '
'        '    Public Property Get ReadStreamJSProcess() As Boolean
'        '    ReadStreamJSProcess = cmc.main_ReadStreamJSProcess
'        'End Property
'        '    '
'        '    ' ReadStreamJSForm
'        '    '
'        '    Public Property Let ReadStreamJSForm(ByVal vNewValue As Boolean)
'        '    cmc.main_ReadStreamJSForm = vNewValue
'        'End Property
'        '    '
'        '    ' ReadStreamJSForm
'        '    '
'        '    Public Property Get ReadStreamJSForm() As Boolean
'        '    ReadStreamJSForm = cmc.main_ReadStreamJSForm
'        'End Property
'        '    '
'        '    ' AllowCookielessDetection
'        '    '
'        '    Public Property Let AllowCookielessDetection(ByVal vNewValue As Boolean)
'        '    cmc.main_AllowCookielessDetection = vNewValue
'        'End Property
'        '    '
'        '    ' AllowCookielessDetection
'        '    '
'        '    Public Property Get AllowCookielessDetection() As Boolean
'        '    AllowCookielessDetection = cmc.main_AllowCookielessDetection
'        'End Property
'        '    '
'        '    ' IconFileDefault
'        '    '
'        '    Public Property Let IconFileDefault(ByVal vNewValue As String)
'        '    cmc.main_IconFileDefault = vNewValue
'        'End Property
'        '    '
'        '    ' IconFileDefault
'        '    '
'        '    Public Property Get IconFileDefault() As String
'        '    IconFileDefault = cmc.main_IconFileDefault
'        'End Property
'        '    '
'        '    ' IconFolderClosed
'        '    '
'        '    Public Property Let IconFolderClosed(ByVal vNewValue As String)
'        '    cmc.main_IconFolderClosed = vNewValue
'        'End Property
'        '    '
'        '    ' IconFolderClosed
'        '    '
'        '    Public Property Get IconFolderClosed() As String
'        '    IconFolderClosed = cmc.main_IconFolderClosed
'        'End Property
'        '    '
'        '    ' IconFolderOpen
'        '    '
'        '    Public Property Let IconFolderOpen(ByVal vNewValue As String)
'        '    cmc.main_IconFolderOpen = vNewValue
'        'End Property
'        '    '
'        '    ' IconFolderOpen
'        '    '
'        '    Public Property Get IconFolderOpen() As String
'        '    IconFolderOpen = cmc.main_IconFolderOpen
'        'End Property
'        '    '
'        '    ' IconFolderUp
'        '    '
'        '    Public Property Let IconFolderUp(ByVal vNewValue As String)
'        '    cmc.main_IconFolderUp = vNewValue
'        'End Property
'        '    '
'        '    ' IconFolderUp
'        '    '
'        '    Public Property Get IconFolderUp() As String
'        '    IconFolderUp = cmc.main_IconFolderUp
'        'End Property
'        '    '
'        '    ' RenderedPageID
'        '    '
'        '    Public Property Let RenderedPageID(ByVal vNewValue As Long)
'        '    cmc.main_RenderedPageID = vNewValue
'        'End Property
'        '    '
'        '    ' RenderedPageID
'        '    '
'        '    Public Property Get RenderedPageID() As Long
'        '    RenderedPageID = cmc.main_RenderedPageID
'        'End Property
'        '    '
'        '    ' RenderedPageName
'        '    '
'        '    Public Property Let RenderedPageName(ByVal vNewValue As String)
'        '    cmc.main_RenderedPageName = vNewValue
'        'End Property
'        '    '
'        '    ' RenderedPageName
'        '    '
'        '    Public Property Get RenderedPageName() As String
'        '    RenderedPageName = cmc.main_RenderedPageName
'        'End Property
'        '    '
'        '    ' RenderedSectionID
'        '    '
'        '    Public Property Let RenderedSectionID(ByVal vNewValue As Long)
'        '    cmc.main_RenderedSectionID = vNewValue
'        'End Property
'        '    '
'        '    ' RenderedSectionID
'        '    '
'        '    Public Property Get RenderedSectionID() As Long
'        '    RenderedSectionID = cmc.main_RenderedSectionID
'        'End Property
'        '    '
'        '    ' RenderedSectionName
'        '    '
'        '    Public Property Let RenderedSectionName(ByVal vNewValue As String)
'        '    cmc.main_RenderedSectionName = vNewValue
'        'End Property
'        '    '
'        '    ' RenderedSectionName
'        '    '
'        '    Public Property Get RenderedSectionName() As String
'        '    RenderedSectionName = cmc.main_RenderedSectionName
'        'End Property
'        '    '
'        '    ' RenderedTemplateID
'        '    '
'        '    Public Property Let RenderedTemplateID(ByVal vNewValue As Long)
'        '    cmc.main_RenderedTemplateID = vNewValue
'        'End Property
'        '    '
'        '    ' RenderedTemplateID
'        '    '
'        '    Public Property Get RenderedTemplateID() As Long
'        '    RenderedTemplateID = cmc.main_RenderedTemplateID
'        'End Property
'        '    '
'        '    ' RenderedTemplateName
'        '    '
'        '    Public Property Let RenderedTemplateName(ByVal vNewValue As String)
'        '    cmc.main_RenderedTemplateName = vNewValue
'        'End Property
'        '    '
'        '    ' RenderedTemplateName
'        '    '
'        '    Public Property Get RenderedTemplateName() As String
'        '    RenderedTemplateName = cmc.main_RenderedTemplateName
'        'End Property
'        '    '
'        '    ' RenderedNavigationStructure
'        '    '
'        '    Public Property Let RenderedNavigationStructure(ByVal vNewValue As String)
'        '    cmc.main_RenderedNavigationStructure = vNewValue
'        'End Property
'        '    '
'        '    ' RenderedNavigationStructure
'        '    '
'        '    Public Property Get RenderedNavigationStructure() As String
'        '    RenderedNavigationStructure = cmc.main_RenderedNavigationStructure
'        'End Property
'        '    '
'        '    ' PageContent
'        '    '
'        '    Public Property Let PageContent(ByVal vNewValue As String)
'        '    cmc.main_PageContent = vNewValue
'        'End Property
'        '    '
'        '    ' PageContent
'        '    '
'        '    Public Property Get PageContent() As String
'        '    PageContent = cmc.main_PageContent
'        'End Property
'        '    '
'        '    ' ContentPageStructure
'        '    '
'        '    Public Property Let ContentPageStructure(ByVal vNewValue As String)
'        '    cmc.main_ContentPageStructure = vNewValue
'        'End Property
'        '    '
'        '    ' ContentPageStructure
'        '    '
'        '    Public Property Get ContentPageStructure() As String
'        '    ContentPageStructure = cmc.main_ContentPageStructure
'        'End Property
'        '    '
'        '    ' PCCCnt
'        '    '
'        '    Public Property Let PCCCnt(ByVal vNewValue As Long)
'        '    cmc.main_PCCCnt = vNewValue
'        'End Property
'        '    '
'        '    ' PCCCnt
'        '    '
'        '    Public Property Get PCCCnt() As Long
'        '    PCCCnt = cmc.main_PCCCnt
'        'End Property
'        '    '
'        '    ' PCCNeedsReload
'        '    '
'        '    Public Property Let PCCNeedsReload(ByVal vNewValue As Boolean)
'        '    cmc.main_PCCNeedsReload = vNewValue
'        'End Property
'        '    '
'        '    ' PCCNeedsReload
'        '    '
'        '    Public Property Get PCCNeedsReload() As Boolean
'        '    PCCNeedsReload = cmc.main_PCCNeedsReload
'        'End Property
'        '    '
'        '    ' PCC
'        '    '
'        '    Public Property Set PCC(vNewValue As Variant)
'        '    cmc.main_PCC = vNewValue
'        'End Property
'        '    '
'        '    ' PCC
'        '    '
'        '    Public Property Let PCC(vNewValue As Variant)
'        '    cmc.main_PCC = vNewValue
'        'End Property
'        '    '
'        '    ' PCC
'        '    '
'        '    Public Property Get PCC() As Variant
'        '    PCC = cmc.main_PCC
'        'End Property
'        '    '
'        '    ' PCCIDIndex
'        '    '
'        '    Public Property Set PCCIDIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        Set cmc.main_PCCIDIndex = vNewValue
'        'End Property
'        '    '
'        '    ' PCCIDIndex
'        '    '
'        '    Public Property Let PCCIDIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        Set cmc.main_PCCIDIndex = vNewValue
'        'End Property
'        '    '
'        '    ' PCCIDIndex
'        '    '
'        '    Public Property Get PCCIDIndex() As cpcom.fastIndex6Class
'        '        Set PCCIDIndex = cmc.main_PCCIDIndex
'        'End Property
'        '    '
'        '    ' PCCParentIDIndex
'        '    '
'        '    Public Property Set PCCParentIDIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        Set cmc.main_PCCParentIDIndex = vNewValue
'        'End Property
'        '    '
'        '    ' PCCParentIDIndex
'        '    '
'        '    Public Property Let PCCParentIDIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        Set cmc.main_PCCParentIDIndex = vNewValue
'        'End Property
'        '    '
'        '    ' PCCParentIDIndex
'        '    '
'        '    Public Property Get PCCParentIDIndex() As cpcom.fastIndex6Class
'        '        Set PCCParentIDIndex = cmc.main_PCCParentIDIndex
'        'End Property
'        '    '
'        '    ' PCCNameIndex
'        '    '
'        '    Public Property Set PCCNameIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        Set cmc.main_PCCNameIndex = vNewValue
'        'End Property
'        '    '
'        '    ' PCCNameIndex
'        '    '
'        '    Public Property Let PCCNameIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        Set cmc.main_PCCNameIndex = vNewValue
'        'End Property
'        '    '
'        '    ' PCCNameIndex
'        '    '
'        '    Public Property Get PCCNameIndex() As cpcom.fastIndex6Class
'        '        Set PCCNameIndex = cmc.main_PCCNameIndex
'        'End Property
'        '    '
'        '    ' PageContentIDIndex
'        '    '
'        '    Public Property Set PageContentIDIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        'cmc.main_PageContentIDIndex = vNewValue
'        '    End Property
'        '    '
'        '    ' PageContentIDIndex
'        '    '
'        '    Public Property Let PageContentIDIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        'cmc.main_PageContentIDIndex = vNewValue
'        '    End Property
'        '    '
'        '    ' PageContentIDIndex
'        '    '
'        '    Public Property Get PageContentIDIndex() As cpcom.fastIndex6Class
'        '        'PageContentIDIndex = cmc.main_PageContentIDIndex
'        '    End Property
'        '    '
'        '    ' PageContentParentIDIndex
'        '    '
'        '    Public Property Set PageContentParentIDIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        'cmc.main_PageContentParentIDIndex = vNewValue
'        '    End Property
'        '    '
'        '    ' PageContentParentIDIndex
'        '    '
'        '    Public Property Let PageContentParentIDIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        'cmc.main_PageContentParentIDIndex = vNewValue
'        '    End Property
'        '    '
'        '    ' PageContentParentIDIndex
'        '    '
'        '    Public Property Get PageContentParentIDIndex() As cpcom.fastIndex6Class
'        '        'PageContentParentIDIndex = cmc.main_PageContentParentIDIndex
'        '    End Property
'        '    '
'        '    ' PageContentNameIndex
'        '    '
'        '    Public Property Set PageContentNameIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        'cmc.main_PageContentNameIndex = vNewValue
'        '    End Property
'        '    '
'        '    ' PageContentNameIndex
'        '    '
'        '    Public Property Let PageContentNameIndex(ByVal vNewValue As cpcom.fastIndex6Class)
'        '        'cmc.main_PageContentNameIndex = vNewValue
'        '    End Property
'        '    '
'        '    ' PageContentNameIndex
'        '    '
'        '    Public Property Get PageContentNameIndex() As cpcom.fastIndex6Class
'        '        'PageContentNameIndex = cmc.main_PageContentNameIndex
'        '    End Property
'        '    '
'        '    ' version
'        '    '
'        '    Public Property Get version() As String
'        '    version = cmc.main_version
'        'End Property
'        '    '
'        '    ' TrapErrors
'        '    '
'        '    Public Property Get TrapErrors() As Boolean
'        '    TrapErrors = cmc.main_TrapErrors
'        'End Property
'        '    '
'        '    ' TrapErrors
'        '    '
'        '    Public Property Let TrapErrors(vNewValue As Boolean)
'        '    cmc.main_TrapErrors = vNewValue
'        'End Property
'        '    '
'        '    ' EmailAdmin
'        '    '
'        '    Public Property Get EmailAdmin() As String
'        '    EmailAdmin = cmc.main_EmailAdmin
'        'End Property
'        '    '
'        '    ' StateString
'        '    '
'        '    Public Property Get StateString() As String
'        '    StateString = cmc.main_StateString
'        'End Property
'        '    '
'        '    ' LicenseKeyHTMLEditor
'        '    '
'        '    Public Property Set LicenseKeyHTMLEditor(ByVal vNewValue As Variant)
'        '    cmc.main_LicenseKeyHTMLEditor = vNewValue
'        'End Property
'        '    '
'        '    ' LicenseKeyHTMLEditor
'        '    '
'        '    Public Property Let LicenseKeyHTMLEditor(ByVal vNewValue As Variant)
'        '    cmc.main_LicenseKeyHTMLEditor = vNewValue
'        'End Property
'        '    '
'        '    ' LicenseKeyOneToOne
'        '    '
'        '    Public Property Set LicenseKeyOneToOne(ByVal vNewValue As Variant)
'        '    cmc.main_LicenseKeyOneToOne = vNewValue
'        'End Property
'        '    '
'        '    ' LicenseKeyOneToOne
'        '    '
'        '    Public Property Let LicenseKeyOneToOne(ByVal vNewValue As Variant)
'        '    cmc.main_LicenseKeyOneToOne = vNewValue
'        'End Property
'        '    '
'        '    ' AllowPostProcessing
'        '    '
'        '    Public Property Set AllowPostProcessing(ByVal vNewValue As Variant)
'        '        'cmc.main_AllowPostProcessing = vNewValue
'        '    End Property
'        '    '
'        '    ' AllowPostProcessing
'        '    '
'        '    Public Property Let AllowPostProcessing(ByVal vNewValue As Variant)
'        '        'cmc.main_AllowPostProcessing = vNewValue
'        '    End Property
'        '    '
'        '    ' AllowPasswordEmail
'        '    '
'        '    Public Property Set AllowPasswordEmail(ByVal vNewValue As Variant)
'        '    cmc.main_AllowPasswordEmail = vNewValue
'        'End Property
'        '    '
'        '    ' AllowPasswordEmail
'        '    '
'        '    Public Property Let AllowPasswordEmail(ByVal vNewValue As Variant)
'        '    cmc.main_AllowPasswordEmail = vNewValue
'        'End Property
'        '    '
'        '    ' OrderID
'        '    '
'        '    Public Property Get OrderID() As Long
'        '        'OrderID = 'cmc.main_OrderID
'        '    End Property
'        '    '
'        '    ' FormInputHTMLCount
'        '    '
'        '    Public Property Get FormInputHTMLCount() As Long
'        '    FormInputHTMLCount = cmc.main_FormInputHTMLCount
'        'End Property
'        '    '
'        '    ' FormInputHTMLCount
'        '    '
'        '    Public Property Let FormInputHTMLCount(ByVal vNewValue As Long)
'        '    cmc.main_FormInputHTMLCount = vNewValue
'        'End Property
'        '    '
'        '    ' RefreshQueryString
'        '    '
'        '    Public Property Get RefreshQueryString() As String
'        '    RefreshQueryString = cmc.main_RefreshQueryString
'        'End Property
'        '    '
'        '    ' RefreshQueryString
'        '    '
'        '    Public Property Let RefreshQueryString(ByVal vNewValue As String)
'        '    cmc.main_RefreshQueryString = vNewValue
'        'End Property
'        '    '
'        '    ' FormInputWidthDefault
'        '    '
'        '    Public Property Get FormInputWidthDefault() As Long
'        '    FormInputWidthDefault = cmc.main_FormInputWidthDefault
'        'End Property
'        '    '
'        '    ' FormInputWidthDefault
'        '    '
'        '    Public Property Let FormInputWidthDefault(ByVal vNewValue As Long)
'        '    cmc.main_FormInputWidthDefault = vNewValue
'        'End Property
'        '    '
'        '    ' AllowEncodeHTML
'        '    '
'        '    Public Property Get AllowEncodeHTML() As Boolean
'        '    AllowEncodeHTML = cmc.main_AllowencodeHTML
'        'End Property
'        '    '
'        '    ' AllowEncodeHTML
'        '    '
'        '    Public Property Let AllowEncodeHTML(ByVal vNewValue As Boolean)
'        '    cmc.main_AllowencodeHTML = vNewValue
'        'End Property
'        '    '
'        '    ' SiteProperty_DefaultFormInputWidth
'        '    '
'        '    Public Property Get SiteProperty_DefaultFormInputWidth() As Long
'        '    SiteProperty_DefaultFormInputWidth = cmc.main_SiteProperty_DefaultFormInputWidth
'        'End Property
'        '    '
'        '    ' SiteProperty_SelectFieldWidthLimit
'        '    '
'        '    Public Property Get SiteProperty_SelectFieldWidthLimit() As Long
'        '    SiteProperty_SelectFieldWidthLimit = cmc.main_SiteProperty_SelectFieldWidthLimit
'        'End Property
'        '    '
'        '    ' SiteProperty_SelectFieldLimit
'        '    '
'        '    Public Property Get SiteProperty_SelectFieldLimit() As Long
'        '    SiteProperty_SelectFieldLimit = cmc.main_SiteProperty_SelectFieldLimit
'        'End Property
'        '    '
'        '    ' SiteProperty_DefaultFormInputTextHeight
'        '    '
'        '    Public Property Get SiteProperty_DefaultFormInputTextHeight() As Long
'        '    SiteProperty_DefaultFormInputTextHeight = cmc.main_SiteProperty_DefaultFormInputTextHeight
'        'End Property
'        '    '
'        '    ' SiteProperty_UseContentWatchLink
'        '    '
'        '    Public Property Get SiteProperty_UseContentWatchLink() As Boolean
'        '    SiteProperty_UseContentWatchLink = cmc.main_SiteProperty_UseContentWatchLink
'        'End Property
'        '    '
'        '    ' SiteProperty_AllowTestPointLogging
'        '    '
'        '    Public Property Get SiteProperty_AllowTestPointLogging() As Boolean
'        '    SiteProperty_AllowTestPointLogging = cmc.main_SiteProperty_AllowTestPointLogging
'        'End Property
'        '    '
'        '    ' SiteProperty_TrapErrors
'        '    '
'        '    Public Property Get SiteProperty_TrapErrors() As Boolean
'        '    SiteProperty_TrapErrors = cmc.main_SiteProperty_TrapErrors
'        'End Property
'        '    '
'        '    ' SiteProperty_EmailAdmin
'        '    '
'        '    Public Property Get SiteProperty_EmailAdmin() As String
'        '    SiteProperty_EmailAdmin = cmc.main_SiteProperty_EmailAdmin
'        'End Property
'        '    '
'        '    ' SiteProperty_Language
'        '    '
'        '    Public Property Get SiteProperty_Language() As String
'        '    SiteProperty_Language = cmc.main_SiteProperty_Language
'        'End Property
'        '    '
'        '    ' SiteProperty_AdminURL
'        '    '
'        '    Public Property Get SiteProperty_AdminURL() As String
'        '    SiteProperty_AdminURL = cmc.main_SiteProperty_AdminURL
'        'End Property
'        '    '
'        '    ' SiteProperty_CalendarYearLimit
'        '    '
'        '    Public Property Get SiteProperty_CalendarYearLimit() As Long
'        '    SiteProperty_CalendarYearLimit = cmc.main_SiteProperty_CalendarYearLimit
'        'End Property
'        '    '
'        '    ' SiteProperty_AllowChildMenuHeadline
'        '    '
'        '    Public Property Get SiteProperty_AllowChildMenuHeadline() As Boolean
'        '    SiteProperty_AllowChildMenuHeadline = cmc.main_SiteProperty_AllowChildMenuHeadline
'        'End Property
'        '    '
'        '    ' SiteProperty_DefaultFormInputHTMLHeight
'        '    '
'        '    Public Property Get SiteProperty_DefaultFormInputHTMLHeight() As Long
'        '    SiteProperty_DefaultFormInputHTMLHeight = cmc.main_SiteProperty_DefaultFormInputHTMLHeight
'        'End Property
'        '    '
'        '    ' SiteProperty_AllowWorkflowAuthoring
'        '    '
'        '    Public Property Get SiteProperty_AllowWorkflowAuthoring() As Boolean
'        '    SiteProperty_AllowWorkflowAuthoring = cmc.main_SiteProperty_AllowWorkflowAuthoring
'        'End Property
'        '    '
'        '    ' SiteProperty_AllowPathBlocking
'        '    '
'        '    Public Property Get SiteProperty_AllowPathBlocking() As Boolean
'        '    SiteProperty_AllowPathBlocking = cmc.main_SiteProperty_AllowPathBlocking
'        'End Property
'        '    '
'        '    ' VisitProperty_AllowHelpIcon
'        '    '
'        '    Public Property Get VisitProperty_AllowHelpIcon() As Boolean
'        '    VisitProperty_AllowHelpIcon = cmc.main_VisitProperty_AllowHelpIcon
'        'End Property
'        '    '
'        '    ' VisitProperty_AllowEditing
'        '    '
'        '    Public Property Get VisitProperty_AllowEditing() As Boolean
'        '    VisitProperty_AllowEditing = cmc.main_VisitProperty_AllowEditing
'        'End Property
'        '    '
'        '    ' VisitProperty_AllowLinkAuthoring
'        '    '
'        '    Public Property Get VisitProperty_AllowLinkAuthoring() As Boolean
'        '    VisitProperty_AllowLinkAuthoring = cmc.main_VisitProperty_AllowLinkAuthoring
'        'End Property
'        '    '
'        '    ' VisitProperty_AllowQuickEditor
'        '    '
'        '    Public Property Get VisitProperty_AllowQuickEditor() As Boolean
'        '    VisitProperty_AllowQuickEditor = cmc.main_VisitProperty_AllowQuickEditor
'        'End Property
'        '    '
'        '    ' VisitProperty_AllowAdvancedEditor
'        '    '
'        '    Public Property Get VisitProperty_AllowAdvancedEditor() As Boolean
'        '    VisitProperty_AllowAdvancedEditor = cmc.main_VisitProperty_AllowAdvancedEditor
'        'End Property
'        '    '
'        '    ' VisitProperty_AllowPresentationAuthoring
'        '    '
'        '    Public Property Get VisitProperty_AllowPresentationAuthoring() As Boolean
'        '    VisitProperty_AllowPresentationAuthoring = cmc.main_VisitProperty_AllowPresentationAuthoring
'        'End Property
'        '    '
'        '    ' VisitProperty_AllowWorkflowRendering
'        '    '
'        '    Public Property Get VisitProperty_AllowWorkflowRendering() As Boolean
'        '    VisitProperty_AllowWorkflowRendering = cmc.main_VisitProperty_AllowWorkflowRendering
'        'End Property
'        '    '
'        '    ' VisitProperty_AllowDebugging
'        '    '
'        '    Public Property Get VisitProperty_AllowDebugging() As Boolean
'        '    VisitProperty_AllowDebugging = cmc.main_VisitProperty_AllowDebugging
'        'End Property
'        '    '
'        '    ' CatalogAllowInventory
'        '    '
'        '    Public Property Set CatalogAllowInventory(ByVal vNewValue As Variant)
'        '        ''cmc.main_CatalogAllowInventory = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowInventory
'        '    '
'        '    Public Property Let CatalogAllowInventory(ByVal vNewValue As Variant)
'        '        ''cmc.main_CatalogAllowInventory = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowSecurePath
'        '    '
'        '    Public Property Set OrderAllowSecurePath(ByVal vNewValue As Variant)
'        '        ''cmc.main_OrderAllowSecurePath = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowSecurePath
'        '    '
'        '    Public Property Let OrderAllowSecurePath(ByVal vNewValue As Variant)
'        '        ''cmc.main_OrderAllowSecurePath = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowPersonalChecks
'        '    '
'        '    Public Property Set OrderAllowPersonalChecks(ByVal vNewValue As Variant)
'        '        ''cmc.main_OrderAllowPersonalChecks = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowPersonalChecks
'        '    '
'        '    Public Property Let OrderAllowPersonalChecks(ByVal vNewValue As Variant)
'        '        ''cmc.main_OrderAllowPersonalChecks = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCompanyChecks
'        '    '
'        '    Public Property Set OrderAllowCompanyChecks(ByVal vNewValue As Variant)
'        '        ''cmc.main_OrderAllowCompanyChecks = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCompanyChecks
'        '    '
'        '    Public Property Let OrderAllowCompanyChecks(ByVal vNewValue As Variant)
'        '        ''cmc.main_OrderAllowCompanyChecks = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCreditCardOnline
'        '    '
'        '    Public Property Set OrderAllowCreditCardOnline(ByVal vNewValue As Variant)
'        '        ''cmc.main_OrderAllowCreditCardOnline = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCreditCardOnline
'        '    '
'        '    Public Property Let OrderAllowCreditCardOnline(ByVal vNewValue As Variant)
'        '        ''cmc.main_OrderAllowCreditCardOnline = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCreditCardByPhone
'        '    '
'        '    Public Property Set OrderAllowCreditCardByPhone(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCreditCardByPhone = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCreditCardByPhone
'        '    '
'        '    Public Property Let OrderAllowCreditCardByPhone(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCreditCardByPhone = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCreditCardByFax
'        '    '
'        '    Public Property Set OrderAllowCreditCardByFax(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCreditCardByFax = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCreditCardByFax
'        '    '
'        '    Public Property Let OrderAllowCreditCardByFax(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCreditCardByFax = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowNetTerms
'        '    '
'        '    Public Property Set OrderAllowNetTerms(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowNetTerms = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowNetTerms
'        '    '
'        '    Public Property Let OrderAllowNetTerms(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowNetTerms = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCODCompanyCheck
'        '    '
'        '    Public Property Set OrderAllowCODCompanyCheck(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCODCompanyCheck = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCODCompanyCheck
'        '    '
'        '    Public Property Let OrderAllowCODCompanyCheck(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCODCompanyCheck = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCODCertifiedFunds
'        '    '
'        '    Public Property Set OrderAllowCODCertifiedFunds(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCODCertifiedFunds = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCODCertifiedFunds
'        '    '
'        '    Public Property Let OrderAllowCODCertifiedFunds(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCODCertifiedFunds = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCredit
'        '    '
'        '    Public Property Set OrderAllowCredit(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCredit = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowCredit
'        '    '
'        '    Public Property Let OrderAllowCredit(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowCredit = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowVisa
'        '    '
'        '    Public Property Set OrderAllowVisa(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowVisa = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowVisa
'        '    '
'        '    Public Property Let OrderAllowVisa(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowVisa = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowMC
'        '    '
'        '    Public Property Set OrderAllowMC(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowMC = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowMC
'        '    '
'        '    Public Property Let OrderAllowMC(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowMC = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowAmex
'        '    '
'        '    Public Property Set OrderAllowAmex(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowAmex = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowAmex
'        '    '
'        '    Public Property Let OrderAllowAmex(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowAmex = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowDiscover
'        '    '
'        '    Public Property Set OrderAllowDiscover(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowDiscover = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowDiscover
'        '    '
'        '    Public Property Let OrderAllowDiscover(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowDiscover = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogIndexFormat
'        '    '
'        '    Public Property Set CatalogIndexFormat(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogIndexFormat = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogIndexFormat
'        '    '
'        '    Public Property Let CatalogIndexFormat(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogIndexFormat = vNewValue
'        '    End Property
'        '    '
'        '    ' CataloglistingColumns
'        '    '
'        '    Public Property Set CataloglistingColumns(ByVal vNewValue As Variant)
'        '        'cmc.main_CataloglistingColumns = vNewValue
'        '    End Property
'        '    '
'        '    ' CataloglistingColumns
'        '    '
'        '    Public Property Let CataloglistingColumns(ByVal vNewValue As Variant)
'        '        'cmc.main_CataloglistingColumns = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowPricing
'        '    '
'        '    Public Property Set CatalogAllowPricing(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowPricing = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowPricing
'        '    '
'        '    Public Property Let CatalogAllowPricing(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowPricing = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowSpecialPrice
'        '    '
'        '    Public Property Set CatalogAllowSpecialPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowSpecialPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowSpecialPrice
'        '    '
'        '    Public Property Let CatalogAllowSpecialPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowSpecialPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasRetailPrice
'        '    '
'        '    Public Property Set CatalogAliasRetailPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasRetailPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasRetailPrice
'        '    '
'        '    Public Property Let CatalogAliasRetailPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasRetailPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasRegularPrice
'        '    '
'        '    Public Property Set CatalogAliasRegularPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasRegularPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasRegularPrice
'        '    '
'        '    Public Property Let CatalogAliasRegularPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasRegularPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasSpecialPrice
'        '    '
'        '    Public Property Set CatalogAliasSpecialPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasSpecialPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasSpecialPrice
'        '    '
'        '    Public Property Let CatalogAliasSpecialPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasSpecialPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasSalePrice
'        '    '
'        '    Public Property Set CatalogAliasSalePrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasSalePrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasSalePrice
'        '    '
'        '    Public Property Let CatalogAliasSalePrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasSalePrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasYourPrice
'        '    '
'        '    Public Property Set CatalogAliasYourPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasYourPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAliasYourPrice
'        '    '
'        '    Public Property Let CatalogAliasYourPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAliasYourPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowOrdering
'        '    '
'        '    Public Property Set CatalogAllowOrdering(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowOrdering = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowOrdering
'        '    '
'        '    Public Property Let CatalogAllowOrdering(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowOrdering = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogMfgPhrase
'        '    '
'        '    Public Property Set CatalogMfgPhrase(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogMfgPhrase = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogMfgPhrase
'        '    '
'        '    Public Property Let CatalogMfgPhrase(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogMfgPhrase = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowSearch
'        '    '
'        '    Public Property Set CatalogAllowSearch(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowSearch = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowSearch
'        '    '
'        '    Public Property Let CatalogAllowSearch(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowSearch = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowRetailPrice
'        '    '
'        '    Public Property Set CatalogAllowRetailPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowRetailPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' CatalogAllowRetailPrice
'        '    '
'        '    Public Property Let CatalogAllowRetailPrice(ByVal vNewValue As Variant)
'        '        'cmc.main_CatalogAllowRetailPrice = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowInventory
'        '    '
'        '    Public Property Set OrderAllowInventory(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowInventory = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderAllowInventory
'        '    '
'        '    Public Property Let OrderAllowInventory(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderAllowInventory = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderNotifyEmail
'        '    '
'        '    Public Property Set OrderNotifyEmail(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderNotifyEmail = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderNotifyEmail
'        '    '
'        '    Public Property Let OrderNotifyEmail(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderNotifyEmail = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderProcessPath
'        '    '
'        '    Public Property Set OrderProcessPath(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderProcessPath = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderProcessPath
'        '    '
'        '    Public Property Let OrderProcessPath(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderProcessPath = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderCardProcessor
'        '    '
'        '    Public Property Set OrderCardProcessor(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderCardProcessor = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderCardProcessor
'        '    '
'        '    Public Property Let OrderCardProcessor(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderCardProcessor = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactAddress2
'        '    '
'        '    Public Property Set OrderContactAddress2(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactAddress2 = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactAddress2
'        '    '
'        '    Public Property Let OrderContactAddress2(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactAddress2 = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactAddress1
'        '    '
'        '    Public Property Set OrderContactAddress1(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactAddress1 = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactAddress1
'        '    '
'        '    Public Property Let OrderContactAddress1(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactAddress1 = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactCheck
'        '    '
'        '    Public Property Set OrderContactCheck(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactCheck = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactCheck
'        '    '
'        '    Public Property Let OrderContactCheck(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactCheck = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactTime
'        '    '
'        '    Public Property Set OrderContactTime(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactTime = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactTime
'        '    '
'        '    Public Property Let OrderContactTime(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactTime = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactName
'        '    '
'        '    Public Property Set OrderContactName(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactName = vNewValue
'        '    End Property
'        '    '
'        '    ' OrderContactName
'        '    '
'        '    Public Property Let OrderContactName(ByVal vNewValue As Variant)
'        '        'cmc.main_OrderContactName = vNewValue
'        '    End Property
'        '    '
'        '    ' SQLCommandTimeout
'        '    '
'        '    Public Property Get SQLCommandTimeout() As Long
'        '    SQLCommandTimeout = cmc.main_SQLCommandTimeout
'        'End Property
'        '    '
'        '    ' SQLCommandTimeout
'        '    '
'        '    Public Property Let SQLCommandTimeout(ByVal vNewValue As Long)
'        '    cmc.main_SQLCommandTimeout = vNewValue
'        'End Property
'        '    '
'        '    ' ResponseRedirect
'        '    '
'        '    Public Property Get ResponseRedirect() As String
'        '    ResponseRedirect = cmc.main_ResponseRedirect
'        'End Property
'        '    '
'        '    ' ResponseHeader
'        '    '
'        '    Public Property Get ResponseHeader() As String
'        '    ResponseHeader = cmc.main_ResponseHeader
'        'End Property
'        '    '
'        '    ' ResponseCookies
'        '    '
'        '    Public Property Get ResponseCookies() As String
'        '    ResponseCookies = cmc.main_ResponseCookies
'        'End Property
'        '    '
'        '    ' ResponseContentType
'        '    '
'        '    Public Property Get ResponseContentType() As String
'        '    ResponseContentType = cmc.main_ResponseContentType
'        'End Property
'        '    '
'        '    ' ResponseStatus
'        '    '
'        '    Public Property Get ResponseStatus() As String
'        '    ResponseStatus = cmc.main_ResponseStatus
'        'End Property
'        '    '
'        '    ' ResponseBuffer
'        '    '
'        '    Public Property Get ResponseBuffer() As String
'        '    ResponseBuffer = cmc.main_ResponseBuffer
'        'End Property
'        '    '
'        '    ' ServerStyleTag
'        '    '
'        '    Public Property Get ServerStyleTag() As String
'        '    ServerStyleTag = cmc.main_ServerStyleTag
'        'End Property
'        '    '
'        '    ' ServerStyleTag
'        '    '
'        '    Public Property Let ServerStyleTag(ByVal vNewValue As String)
'        '    cmc.main_ServerStyleTag = vNewValue
'        'End Property
'        '    '
'        '    ' SiteProperty_AllowTemplateLinkVerification
'        '    '
'        '    Public Property Get SiteProperty_AllowTemplateLinkVerification() As Boolean
'        '    SiteProperty_AllowTemplateLinkVerification = cmc.main_SiteProperty_AllowTemplateLinkVerification
'        'End Property
'        '    '
'        '    ' SiteProperty_BuildVersion
'        '    '
'        '    Public Property Get SiteProperty_BuildVersion() As String
'        '    SiteProperty_BuildVersion = cmc.main_SiteProperty_BuildVersion
'        'End Property
'        '    '
'        '    ' MetaContentNoFollow
'        '    '
'        '    Public Property Get MetaContentNoFollow() As Boolean
'        '    MetaContentNoFollow = cmc.main_MetaContentNoFollow
'        'End Property
'        '    '
'        '    ' MetaContentNoFollow
'        '    '
'        '    Public Property Let MetaContentNoFollow(ByVal vNewValue As Boolean)
'        '    cmc.main_MetaContentNoFollow = vNewValue
'        'End Property
'        '    '
'        '    ' ToolsPanelTimerTrace
'        '    '
'        '    Public Property Get ToolsPanelTimerTrace() As String
'        '    ToolsPanelTimerTrace = cmc.main_ToolsPanelTimerTrace
'        'End Property
'        '    '
'        '    ' ToolsPanelTimerTrace
'        '    '
'        '    Public Property Let ToolsPanelTimerTrace(ByVal vNewValue As String)
'        '    cmc.main_ToolsPanelTimerTrace = vNewValue
'        'End Property
'        '    '
'        '    ' docType
'        '    '
'        '    Public Property Get docType() As String
'        '    docType = cmc.main_docType
'        'End Property
'        '    '
'        '    ' DocTypeAdmin
'        '    '
'        '    Public Property Get DocTypeAdmin() As String
'        '    DocTypeAdmin = cmc.main_DocTypeAdmin
'        'End Property
'        '    '
'        '    ' SiteStructure
'        '    '
'        '    Public Property Get SiteStructure() As String
'        '    SiteStructure = cmc.main_SiteStructure
'        'End Property
'        '    '
'        '    ' VisitProperty_AllowVerboseReporting
'        '    '
'        '    Public Property Get VisitProperty_AllowVerboseReporting() As Boolean
'        '    VisitProperty_AllowVerboseReporting = cmc.main_VisitProperty_AllowVerboseReporting
'        'End Property
'        '    '
'        '    ' ServerDomainList
'        '    '
'        '    Public Property Get ServerDomainList() As String
'        '    ServerDomainList = cmc.main_ServerDomainList
'        'End Property
'        '    '
'        '    ' ServerDomainCrossList
'        '    '
'        '    Public Property Get ServerDomainCrossList() As String
'        '    ServerDomainCrossList = cmc.main_ServerDomainCrossList
'        'End Property
'        '    '
'        '    ' ServerDomainList
'        '    '
'        '    Public Property Let ServerDomainList(ByVal vNewValue As String)
'        '    cmc.main_ServerDomainList = vNewValue
'        'End Property
'        '    Public Property Set cmcObj(newValue As Object)
'        '        Set cmc = newValue
'        'End Property
'        '    Public Sub Class_initialize()
'        '    End Sub
'        '    Public Sub Class_terminate()
'        '    Set cmc = Nothing
'        'End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Function AddEvent(HtmlId As String, DOMEvent As String, Javascript As String)
'        '        AddEvent = cmc.main_AddEvent(HtmlId, DOMEvent, Javascript)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function AddGroup(GroupName As Variant, Optional GroupCaption As Variant) As Long
'        '        AddGroup = cmc.main_AddGroup(GroupName, GroupCaption)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function AddRefreshQueryString(Name As Variant, Optional Value As Variant) As String
'        '        AddRefreshQueryString = cmc.main_AddRefreshQueryString(Name, Value)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ConvertHTML2Text(Source As Variant) As String
'        '        ConvertHTML2Text = cmc.main_ConvertHTML2Text(Source)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ConvertText2HTML(Source As Variant) As String
'        '        ConvertText2HTML = cmc.main_ConvertText2HTML(Source)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function cpNew() As Object
'        '        'Set cpNew = cmc.main_cpNew()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function CreateGuid() As String
'        '        CreateGuid = cmc.main_CreateGuid()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function CreateMember(username As Variant, Optional password As Variant, Optional email As Variant) As Long
'        '        CreateMember = cmc.main_CreateMember(username, password, email)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function CSOK(CSPointer As Variant) As Boolean
'        '        CSOK = cmc.main_CSOK(CSPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function DecodeAddonOptionArgument(EncodedArg As String)
'        '        DecodeAddonOptionArgument = cmc.main_DecodeAddonOptionArgument(EncodedArg)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function DecodeContent(Source As Variant) As String
'        '        DecodeContent = cmc.main_DecodeContent(Source)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function DecodeHTML(Source As Variant) As String
'        '        DecodeHTML = cmc.main_DecodeHTML(Source)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function DecodeKeyNumber(EncodedKey As Variant) As Long
'        '        DecodeKeyNumber = cmc.main_DecodeKeyNumber(EncodedKey)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function DecodeKeyTime(EncodedKey As Variant) As Date
'        '        DecodeKeyTime = cmc.main_DecodeKeyTime(EncodedKey)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function decodeNvaArgument(EncodedArg As String) As String
'        '        decodeNvaArgument = cmc.main_decodeNvaArgument(EncodedArg)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function DecodeUrl(ByVal sUrl As Variant) As String
'        '        DecodeUrl = cmc.main_DecodeUrl(sUrl)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function DeleteChildRecords(ContentName As String, recordId As Long, Optional ReturnListWithoutDelete As Boolean) As String
'        '        DeleteChildRecords = cmc.main_DeleteChildRecords(ContentName, recordId, ReturnListWithoutDelete)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function DumpAscii(ContentName As Variant)
'        '        DumpAscii = cmc.main_DumpAscii(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeAddonOptionArgument(argToEncode As String)
'        '        EncodeAddonOptionArgument = cmc.main_EncodeAddonOptionArgument(argToEncode)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeAppRootPath(Link As Variant) As String
'        '        EncodeAppRootPath = cmc.main_EncodeAppRootPath(Link)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeBoolean(InputValue As Variant) As Boolean
'        '        EncodeBoolean = cmc.main_EncodeBoolean(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeContent(Source As Variant, Optional ForMemberID As Variant, Optional CSFormattingContext As Variant, Optional PlainText As Boolean, Optional AddLinkEID As Boolean, Optional EncodeActiveFormatting As Boolean, Optional EncodeActiveImages As Boolean, Optional EncodeActiveEditIcons As Boolean, Optional EncodeActivePersonalization As Boolean, Optional AddAnchorQuery As String, Optional ProtocolHostString As String) As String
'        '        EncodeContent = cmc.main_EncodeContent(Source, ForMemberID, CSFormattingContext, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeContent2(Source As String, ForMemberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String) As String
'        '        EncodeContent2 = cmc.main_EncodeContent2(Source, ForMemberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeContent3(Source As String, ForMemberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean) As String
'        '        EncodeContent3 = cmc.main_EncodeContent3(Source, ForMemberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString, IsEmailContent)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeContent4(Source As String, ForMemberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean) As String
'        '        EncodeContent4 = cmc.main_EncodeContent4(Source, ForMemberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString, IsEmailContent)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeContent5(Source As String, ForMemberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean, DefaultWrapperID As Long) As String
'        '        EncodeContent5 = cmc.main_EncodeContent5(Source, ForMemberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString, IsEmailContent, DefaultWrapperID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeContent8(Source As String, ForMemberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean, DefaultWrapperID As Long, ignore_TemplateCaseOnly_Content As String, Context As AddonContextEnum) As String
'        '        EncodeContent8 = cmc.main_EncodeContent8(Source, ForMemberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString, IsEmailContent, DefaultWrapperID, ignore_TemplateCaseOnly_Content, Context)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeContent9(Source As String, personalizationPeopleId As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean, DefaultWrapperID As Long, ignore_TemplateCaseOnly_Content As String, addonContext As AddonContextEnum) As String
'        '        EncodeContent9 = cmc.main_EncodeContent9(Source, personalizationPeopleId, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString, IsEmailContent, DefaultWrapperID, ignore_TemplateCaseOnly_Content, addonContext)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeContentForWeb(Source As String, ContextContentName As String, ContextRecordID As Long, Ignore_BasePath As String, WrapperID As Long) As String
'        '        EncodeContentForWeb = cmc.main_EncodeContentForWeb(Source, ContextContentName, ContextRecordID, Ignore_BasePath, WrapperID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeCRLF(Source As Variant) As String
'        '        EncodeCRLF = cmc.main_EncodeCRLF(Source)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeDate(InputValue As Variant) As Date
'        '        EncodeDate = cmc.main_EncodeDate(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function encodeHTML(Source As Variant) As String
'        '        encodeHTML = cmc.main_encodeHTML(Source)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeInteger(InputValue As Variant) As Long
'        '        EncodeInteger = cmc.main_EncodeInteger(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeKeyNumber(Key As Variant, EncodeTime As Variant) As String
'        '        EncodeKeyNumber = cmc.main_EncodeKeyNumber(Key, EncodeTime)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeMessaging(Text As Variant) As String
'        '        EncodeMessaging = cmc.main_EncodeMessaging(Text)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeMessagingByMemberID(memberID As Variant, Text As Variant) As String
'        '        EncodeMessagingByMemberID = cmc.main_EncodeMessagingByMemberID(memberID, Text)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeMissing(InputValue As Variant, DefaultValue As Variant) As Variant
'        '        EncodeMissing = cmc.main_EncodeMissing(InputValue, DefaultValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeNumber(InputValue As Variant) As Double
'        '        EncodeNumber = cmc.main_EncodeNumber(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function encodeNvaArgument(Arg As String) As String
'        '        encodeNvaArgument = cmc.main_encodeNvaArgument(Arg)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodePage(PageSourceFilename As String) As String
'        '        EncodePage = cmc.main_EncodePage(PageSourceFilename)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodePage_old(PageSource As String)
'        '        EncodePage_old = cmc.main_EncodePage_old(PageSource)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeRequestVariable(Source As Variant) As String
'        '        EncodeRequestVariable = cmc.main_EncodeRequestVariable(Source)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeSQLBoolean(SourceBoolean As Variant) As String
'        '        EncodeSQLBoolean = cmc.main_EncodeSQLBoolean(SourceBoolean)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeSQLDate(SourceDate As Variant) As String
'        '        EncodeSQLDate = cmc.main_EncodeSQLDate(SourceDate)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeSQLNumber(SourceNumber As Variant) As String
'        '        EncodeSQLNumber = cmc.main_EncodeSQLNumber(SourceNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeSQLText(SourceText As Variant) As String
'        '        EncodeSQLText = cmc.main_EncodeSQLText(SourceText)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeText(InputValue As Variant) As String
'        '        EncodeText = cmc.main_EncodeText(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function EncodeURL(Source As Variant) As String
'        '        EncodeURL = cmc.main_EncodeURL(Source)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ErrorTrapping() As Boolean
'        '        ErrorTrapping = cmc.main_ErrorTrapping()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ExecuteAddon(addonid As Long, AddonNameOrGuid As String, OptionString As String, Context As AddonContextEnum, HostContentName As String, HostRecordID As Long, HostFieldName As String, ACInstanceID As String, DefaultWrapperID As Long) As String
'        '        ExecuteAddon = cmc.main_executeAddon(addonid, AddonNameOrGuid, OptionString, Context, HostContentName, HostRecordID, HostFieldName, ACInstanceID, DefaultWrapperID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ExecuteAddon2(AddonIDGuidOrName As String, Optional OptionString As String, Optional WrapperID As Long, Optional CPObjOrNothing As Object) As String
'        '        ExecuteAddon2 = cmc.main_executeAddon2(AddonIDGuidOrName, OptionString, WrapperID, CPObjOrNothing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ExecuteAddon3(AddonIDGuidOrName As String, Optional OptionString As String, Optional Context As AddonContextEnum, Optional CPObjOrNothing As Object) As String
'        '        ExecuteAddon3 = cmc.main_executeAddon3(AddonIDGuidOrName, OptionString, Context, CPObjOrNothing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ExecuteAddonAsProcess(AddonIDGuidOrName As String, Optional OptionString As String, Optional CPObjOrNothing As Object, Optional WaitForResults As Boolean) As String
'        '        ExecuteAddonAsProcess = cmc.main_executeAddonAsProcess(AddonIDGuidOrName, OptionString, CPObjOrNothing, WaitForResults)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function executeAddon_getAddonStylesBubble(addonid As Long, Return_DialogList As String) As String
'        '        executeAddon_getAddonStylesBubble = cmc.main_executeAddon_GetAddonStylesBubble(addonid, Return_DialogList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function executeAddon_getHelpBubble(addonid As Long, helpCopy As String, CollectionID As Long, Return_DialogList As String) As String
'        '        executeAddon_getHelpBubble = cmc.main_executeAddon_GetHelpBubble(addonid, helpCopy, CollectionID, Return_DialogList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function executeAddon_getHTMLViewerBubble(addonid As Long, HTMLSourceID As String, Return_DialogList As String) As String
'        '        executeAddon_getHTMLViewerBubble = cmc.main_executeAddon_GetHTMLViewerBubble(addonid, HTMLSourceID, Return_DialogList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function executeAddon_getInstanceBubble(AddonName As String, OptionString As String, ContentName As String, recordId As Long, FieldName As String, ACInstanceID As String, Context As AddonContextEnum, Return_DialogList As String) As String
'        '        executeAddon_getInstanceBubble = cmc.main_executeAddon_GetInstanceBubble(AddonName, OptionString, ContentName, recordId, FieldName, ACInstanceID, Context, Return_DialogList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ExecuteSQL(DataSourcename As Variant, SQL As Variant, Optional Retries As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Recordset
'        '    Set ExecuteSQL = cmc.main_ExecuteSQL(DataSourcename, SQL, Retries, PageSize, PageNumber)
'        'End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ExecuteSQLCommand(DataSourcename As String, SQL As String, Optional CommandTimeout, Optional PageSize, Optional PageNumber) As Recordset
'        '    Set ExecuteSQLCommand = cmc.main_ExecuteSQLCommand(DataSourcename, SQL, CommandTimeout, PageSize, PageNumber)
'        'End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ExecuteTraplessSQL(DataSourcename As String, SQL As String, Optional Retries, Optional PageSize, Optional PageNumber) As Recordset
'        '    Set ExecuteTraplessSQL = cmc.main_ExecuteTraplessSQL(DataSourcename, SQL, Retries, PageSize, PageNumber)
'        'End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetActiveEditor(ContentName As Variant, recordId As Variant, FieldName As String, Optional FormElements As Variant) As String
'        '        GetActiveEditor = cmc.main_GetActiveEditor(ContentName, recordId, FieldName, FormElements)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAddonContent(addonid As Long, AddonName As String, OptionString As String, Context As AddonContextEnum, ContentName As String, recordId As Long, FieldName As String, ACInstanceID As Long) As String
'        '        GetAddonContent = cmc.main_GetAddonContent(addonid, AddonName, OptionString, Context, ContentName, recordId, FieldName, ACInstanceID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        Public Function GetAddonOption(OptionName As String, OptionString As String) As String
'            Return cp.Doc.GetText(OptionName)
'        End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAdminForm(Optional Content As String) As String
'        '        GetAdminForm = cmc.main_GetAdminForm(Content)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAdminFormBody(Caption As String, ButtonListLeft As String, ButtonListRight As String, AllowAdd As Boolean, AllowDelete As Boolean, Description As String, ContentSummary As String, ContentPadding As Long, Content As String) As String
'        '        GetAdminFormBody = cmc.main_GetAdminFormBody(Caption, ButtonListLeft, ButtonListRight, AllowAdd, AllowDelete, Description, ContentSummary, ContentPadding, Content)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAdminHintWrapper(Content As Variant) As String
'        '        GetAdminHintWrapper = cmc.main_GetAdminHintWrapper(Content)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAdminPage(Optional Content As Variant) As String
'        '        GetAdminPage = cmc.main_GetAdminPage(Content)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAdminPage2(Optional Content As Variant, Optional doNotDisposeOnExit As Boolean) As String
'        '        GetAdminPage2 = cmc.main_GetAdminPage2(Content, doNotDisposeOnExit)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAggrOption(Name As String, OptionString As String) As String
'        '        GetAggrOption = cmc.main_GetAggrOption(Name, OptionString)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAsciiExport(ContentName As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As String
'        '        GetAsciiExport = cmc.main_GetAsciiExport(ContentName, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAuthoringLink(Label As String, SideCaption As String, Link As String, NewWindow As Boolean, Optional ignore0 As Boolean, Optional ignore1 As String) As String
'        '        GetAuthoringLink = cmc.main_GetAuthoringLink(Label, SideCaption, Link, NewWindow, ignore0, ignore1)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAuthoringStatusMessage(IsContentWorkflowAuthoring As Boolean, RecordEditLocked As Boolean, EditLockName As String, EditLockExpires As Date, RecordApproved As Boolean, ApprovedBy As String, RecordSubmitted As Boolean, SubmittedBy As String, RecordDeleted As Boolean, RecordInserted As Boolean, RecordModified As Boolean, ModifiedBy As String) As String
'        '        GetAuthoringStatusMessage = cmc.main_GetAuthoringStatusMessage(IsContentWorkflowAuthoring, RecordEditLocked, EditLockName, EditLockExpires, RecordApproved, ApprovedBy, RecordSubmitted, SubmittedBy, RecordDeleted, RecordInserted, RecordModified, ModifiedBy)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetAutoSite() As String
'        '        GetAutoSite = cmc.main_GetAutoSite()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetBanner() As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetBanner]")
'        '        'GetBanner = cmc.main_GetBanner()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetBody(HTMLDoc As String) As String
'        '        GetBody = cmc.main_GetBody(HTMLDoc)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetBoolean(InputValue As Variant) As Boolean
'        '        GetBoolean = cmc.main_GetBoolean(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetBrowserLanguageID() As Long
'        '        GetBrowserLanguageID = cmc.main_GetBrowserLanguageID()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCalendar(Optional CalendarName As Variant) As String
'        '        GetCalendar = cmc.main_GetCalendar(CalendarName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCatalog(Optional ItemContentName As Variant, Optional CategoryContentName As Variant) As String
'        '        'GetCatalog = cmc.main_GetCatalog(ItemContentName, CategoryContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCDefAdminColumns(ContentName As Variant) As Variant
'        '        GetCDefAdminColumns = cmc.main_GetCDefAdminColumns(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetChildPageList(RequestedListName As String, ContentName As String, ParentPageID As Long, AllowChildListDisplay As Boolean, Optional ArchivePages As Boolean) As String
'        '        GetChildPageList = cmc.main_GetChildPageList(RequestedListName, ContentName, ParentPageID, AllowChildListDisplay, ArchivePages)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetClosePage(Optional AllowLogin As Variant, Optional AllowTools As Variant)
'        '        GetClosePage = cmc.main_GetClosePage(AllowLogin, AllowTools)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetClosePage2(AllowLogin As Boolean, AllowTools As Boolean, BlockNonContentExtras As Boolean)
'        '        GetClosePage2 = cmc.main_GetClosePage2(AllowLogin, AllowTools, BlockNonContentExtras)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetClosePage3(AllowLogin As Boolean, AllowTools As Boolean, BlockNonContentExtras As Boolean, doNotDisposeOnExit As Boolean)
'        '        GetClosePage3 = cmc.main_GetClosePage3(AllowLogin, AllowTools, BlockNonContentExtras, doNotDisposeOnExit)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetConnectionString(DataSourcename As Variant) As String
'        '        GetConnectionString = cmc.main_GetConnectionString(DataSourcename)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContactManager(OptionString As String) As String
'        '        GetContactManager = cmc.main_GetContactManager(OptionString)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentControlCriteria(ContentName As Variant) As String
'        '        GetContentControlCriteria = cmc.main_GetContentControlCriteria(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentCopy(CopyName As Variant, Optional ContentName As Variant) As String
'        '        GetContentCopy = cmc.main_GetContentCopy(CopyName, ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentCopy2(CopyName As Variant, Optional ContentName As Variant, Optional DefaultContent As Variant) As String
'        '        GetContentCopy2 = cmc.main_GetContentCopy2(CopyName, ContentName, DefaultContent)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentCopy3(CopyName As String, Optional ContentName As String, Optional DefaultContent As String, Optional AllowEditWrapper As Boolean) As String
'        '        GetContentCopy3 = cmc.main_GetContentCopy3(CopyName, ContentName, DefaultContent, AllowEditWrapper)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentDataSource(ContentName As Variant) As String
'        '        GetContentDataSource = cmc.main_GetContentDataSource(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentFieldCount(ContentPointer As Variant) As Long
'        '        GetContentFieldCount = cmc.main_GetContentFieldCount(ContentPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentFieldLookupID(ContentPointer As Variant, FieldPointer As Variant) As String
'        '        GetContentFieldLookupID = cmc.main_GetContentFieldLookupID(ContentPointer, FieldPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentFieldMax(ContentPointer As Variant) As Long
'        '        GetContentFieldMax = cmc.main_GetContentFieldMax(ContentPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentFieldName(ContentPointer As Variant, FieldPointer As Variant)
'        '        GetContentFieldName = cmc.main_GetContentFieldName(ContentPointer, FieldPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentFieldProperty(ContentName As Variant, FieldName As Variant, PropertyName As Variant) As Variant
'        '        GetContentFieldProperty = cmc.main_GetContentFieldProperty(ContentName, FieldName, PropertyName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentFieldType(ContentPointer As Variant, FieldPointer As Variant)
'        '        GetContentFieldType = cmc.main_GetContentFieldType(ContentPointer, FieldPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentFileView(StartPath As String) As String
'        '        GetContentFileView = cmc.main_GetContentFileView(StartPath)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentFileView_GetRow(cell0, cell1, RowEven As Boolean)
'        '        GetContentFileView_GetRow = cmc.main_GetContentFileView_GetRow(cell0, cell1, RowEven)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentID(ContentName As Variant) As Long
'        '        GetContentID = cmc.main_GetContentID(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentIDByTablename(tableName As Variant) As Long
'        '        GetContentIDByTablename = cmc.main_GetContentIDByTablename(tableName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentManagementList() As String
'        '        GetContentManagementList = cmc.main_GetContentManagementList()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentNameByID(ContentID As Variant) As String
'        '        GetContentNameByID = cmc.main_GetContentNameByID(ContentID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentPage(RootPageName As Variant, Optional ContentName As Variant, Optional OrderByClause As Variant, Optional AllowChildPageList As Variant, Optional AllowReturnLink As Variant, Optional Bid As Variant) As String
'        '        GetContentPage = cmc.main_GetContentPage(RootPageName, ContentName, OrderByClause, AllowChildPageList, AllowReturnLink, Bid)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentPageArchive(RootPageName As Variant, Optional ContentName As Variant, Optional OrderByClause As Variant) As String
'        '        GetContentPageArchive = cmc.main_GetContentPageArchive(RootPageName, ContentName, OrderByClause)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentPageField(FieldName As Variant) As String
'        '        GetContentPageField = cmc.main_GetContentPageField(FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentPageMenu(RootPageName As Variant, Optional ContentName As Variant, Optional Link As Variant, Optional RootPageRecordID As Variant, Optional DepthLimit As Variant, Optional MenuStyle As Variant, Optional StyleSheetPrefix As Variant, Optional MenuImage As Variant) As String
'        '        GetContentPageMenu = cmc.main_GetContentPageMenu(RootPageName, ContentName, Link, RootPageRecordID, DepthLimit, MenuStyle, StyleSheetPrefix, MenuImage)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentPageWhatsRelated(Optional SortFieldList As Variant) As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetContentPageWhatsRelated]")
'        '        ' GetContentPageWhatsRelated = cmc.main_GetContentPageWhatsRelated(SortFieldList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentProperty(ContentName As Variant, PropertyName As Variant) As String
'        '        GetContentProperty = cmc.main_GetContentProperty(ContentName, PropertyName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentTablename(ContentName As Variant) As String
'        '        GetContentTablename = cmc.main_GetContentTablename(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentWatchLinkByID(ContentID As Variant, recordId As Variant, Optional DefaultLink As Variant, Optional IncrementClicks As Variant) As String
'        '        GetContentWatchLinkByID = cmc.main_GetContentWatchLinkByID(ContentID, recordId, DefaultLink, IncrementClicks)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentWatchLinkByKey(ContentRecordKey As Variant, Optional DefaultLink As Variant, Optional IncrementClicks As Variant) As String
'        '        GetContentWatchLinkByKey = cmc.main_GetContentWatchLinkByKey(ContentRecordKey, DefaultLink, IncrementClicks)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetContentWatchLinkByName(ContentName As Variant, recordId As Variant, Optional DefaultLink As Variant, Optional IncrementClicks As Variant) As String
'        '        GetContentWatchLinkByName = cmc.main_GetContentWatchLinkByName(ContentName, recordId, DefaultLink, IncrementClicks)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCS(CSPointer As Variant, FieldName As Variant) As String
'        '        GetCS = cmc.main_GetCS(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCS2(CSPointer As Long, FieldName As String) As String
'        '        GetCS2 = cmc.main_GetCS2(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCS2Text(CSPointer As Variant, FieldName As Variant) As String
'        '        GetCS2Text = cmc.main_GetCS2Text(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSBoolean(CSPointer As Variant, FieldName As Variant) As Boolean
'        '        GetCSBoolean = cmc.main_GetCSBoolean(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSDate(CSPointer As Variant, FieldName As Variant) As Date
'        '        GetCSDate = cmc.main_GetCSDate(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSEncodedField(CSPointer As Variant, FieldName As Variant) As String
'        '        GetCSEncodedField = cmc.main_GetCSEncodedField(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSField(CSPointer As Variant, FieldName As Variant) As Variant
'        '        GetCSField = cmc.main_GetCSField(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSFieldCaption(CSPointer As Variant, FieldName As Variant) As String
'        '        GetCSFieldCaption = cmc.main_GetCSFieldCaption(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSFieldType(CSPointer As Variant, FieldName As Variant) As Long
'        '        GetCSFieldType = cmc.main_GetCSFieldType(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSFilename(CSPointer As Variant, FieldName As Variant, OriginalFilename As Variant, Optional ContentName As Variant) As String
'        '        GetCSFilename = cmc.main_GetCSFilename(CSPointer, FieldName, OriginalFilename, ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSFirstFieldName(CSPointer As Variant) As String
'        '        GetCSFirstFieldName = cmc.main_GetCSFirstFieldName(CSPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSGroupMembers(GroupName As Variant, Optional Criteria As Variant) As Long
'        '        GetCSGroupMembers = cmc.main_GetCSGroupMembers(GroupName, Criteria)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        Public Function GetCSInteger(CSPointer As Integer, FieldName As String) As Integer
'            Return cp.Utils.EncodeInteger(GetCSText(CSPointer, FieldName))
'        End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSLookup(CSPointer As Variant, FieldName As Variant) As String
'        '        GetCSLookup = cmc.main_GetCSLookup(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSNextFieldName(CSPointer As Variant) As String
'        '        GetCSNextFieldName = cmc.main_GetCSNextFieldName(CSPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSNumber(CSPointer As Variant, FieldName As Variant) As Double
'        '        GetCSNumber = cmc.main_GetCSNumber(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSRecordAddLink(CSPointer As Variant, Optional PresetNameValueList As Variant, Optional AllowPaste As Variant) As String
'        '        GetCSRecordAddLink = cmc.main_GetCSRecordAddLink(CSPointer, PresetNameValueList, AllowPaste)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSRecordEditLink(CSPointer As Variant, Optional AllowCut As Variant) As String
'        '        GetCSRecordEditLink = cmc.main_GetCSRecordEditLink(CSPointer, AllowCut)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSRow(CSPointer As Variant) As Variant
'        '        GetCSRow = cmc.main_GetCSRow(CSPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSRowCount(CSPointer As Variant) As Long
'        '        GetCSRowCount = cmc.main_GetCSRowCount(CSPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSRows(CSPointer As Variant) As Variant
'        '        GetCSRows = cmc.main_GetCSRows(CSPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSSelectFieldList(CSPointer As Variant) As String
'        '        GetCSSelectFieldList = cmc.main_GetCSSelectFieldList(CSPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSSource(CSPointer As Variant) As String
'        '        GetCSSource = cmc.main_GetCSSource(CSPointer)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        Public Function GetCSText(CSPointer As Integer, FieldName As String) As String
'            Dim result As String = ""
'            If (CSPointer <= 100) Then
'                If (csArray(CSPointer) IsNot Nothing) Then
'                    If (csArray(CSPointer).OK()) Then
'                        result = csArray(CSPointer).GetText(FieldName)
'                    End If
'                End If
'            End If
'            Return result
'        End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetCSTextFile(CSPointer As Variant, FieldName As Variant) As String
'        '        GetCSTextFile = cmc.main_GetCSTextFile(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetDataSourceByID(DataSourceID As Long) As String
'        '        GetDataSourceByID = cmc.main_GetDataSourceByID(DataSourceID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetDataSourceType(DataSourcename As Variant) As Long
'        '        GetDataSourceType = cmc.main_GetDataSourceType(DataSourcename)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetDate(InputValue As Variant) As Date
'        '        GetDate = cmc.main_GetDate(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetDefaultValue(Key As String) As String
'        '        GetDefaultValue = cmc.main_GetDefaultValue(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetDynamicMenu(addonOptionString As String, UseContentWatchLink As Boolean) As String
'        '        GetDynamicMenu = cmc.main_GetDynamicMenu(addonOptionString, UseContentWatchLink)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetEditLockDateExpires(ContentName As Variant, recordId As Variant) As Date
'        '        GetEditLockDateExpires = cmc.main_GetEditLockDateExpires(ContentName, recordId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetEditLockMemberName(ContentName As Variant, recordId As Variant) As String
'        '        GetEditLockMemberName = cmc.main_GetEditLockMemberName(ContentName, recordId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetEditLockStatus(ContentName As Variant, recordId As Variant) As Boolean
'        '        GetEditLockStatus = cmc.main_GetEditLockStatus(ContentName, recordId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetEditWrapper(Caption As Variant, Content As Variant) As String
'        '        GetEditWrapper = cmc.main_GetEditWrapper(Caption, Content)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetErrorMessage() As String
'        '        GetErrorMessage = cmc.main_GetErrorMessage()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFeedbackForm(ContentName As Variant, recordId As Variant, ToMemberID As Variant, Optional Headline As Variant) As String
'        '        GetFeedbackForm = cmc.main_GetFeedbackForm(ContentName, recordId, ToMemberID, Headline)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFileCount(FolderPath As Variant) As Long
'        '        GetFileCount = cmc.main_GetFileCount(FolderPath)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFileList(FolderPath As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As String
'        '        GetFileList = cmc.main_GetFileList(FolderPath, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFolderList(FolderPath As Variant) As String
'        '        GetFolderList = cmc.main_GetFolderList(FolderPath)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormButton(ButtonLabel As Variant, Optional Name As Variant, Optional IDElement As Variant, Optional onClick As Variant) As String
'        '        GetFormButton = cmc.main_GetFormButton(ButtonLabel, Name, IDElement, onClick)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormButton2(ButtonLabel As Variant, Optional Name As Variant, Optional IDElement As Variant, Optional onClick As Variant, Optional Disabled As Boolean) As String
'        '        GetFormButton2 = cmc.main_GetFormButton2(ButtonLabel, Name, IDElement, onClick, Disabled)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormCSInput(CSPointer As Variant, FieldName As Variant) As String
'        '        GetFormCSInput = cmc.main_GetFormCSInput(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormEnd() As String
'        '        GetFormEnd = cmc.main_GetFormEnd()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputActiveContent(TagName As Variant, Optional DefaultValue As Variant, Optional Height As Variant, Optional Width As Variant)
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetFormInputActiveContent]")
'        '        'GetFormInputActiveContent = cmc.main_GetFormInputActiveContent(TagName, DefaultValue, Height, Width)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        Public Function GetFormInputCheckBox(TagName As String, Optional DefaultValue As Boolean = False, Optional IDElement As String = "") As String
'            Dim result As String = "<input type=checkbox name=""" & TagName & """ value=1 "
'            result += If(String.IsNullOrEmpty(IDElement), "", "id=""" & IDElement & """")
'            result += If(DefaultValue, " checked", "")
'            Return result & ">"
'        End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputCheckBox2(TagName As String, Optional DefaultValue As Boolean, Optional HtmlId As String, Optional Disabled As Boolean, Optional HtmlClass As String) As String
'        '        GetFormInputCheckBox2 = cmc.main_GetFormInputCheckBox2(TagName, DefaultValue, HtmlId, Disabled, HtmlClass)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputCheckList(TagName As String, PrimaryContentName As String, PrimaryRecordID As Long, SecondaryContentName As String, RulesContentName As String, RulesPrimaryFieldname As String, RulesSecondaryFieldName As String, Optional SecondaryContentSelectCriteria As String, Optional CaptionFieldName As Variant, Optional ReadOnly As Boolean) As String
'        '        GetFormInputCheckList = cmc.main_GetFormInputCheckList(TagName, PrimaryContentName, PrimaryRecordID, SecondaryContentName, RulesContentName, RulesPrimaryFieldname, RulesSecondaryFieldName, SecondaryContentSelectCriteria, CaptionFieldName, readOnly)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputCheckListByIDList(TagName As String, SecondaryContentName As String, CheckedIDList As String, Optional CaptionFieldName As Variant, Optional ReadOnly As Boolean) As String
'        '        GetFormInputCheckListByIDList = cmc.main_GetFormInputCheckListByIDList(TagName, SecondaryContentName, CheckedIDList, CaptionFieldName, readOnly)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputCheckListCategories(TagName As String, PrimaryContentName As String, PrimaryRecordID As Long, SecondaryContentName As String, RulesContentName As String, RulesPrimaryFieldname As String, RulesSecondaryFieldName As String, Optional SecondaryContentSelectCriteria As String, Optional CaptionFieldName As String, Optional ReadOnly As Boolean, Optional RightSideHeader As String, Optional DefaultSecondaryIDList As String) As String
'        '        GetFormInputCheckListCategories = cmc.main_GetFormInputCheckListCategories(TagName, PrimaryContentName, PrimaryRecordID, SecondaryContentName, RulesContentName, RulesPrimaryFieldname, RulesSecondaryFieldName, SecondaryContentSelectCriteria, CaptionFieldName, readOnly, RightSideHeader, DefaultSecondaryIDList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputCS(CSPointer As Variant, ContentName As Variant, FieldName As Variant, Optional Height As Variant, Optional Width As Variant, Optional IDElement As Variant)
'        '        GetFormInputCS = cmc.main_GetFormInputCS(CSPointer, ContentName, FieldName, Height, Width, IDElement)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputDate(TagName As Variant, Optional DefaultValue As Variant, Optional Width As Variant, Optional Id As Variant) As String
'        '        GetFormInputDate = cmc.main_GetFormInputDate(TagName, DefaultValue, Width, Id)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function getFormInputField(ContentName As String, FieldName As String, Optional htmlName As String, Optional HtmlValue As String, Optional HtmlClass As String, Optional HtmlId As String, Optional HtmlStyle As String, Optional ManyToManySourceRecordID As Long) As String
'        '        getFormInputField = cmc.main_GetFormInputField(ContentName, FieldName, htmlName, HtmlValue, HtmlClass, HtmlId, HtmlStyle, ManyToManySourceRecordID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputFile(TagName As Variant, Optional IDElement As Variant) As String
'        '        GetFormInputFile = cmc.main_GetFormInputFile(TagName, IDElement)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputFile2(TagName As String, Optional IDElement As String, Optional HtmlClass As String) As String
'        '        GetFormInputFile2 = cmc.main_GetFormInputFile2(TagName, IDElement, HtmlClass)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputHidden(TagName As Variant, TagValue As Variant, Optional IDElement As Variant) As String
'        '        GetFormInputHidden = cmc.main_GetFormInputHidden(TagName, TagValue, IDElement)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputHTML(TagName As Variant, Optional DefaultValue As Variant, Optional Height As Variant, Optional Width As Variant)
'        '        GetFormInputHTML = cmc.main_GetFormInputHTML(TagName, DefaultValue, Height, Width)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputHTML2(htmlName As String, Optional DefaultValue As String, Optional Height As String, Optional Width As String, Optional ReadOnly As Boolean, Optional allowActiveContent As Boolean, Optional addonListJSON As String, Optional styleList As String, Optional styleOptionList As String)
'        '        GetFormInputHTML2 = cmc.main_GetFormInputHTML2(htmlName, DefaultValue, Height, Width, readOnly, allowActiveContent, addonListJSON, styleList, styleOptionList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputHTML3(htmlName As String, Optional DefaultValue As String, Optional Height As String, Optional Width As String, Optional ReadOnly As Boolean, Optional allowActiveContent As Boolean, Optional addonListJSON As String, Optional styleList As String, Optional styleOptionList As String, Optional allowResourceLibrary As Boolean)
'        '        GetFormInputHTML3 = cmc.main_GetFormInputHTML3(htmlName, DefaultValue, Height, Width, readOnly, allowActiveContent, addonListJSON, styleList, styleOptionList, allowResourceLibrary)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputMemberSelect(MenuName As String, CurrentValue As Long, GroupID As Long, Optional ignore As Variant, Optional NoneCaption As Variant, Optional IDElement As Variant) As String
'        '        GetFormInputMemberSelect = cmc.main_GetFormInputMemberSelect(MenuName, CurrentValue, GroupID, ignore, NoneCaption, IDElement)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputMemberSelect2(MenuName As String, CurrentValue As Long, GroupID As Long, Optional ignore As Variant, Optional NoneCaption As Variant, Optional HtmlId As Variant, Optional HtmlClass As String) As String
'        '        GetFormInputMemberSelect2 = cmc.main_GetFormInputMemberSelect2(MenuName, CurrentValue, GroupID, ignore, NoneCaption, HtmlId, HtmlClass)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputRadioBox(TagName As Variant, TagValue As Variant, CurrentValue As Variant, Optional IDElement As Variant) As String
'        '        GetFormInputRadioBox = cmc.main_GetFormInputRadioBox(TagName, TagValue, CurrentValue, IDElement)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputSelect(MenuName As Variant, CurrentValue As Variant, ContentName As Variant, Optional Criteria As Variant, Optional NoneCaption As Variant, Optional IDElement As Variant) As String
'        '        GetFormInputSelect = cmc.main_GetFormInputSelect(MenuName, CurrentValue, ContentName, Criteria, NoneCaption, IDElement)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputSelect2(MenuName As String, CurrentValue As Long, ContentName As String, Criteria As String, NoneCaption As String, IDElement As String, Return_IsEmptyList As Boolean, Optional HtmlClass As String) As String
'        '        GetFormInputSelect2 = cmc.main_GetFormInputSelect2(MenuName, CurrentValue, ContentName, Criteria, NoneCaption, IDElement, Return_IsEmptyList, HtmlClass)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputSelectList(MenuName As Variant, CurrentValue As Variant, SelectList As Variant, Optional NoneCaption As Variant, Optional IDElement As Variant) As String
'        '        GetFormInputSelectList = cmc.main_GetFormInputSelectList(MenuName, CurrentValue, SelectList, NoneCaption, IDElement)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputSelectList2(MenuName As String, CurrentValue As Long, SelectList As String, NoneCaption As String, IDElement As String, Optional HtmlClass As String) As String
'        '        GetFormInputSelectList2 = cmc.main_GetFormInputSelectList2(MenuName, CurrentValue, SelectList, NoneCaption, IDElement, HtmlClass)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputStyles(TagName As String, StyleCopy As String, Optional HtmlId As String, Optional HtmlClass As String) As String
'        '        GetFormInputStyles = cmc.main_GetFormInputStyles(TagName, StyleCopy, HtmlId, HtmlClass)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputText(TagName As Variant, Optional DefaultValue As Variant, Optional Height As Variant, Optional Width As Variant, Optional Id As Variant, Optional PasswordField As Boolean) As String
'        '        GetFormInputText = cmc.main_GetFormInputText(TagName, DefaultValue, Height, Width, Id, PasswordField)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputText2(htmlName As String, Optional DefaultValue As String, Optional Height As String, Optional Width As String, Optional HtmlId As String, Optional PasswordField As Boolean, Optional Disabled As Boolean, Optional HtmlClass As String) As String
'        '        GetFormInputText2 = cmc.main_GetFormInputText2(htmlName, DefaultValue, Height, Width, HtmlId, PasswordField, Disabled, HtmlClass)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputTextExpandable(TagName As String, Optional Value As String, Optional Rows As Long, Optional styleWidth As String, Optional Id As String, Optional PasswordField As Boolean) As String
'        '        GetFormInputTextExpandable = cmc.main_GetFormInputTextExpandable(TagName, Value, Rows, styleWidth, Id, PasswordField)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputTextExpandable2(TagName As String, Optional Value As String, Optional Rows As Long, Optional styleWidth As String, Optional Id As String, Optional PasswordField As Boolean, Optional Disabled As Boolean, Optional HtmlClass As String) As String
'        '        GetFormInputTextExpandable2 = cmc.main_GetFormInputTextExpandable2(TagName, Value, Rows, styleWidth, Id, PasswordField, Disabled, HtmlClass)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputTopics(TagName As Variant, TopicContentName As Variant, ContentName As Variant, recordId As Variant) As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetFormInputTopics]")
'        '        'GetFormInputTopics = cmc.main_GetFormInputTopics(TagName, TopicContentName, ContentName, recordId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormInputWysiwyg(ByVal htmlName As String, Optional ByVal HtmlValue As String = "", Optional ByVal UserScope As EditorUserScopeEnum, Optional ByVal ContentScope As EditorContentScopeEnum, Optional ByVal Height As String = "", Optional ByVal Width As String = "", Optional ByVal HtmlClass As String = "", Optional ByVal HtmlId As String = "", Optional ReadOnly As Boolean, Optional TemplateIDForStyles As Long) As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetFormInputWysiwyg]")
'        '        'GetFormInputWysiwyg = cmc.main_GetFormInputWysiwyg(htmlName, HtmlValue, UserScope, ContentScope, Height, Width, HtmlClass, HtmlId, readOnly, TemplateIDForStyles)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormJoin() As String
'        '        GetFormJoin = cmc.main_GetFormJoin()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormLogin() As String
'        '        GetFormLogin = cmc.main_GetFormLogin()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormMyAccount(PeopleID As Variant) As String
'        '        GetFormMyAccount = cmc.main_GetFormMyAccount(PeopleID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormSendPassword() As String
'        '        GetFormSendPassword = cmc.main_GetFormSendPassword()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormSN() As String
'        '        GetFormSN = cmc.main_GetFormSN()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormStart(Optional ActionQueryString As Variant) As String
'        '        GetFormStart = cmc.main_GetFormStart(ActionQueryString)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormStart2(Optional ActionQueryString As Variant, Optional htmlName As String, Optional HtmlId As String) As String
'        '        GetFormStart2 = cmc.main_GetFormStart2(ActionQueryString, htmlName, HtmlId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetFormStart3(Optional ActionQueryString As Variant, Optional htmlName As String, Optional HtmlId As String, Optional Method As String) As String
'        '        GetFormStart3 = cmc.main_GetFormStart3(ActionQueryString, htmlName, HtmlId, Method)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetForum(ForumName As Variant, Optional ContentName As Variant) As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetForum]")
'        '        'GetForum = cmc.main_GetForum(ForumName, ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetForumList(PageTitle As Variant, Optional ContentName As Variant) As String
'        '        'GetForumList = cmc.main_GetForumList(PageTitle, ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetGroupByID(GroupID As Variant) As String
'        '        GetGroupByID = cmc.main_GetGroupByID(GroupID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetGroupID(GroupName As Variant) As Long
'        '        GetGroupID = cmc.main_GetGroupID(GroupName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHelpLink(Code As Variant, Caption As Variant, Optional BubbleCopy As Variant, Optional Link As Variant) As String
'        '        'GetHelpLink = cmc.main_GetHelpLink(Code, Caption, BubbleCopy, Link)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHTML(Optional ApplicationName As String)
'        '        GetHTML = cmc.main_GetHTML4(ApplicationName, False, False, Nothing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHTML2(Optional ApplicationName As String, Optional forceAdminPage As Boolean)
'        '        GetHTML2 = cmc.main_GetHTML4(ApplicationName, forceAdminPage, False, Nothing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHTML3(Optional ByVal ApplicationName As String, Optional ByVal forceAdminPage As Boolean, Optional ByVal doNotDisposeOnExit As Boolean)
'        '        GetHTML3 = cmc.main_GetHTML4(ApplicationName, forceAdminPage, doNotDisposeOnExit, Nothing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHTML4(Optional ByVal ApplicationName As String, Optional ByVal forceAdminPage As Boolean, Optional ByVal doNotDisposeOnExit As Boolean, Optional cpOrNothing As Object)
'        '        GetHTML4 = cmc.main_GetHTML4(ApplicationName, forceAdminPage, doNotDisposeOnExit, cpOrNothing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHTMLBody()
'        '        GetHTMLBody = cmc.main_GetHTMLBody()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHTMLDoc()
'        '        GetHTMLDoc = cmc.main_GetHTMLDoc()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHTMLDoc2(doNotDisposeOnExit As Boolean)
'        '        GetHTMLDoc2 = cmc.main_GetHTMLDoc2(doNotDisposeOnExit)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHTMLHead() As String
'        '        GetHTMLHead = cmc.main_GetHTMLHead()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetHTMLInternalHead(IsAdminSite As Boolean) As String
'        '        GetHTMLInternalHead = cmc.main_GetHTMLInternalHead(IsAdminSite)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetJoinForm() As String
'        '        GetJoinForm = cmc.main_GetJoinForm()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLandingPageID() As Long
'        '        GetLandingPageID = cmc.main_GetLandingPageID()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLastMetaDescription() As String
'        '        GetLastMetaDescription = cmc.main_GetLastMetaDescription()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLastMetaKeywordList() As String
'        '        GetLastMetaKeywordList = cmc.main_GetLastMetaKeywordList()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLastMetaTitle() As String
'        '        GetLastMetaTitle = cmc.main_GetLastMetaTitle()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLastOtherHeadTags() As String
'        '        GetLastOtherHeadTags = cmc.main_GetLastOtherHeadTags()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLinkAliasByPageID(pageId As Long, QueryStringSuffix As String, DefaultLink As String) As String
'        '        GetLinkAliasByPageID = cmc.main_GetLinkAliasByPageID(pageId, QueryStringSuffix, DefaultLink)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLinkedText(AnchorTag As Variant, AnchorText As Variant) As String
'        '        GetLinkedText = cmc.main_GetLinkedText(AnchorTag, AnchorText)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLiveTabs() As String
'        '        GetLiveTabs = cmc.main_GetLiveTabs()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLoginForm() As String
'        '        GetLoginForm = cmc.main_GetLoginForm()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLoginLink() As String
'        '        GetLoginLink = cmc.main_GetLoginLink()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLoginMemberID(loginFieldValue As String, password As String) As Long
'        '        GetLoginMemberID = cmc.main_GetLoginMemberID(loginFieldValue, password)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLoginPage() As String
'        '        GetLoginPage = cmc.main_GetLoginPage()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLoginPage2(forceDefaultLogin As Boolean) As String
'        '        GetLoginPage2 = cmc.main_GetLoginPage2(forceDefaultLogin)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetLoginPanel() As String
'        '        GetLoginPanel = cmc.main_GetLoginPanel()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMeetingSmart() As String
'        '        'GetMeetingSmart = cmc.main_GetMeetingSmart()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMemberLoginForm() As String
'        '        GetMemberLoginForm = cmc.main_GetMemberLoginForm()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMemberProfileForm() As String
'        '        GetMemberProfileForm = cmc.main_GetMemberProfileForm()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMemberProperty(PropertyName As Variant, Optional DefaultValue As Variant, Optional TargetMemberID As Variant) As String
'        '        GetMemberProperty = cmc.main_GetMemberProperty(PropertyName, DefaultValue, TargetMemberID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMemberSmart() As String
'        '        'GetMemberSmart = cmc.main_GetMemberSmart()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMenu(MenuName As Variant, MenuStyle As Variant, Optional StyleSheetPrefix As Variant) As String
'        '        GetMenu = cmc.main_GetMenu(MenuName, MenuStyle, StyleSheetPrefix)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMenuClose() As String
'        '        GetMenuClose = cmc.main_GetMenuClose()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMissing(InputValue As Variant, DefaultValue As Variant) As Variant
'        '        GetMissing = cmc.main_GetMissing(InputValue, DefaultValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMoreInfo(contactMemberID As Variant) As String
'        '        GetMoreInfo = cmc.main_GetMoreInfo(contactMemberID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMyAccountForm(PeopleID As Variant) As String
'        '        GetMyAccountForm = cmc.main_GetMyAccountForm(PeopleID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetMyProfileForm(PeopleID As Variant) As String
'        '        GetMyProfileForm = cmc.main_GetMyProfileForm(PeopleID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetNameValue(Tag As Variant, Name As Variant) As String
'        '        GetNameValue = cmc.main_GetNameValue(Tag, Name)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetNumber(InputValue As Variant) As Double
'        '        GetNumber = cmc.main_GetNumber(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function getNvaValue(Name As String, nvaEncodedString As String) As String
'        '        getNvaValue = cmc.main_GetNvaValue(Name, nvaEncodedString)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetOnLoadJavascript() As String
'        '        GetOnLoadJavascript = cmc.main_GetOnLoadJavascript()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetOrderProcess() As String
'        '        'GetOrderProcess = cmc.main_GetOrderProcess()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPageEnd() As String
'        '        GetPageEnd = cmc.main_GetPageEnd()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPageLink(pageId As Long) As String
'        '        GetPageLink = cmc.main_GetPageLink(pageId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPageLink2(pageId As Long, QueryStringSuffix As String) As String
'        '        GetPageLink2 = cmc.main_GetPageLink2(pageId, QueryStringSuffix)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPageLink3(pageId As Long, QueryStringSuffix As String, AllowLinkAlias As Boolean) As String
'        '        GetPageLink3 = cmc.main_GetPageLink3(pageId, QueryStringSuffix, AllowLinkAlias)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPageLink4(pageId As Long, QueryStringSuffix As String, AllowLinkAliasIfEnabled As Boolean, UseContentWatchNotDefaultPage As Boolean) As String
'        '        GetPageLink4 = cmc.main_GetPageLink4(pageId, QueryStringSuffix, AllowLinkAliasIfEnabled, UseContentWatchNotDefaultPage)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPageStart(Optional Title As String, Optional PageMargin As Long) As String
'        '        GetPageStart = cmc.main_GetPageStart(Title, PageMargin)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPageStartAdmin(Optional Title As String, Optional PageMargin As Long) As String
'        '        GetPageStartAdmin = cmc.main_GetPageStartAdmin(Title, PageMargin)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        'Public Function GetPanel(Panel As String, Optional StylePanel As String, Optional StyleHilite As String, Optional StyleShadow As String, Optional Width As Integer, Optional Padding As Integer, Optional HeightMin As Integer) As String
'        'End Function
'        Public Function GetPanel(Panel As String) As String
'            Dim result As String = "<div "
'            'result += If(String.IsNullOrEmpty(htmlClass), "", " class=""" & htmlClass & """")
'            'result += If(String.IsNullOrEmpty(htmlId), "", " id=""" & htmlId & """")
'            Return result & ">" & Panel & "</div>"
'        End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPanelBottom(Optional StylePanel As Variant, Optional StyleHilite As Variant, Optional StyleShadow As Variant, Optional Width As Variant, Optional Padding As Variant) As String
'        '        GetPanelBottom = cmc.main_GetPanelBottom(StylePanel, StyleHilite, StyleShadow, Width, Padding)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPanelButtons(ButtonValueList As Variant, ButtonName As Variant, Optional PanelWidth As Variant, Optional PanelHeightMin As Variant) As String
'        '        GetPanelButtons = cmc.main_GetPanelButtons(ButtonValueList, ButtonName, PanelWidth, PanelHeightMin)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPanelHeader(HeaderMessage As Variant, Optional RightSideMessage As Variant) As String
'        '        GetPanelHeader = cmc.main_GetPanelHeader(HeaderMessage, RightSideMessage)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        'Public Function GetPanelInput(PanelContent As String, Optional PanelWidth As Integer, Optional PanelHeightMin As Integer) As String
'        'End Function
'        Public Function GetPanelInput(PanelContent As String) As String
'            Dim result As String = "<div "
'            Return result & ">" & PanelContent & "</div>"
'        End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPanelRev(PanelContent As Variant, Optional PanelWidth As Variant, Optional PanelHeightMin As Variant) As String
'        '        GetPanelRev = cmc.main_GetPanelRev(PanelContent, PanelWidth, PanelHeightMin)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPanelTop(Optional StylePanel As Variant, Optional StyleHilite As Variant, Optional StyleShadow As Variant, Optional Width As Variant, Optional Padding As Variant, Optional HeightMin As Variant) As String
'        '        GetPanelTop = cmc.main_GetPanelTop(StylePanel, StyleHilite, StyleShadow, Width, Padding, HeightMin)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPCC(IsWorkflowRendering As Boolean, IsQuickEditing As Boolean) As Variant
'        '        GetPCC = cmc.main_GetPCC(IsWorkflowRendering, IsQuickEditing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPCCFirstChildPtr(pageId As Long, IsWorkflowRendering As Boolean, IsQuickEditing As Boolean) As Long
'        '        GetPCCFirstChildPtr = cmc.main_GetPCCFirstChildPtr(pageId, IsWorkflowRendering, IsQuickEditing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPCCFirstNamePtr(PageName As String, IsWorkflowRendering As Boolean, IsQuickEditing As Boolean) As Long
'        '        GetPCCFirstNamePtr = cmc.main_GetPCCFirstNamePtr(PageName, IsWorkflowRendering, IsQuickEditing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPCCPtr(pageId As Long, IsWorkflowRendering As Boolean, IsQuickEditing As Boolean) As Long
'        '        GetPCCPtr = cmc.main_GetPCCPtr(pageId, IsWorkflowRendering, IsQuickEditing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPCCPtrsSorted(PCCPtrs() As Long, OrderByCriteria As String) As Variant
'        '        GetPCCPtrsSorted = cmc.main_GetPCCPtrsSorted(PCCPtrs(), OrderByCriteria)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPleaseWaitEnd() As String
'        '        GetPleaseWaitEnd = cmc.main_GetPleaseWaitEnd()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPleaseWaitStart() As String
'        '        GetPleaseWaitStart = cmc.main_GetPleaseWaitStart()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPopupDialog(URI As Variant, Optional WindowWidth As Variant, Optional WindowHeight As Variant, Optional WindowScrollBars As Variant, Optional WindowResizable As Variant, Optional WindowName As Variant) As String
'        '        GetPopupDialog = cmc.main_GetPopupDialog(URI, WindowWidth, WindowHeight, WindowScrollBars, WindowResizable, WindowName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPopupMessage(Message As Variant, Optional WindowWidth As Variant, Optional WindowHeight As Variant, Optional WindowScrollBars As Variant, Optional WindowResizable As Variant, Optional WindowName As Variant) As String
'        '        GetPopupMessage = cmc.main_GetPopupMessage(Message, WindowWidth, WindowHeight, WindowScrollBars, WindowResizable, WindowName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetPopupPage(URI As Variant, Optional WindowWidth As Variant, Optional WindowHeight As Variant, Optional WindowScrollBars As Variant, Optional WindowResizable As Variant, Optional WindowName As Variant) As String
'        '        GetPopupPage = cmc.main_GetPopupPage(URI, WindowWidth, WindowHeight, WindowScrollBars, WindowResizable, WindowName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRandomLong() As Long
'        '        GetRandomLong = cmc.main_GetRandomLong()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRecordAddLink(ContentName As Variant, PresetNameValueList As Variant, Optional AllowPaste As Variant) As String
'        '        GetRecordAddLink = cmc.main_GetRecordAddLink(ContentName, PresetNameValueList, AllowPaste)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRecordAddLink2(ContentName As String, PresetNameValueList As String, AllowPaste As Boolean, IsEditing As Boolean) As String
'        '        GetRecordAddLink2 = cmc.main_GetRecordAddLink2(ContentName, PresetNameValueList, AllowPaste, IsEditing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRecordEditLink(ContentName As Variant, recordId As Variant, Optional AllowCut As Variant) As String
'        '        GetRecordEditLink = cmc.main_GetRecordEditLink(ContentName, recordId, AllowCut)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRecordEditLink2(ContentName As Variant, recordId As Variant, AllowCut As Boolean, RecordName As String, IsEditing As Boolean) As String
'        '        GetRecordEditLink2 = cmc.main_GetRecordEditLink2(ContentName, recordId, AllowCut, RecordName, IsEditing)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRecordEditLinkByContent(ContentID As Long, RecordIDVariant As Variant, Criteria As String) As String
'        '        GetRecordEditLinkByContent = cmc.main_GetRecordEditLinkByContent(ContentID, RecordIDVariant, Criteria)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRecordID(ContentName As Variant, RecordName As Variant) As Long
'        '        GetRecordID = cmc.main_GetRecordID(ContentName, RecordName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRecordName(ContentName As Variant, recordId As Variant) As String
'        '        GetRecordName = cmc.main_GetRecordName(ContentName, recordId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRefCount(ByRef Obj As IUnknown) As Long
'        '        'GetRefCount = cmc.main_GetRefCount(Obj)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function getRemoteQueryKey(SQL As String, Optional DataSourcename As String, Optional maxRows As Long) As String
'        '        getRemoteQueryKey = cmc.main_GetRemoteQueryKey(SQL, DataSourcename, maxRows)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetReport(RowCount As Long, ColCaption() As String, ColAlign() As String, ColWidth() As String, Cells() As String, PageSize As Long, PageNumber As Long, PreTableCopy As String, PostTableCopy As String, DataRowCount As Long, ClassStyle As String) As String
'        '        GetReport = cmc.main_GetReport(RowCount, ColCaption(), ColAlign(), ColWidth(), Cells(), PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, ClassStyle)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetResourceLibrary(Optional RootFolderName As String, Optional AllowSelectResource As Boolean, Optional SelectResourceEditorName As String) As String
'        '        GetResourceLibrary = cmc.main_GetResourceLibrary(RootFolderName, AllowSelectResource, SelectResourceEditorName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetResourceLibrary2(RootFolderName As String, AllowSelectResource As Boolean, SelectResourceEditorName As String, SelectLinkObjectName As String, AllowGroupAdd As Boolean) As String
'        '        GetResourceLibrary2 = cmc.main_GetResourceLibrary2(RootFolderName, AllowSelectResource, SelectResourceEditorName, SelectLinkObjectName, AllowGroupAdd)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetReversePanel(Panel As Variant, Optional StylePanel As Variant, Optional StyleHilite As Variant, Optional StyleShadow As Variant, Optional Width As Variant, Optional Padding As Variant, Optional HeightMin As Variant) As String
'        '        GetReversePanel = cmc.main_GetReversePanel(Panel, StylePanel, StyleHilite, StyleShadow, Width, Padding, HeightMin)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetRSField(RS As Recordset, FieldName As Variant) As Variant
'        '        GetRSField = cmc.main_GetRSField(RS, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSectionMenu(Optional DepthLimit As Variant, Optional MenuStyle As Variant, Optional StyleSheetPrefix As Variant) As String
'        '        GetSectionMenu = cmc.main_GetSectionMenu(DepthLimit, MenuStyle, StyleSheetPrefix)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSectionMenuNamed(Optional DepthLimit As Variant, Optional MenuStyle As Variant, Optional StyleSheetPrefix As Variant, Optional MenuName As Variant) As String
'        '        GetSectionMenuNamed = cmc.main_GetSectionMenuNamed(DepthLimit, MenuStyle, StyleSheetPrefix, MenuName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSectionPage(Optional AllowChildPageList As Variant, Optional AllowReturnLink As Variant) As String
'        '        GetSectionPage = cmc.main_GetSectionPage(AllowChildPageList, AllowReturnLink)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSeeAlso(ContentName As Variant, recordId As Variant) As String
'        '        GetSeeAlso = cmc.main_GetSeeAlso(ContentName, recordId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSendPasswordForm() As String
'        '        GetSendPasswordForm = cmc.main_GetSendPasswordForm()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSiteProperty(FieldName As Variant, Optional DefaultValue As Variant, Optional AllowAdminAccess As Variant) As String
'        '        GetSiteProperty = cmc.main_GetSiteProperty(FieldName, DefaultValue, AllowAdminAccess)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSiteProperty2(FieldName As String, Optional DefaultValue As String, Optional AllowAdminAccess As Boolean) As String
'        '        GetSiteProperty2 = cmc.main_GetSiteProperty2(FieldName, DefaultValue, AllowAdminAccess)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSiteTree() As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetSiteTree]")
'        '        'GetSiteTree = cmc.main_GetSiteTree()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSiteTree2() As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetSiteTree2]")
'        '        'GetSiteTree2 = cmc.main_GetSiteTree2()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSortMethodByID(SortMethodID As Long) As String
'        '        GetSortMethodByID = cmc.main_GetSortMethodByID(SortMethodID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSQLAlterColumnType(DataSourcename As String, fieldType As Long) As String
'        '        GetSQLAlterColumnType = cmc.main_GetSQLAlterColumnType(DataSourcename, fieldType)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSSC() As Variant
'        '        GetSSC = cmc.main_GetSSC()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamActiveContent(RequestName As Variant) As String
'        '        GetStreamActiveContent = cmc.main_GetStreamActiveContent(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamBoolean(RequestName As Variant) As Boolean
'        '        GetStreamBoolean = cmc.main_GetStreamBoolean(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamBoolean2(RequestName As String) As Boolean
'        '        GetStreamBoolean2 = cmc.main_GetStreamBoolean2(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamCookie(CookieName As Variant) As String
'        '        GetStreamCookie = cmc.main_GetStreamCookie(CookieName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamDate(RequestName As Variant) As Date
'        '        GetStreamDate = cmc.main_GetStreamDate(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamDate2(RequestName As String) As Date
'        '        GetStreamDate2 = cmc.main_GetStreamDate2(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamInteger(RequestName As Variant) As Long
'        '        GetStreamInteger = cmc.main_GetStreamInteger(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamInteger2(RequestName As String) As Long
'        '        GetStreamInteger2 = cmc.main_GetStreamInteger2(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamNumber(RequestName As Variant) As Double
'        '        GetStreamNumber = cmc.main_GetStreamNumber(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamNumber2(RequestName As String) As Double
'        '        GetStreamNumber2 = cmc.main_GetStreamNumber2(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamText(RequestName As Variant) As String
'        '        GetStreamText = cmc.main_GetStreamText(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStreamText2(RequestName As String) As String
'        '        GetStreamText2 = cmc.main_GetStreamText2(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStyleSheet() As String
'        '        GetStyleSheet = cmc.main_GetStyleSheet()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStyleSheet2(contentType As contentTypeEnum, Optional templateId As Long, Optional emailId As Long) As String
'        '        GetStyleSheet2 = cmc.main_GetStyleSheet2(contentType, templateId, emailId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStyleSheetDefault() As String
'        '        GetStyleSheetDefault = cmc.main_GetStyleSheetDefault()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetStyleSheetProcessed() As String
'        '        GetStyleSheetProcessed = cmc.csv_getStyleSheetProcessed()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSurveyForm(SurveyName As Variant, Optional ContentName As Variant) As String
'        '        'GetSurveyForm = cmc.main_GetSurveyForm(SurveyName, ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetSurveyList(ListTitle As Variant, Optional ContentName As Variant) As String
'        '        'GetSurveyList = cmc.main_GetSurveyList(ListTitle, ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetTableID(tableName As String) As Long
'        '        GetTableID = cmc.main_GetTableID(tableName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetTabs() As String
'        '        GetTabs = cmc.main_GetTabs()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetTemplateEditor(TagName As Variant, Optional DefaultValue As Variant, Optional Height As Variant, Optional Width As Variant)
'        '        'GetTemplateEditor = cmc.main_GetTemplateEditor(TagName, DefaultValue, Height, Width)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetTemplateEditor2(templateId As Long, TagName As Variant, Optional DefaultValue As Variant, Optional Height As Variant, Optional Width As Variant)
'        '        'GetTemplateEditor2 = cmc.main_GetTemplateEditor2(templateId, TagName, DefaultValue, Height, Width)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetTemplateLink(templateId As Long) As String
'        '        GetTemplateLink = cmc.main_GetTemplateLink(templateId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetText(InputValue As Variant) As String
'        '        GetText = cmc.main_GetText(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetTextSearch(KeywordList As Variant, Optional TopicIDList As Variant, Optional ContentIDList As Variant, Optional PageSize As Variant, Optional InitialPageNumber As Variant, Optional LanguageID As Variant) As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetTextSearch]")
'        '        'GetTextSearch = cmc.main_GetTextSearch(KeywordList, TopicIDList, ContentIDList, PageSize, InitialPageNumber, LanguageID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetTextSearchForm() As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [GetTextSearchForm]")
'        '        'GetTextSearchForm = cmc.main_GetTextSearchForm()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetTitle(Title As Variant) As String
'        '        GetTitle = cmc.main_GetTitle(Title)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetToolsForm() As String
'        '        GetToolsForm = cmc.main_GetToolsForm()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetToolsPanel() As String
'        '        GetToolsPanel = cmc.main_GetToolsPanel()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetToolsPanelLeft() As String
'        '        'GetToolsPanelLeft = cmc.main_GetToolsPanelLeft()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetUploadFormEnd() As String
'        '        GetUploadFormEnd = cmc.main_GetUploadFormEnd()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetUploadFormStart(Optional ActionQueryString As Variant) As String
'        '        GetUploadFormStart = cmc.main_GetUploadFormStart(ActionQueryString)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetUserError() As String
'        '        GetUserError = cmc.main_GetUserError()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetViewingProperty(PropertyName As String, Optional DefaultValue As String) As String
'        '        GetViewingProperty = cmc.main_GetViewingProperty(PropertyName, DefaultValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetVirtualFileCount(FolderPath As Variant) As Long
'        '        GetVirtualFileCount = cmc.main_GetVirtualFileCount(FolderPath)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetVirtualFileList(FolderPath As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As String
'        '        GetVirtualFileList = cmc.main_GetVirtualFileList(FolderPath, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetVirtualFilename(ContentName As Variant, FieldName As Variant, recordId As Variant, Optional OriginalFilename As Variant) As String
'        '        GetVirtualFilename = cmc.main_GetVirtualFilename(ContentName, FieldName, recordId, OriginalFilename)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetVirtualFolderList(FolderPath As Variant) As String
'        '        GetVirtualFolderList = cmc.main_GetVirtualFolderList(FolderPath)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetVisitorProperty(PropertyName As Variant, Optional DefaultValue As Variant, Optional TargetVisitorid As Variant) As String
'        '        GetVisitorProperty = cmc.main_GetVisitorProperty(PropertyName, DefaultValue, TargetVisitorid)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetVisitProperty(PropertyName As Variant, Optional DefaultValue As Variant, Optional TargetVisitID As Variant) As String
'        '        GetVisitProperty = cmc.main_GetVisitProperty(PropertyName, DefaultValue, TargetVisitID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetWatchList(ListName As String, SortField As String, SortReverse As Boolean) As String
'        '        GetWatchList = cmc.main_GetWatchList(ListName, SortField, SortReverse)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetWhatsNew(Optional SortFieldList As Variant) As String
'        '        GetWhatsNew = cmc.main_GetWhatsNew(SortFieldList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetWhatsRelated(ContentName As Variant, recordId As Variant, Optional SortFieldList As Variant) As String
'        '        'GetWhatsRelated = cmc.main_GetWhatsRelated(ContentName, recordId, SortFieldList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetWindowDialogJScript(URI As Variant, Optional WindowWidth As Variant, Optional WindowHeight As Variant, Optional WindowScrollBars As Variant, Optional WindowResizable As Variant, Optional WindowName As Variant) As String
'        '        GetWindowDialogJScript = cmc.main_GetWindowDialogJScript(URI, WindowWidth, WindowHeight, WindowScrollBars, WindowResizable, WindowName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetWindowOpenJScript(URI As Variant, Optional WindowWidth As Variant, Optional WindowHeight As Variant, Optional WindowScrollBars As Variant, Optional WindowResizable As Variant, Optional WindowName As Variant) As String
'        '        GetWindowOpenJScript = cmc.main_GetWindowOpenJScript(URI, WindowWidth, WindowHeight, WindowScrollBars, WindowResizable, WindowName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetWizardContent(HeaderCaption As String, ButtonListLeft As String, ButtonListRight As String, AllowAdd As Boolean, AllowDelete As Boolean, Description As String, Content As String) As String
'        '        GetWizardContent = cmc.main_GetWizardContent(HeaderCaption, ButtonListLeft, ButtonListRight, AllowAdd, AllowDelete, Description, Content)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetWordSearchExcludeList() As String
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [main_GetWordSearchExcludeList]")
'        '        'GetWordSearchExcludeList = cmc.main_GetWordSearchExcludeList()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function GetYesNo(InputValue As Variant) As String
'        '        GetYesNo = cmc.main_GetYesNo(InputValue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function Get_old_MyProfileForm(PeopleID As Variant) As String
'        '        Get_old_MyProfileForm = cmc.main_Get_old_MyProfileForm(PeopleID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IISReset() As Boolean
'        '        IISReset = cmc.main_IISReset()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ImportCollection(CollectionFileData As String) As Boolean
'        '        ImportCollection = cmc.main_ImportCollection(CollectionFileData)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ImportCollectionFile(CollectionFilename As String) As Boolean
'        '        ImportCollectionFile = cmc.main_ImportCollectionFile(CollectionFilename)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function InsertContentRecord(ContentName As Variant) As Long
'        '        InsertContentRecord = cmc.main_InsertContentRecord(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function InsertContentRecordGetID(ContentName As Variant) As Long
'        '        InsertContentRecordGetID = cmc.main_InsertContentRecordGetID(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function InsertContentRecordGetID_Fast(ContentName As String) As Long
'        '        InsertContentRecordGetID_Fast = cmc.main_InsertContentRecordGetID_Fast(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function InsertCSContent(ContentName As Variant) As Long
'        '        InsertCSContent = cmc.main_InsertCSContent(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function InsertCSRecord(ContentName As Variant) As Long
'        '        InsertCSRecord = cmc.main_InsertCSRecord(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function InStream(RequestName As String) As Boolean
'        '        InStream = cmc.main_InStream(RequestName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsAdmin() As Boolean
'        '        IsAdmin = cmc.main_IsAdmin()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsAdvancedEditing(ContentName As Variant) As Boolean
'        '        IsAdvancedEditing = cmc.main_IsAdvancedEditing(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsAuthenticated() As Boolean
'        '        IsAuthenticated = cmc.main_IsAuthenticated()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsAuthoring(ContentName As Variant) As Boolean
'        '        IsAuthoring = cmc.main_IsAuthoring(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsContentFieldSupported(ContentName As String, FieldName As String) As Boolean
'        '        IsContentFieldSupported = cmc.main_IsContentFieldSupported(ContentName, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsContentManager(Optional ContentName As Variant) As Boolean
'        '        IsContentManager = cmc.main_IsContentManager2(kmaEncodeText(ContentName))
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsContentManager2(Optional ContentName As String) As Boolean
'        '        IsContentManager2 = cmc.main_IsContentManager2(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsCSFieldSupported(CSPointer As Variant, FieldName As Variant) As Boolean
'        '        IsCSFieldSupported = cmc.main_IsCSFieldSupported(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        Public Function IsCSOK(CSPointer As Integer) As Boolean
'            Dim result As Boolean = False
'            If (CSPointer <= 100) Then
'                If (csArray(CSPointer) IsNot Nothing) Then
'                    result = csArray(CSPointer).OK()
'                End If
'            End If
'            Return result
'        End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsDeveloper() As Boolean
'        '        IsDeveloper = cmc.main_IsDeveloper()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsEditing(ContentName As Variant) As Boolean
'        '        IsEditing = cmc.main_IsEditing(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsEditingAnything() As Boolean
'        '        IsEditingAnything = cmc.main_IsEditingAnything()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsGroupListMember(GroupIDList As String, Optional CheckMemberID As Variant) As Boolean
'        '        IsGroupListMember = cmc.main_IsGroupListMember(GroupIDList, CheckMemberID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsGroupListMember2(GroupIDList As String, Optional CheckMemberID As Long, Optional adminReturnsTrue As Boolean) As Boolean
'        '        IsGroupListMember2 = cmc.main_IsGroupListMember2(GroupIDList, CheckMemberID, adminReturnsTrue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsGroupMember(GroupName As Variant, Optional CheckMemberID As Variant) As Boolean
'        '        IsGroupMember = cmc.main_IsGroupMember(GroupName, CheckMemberID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsGroupMember2(GroupName As Variant, Optional CheckMemberID As Long, Optional adminReturnsTrue As Boolean) As Boolean
'        '        IsGroupMember2 = cmc.main_IsGroupMember2(GroupName, CheckMemberID, adminReturnsTrue)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsGuest() As Boolean
'        '        IsGuest = cmc.main_IsGuest()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsLinkAuthoring(ContentName As Variant) As Boolean
'        '        IsLinkAuthoring = cmc.main_IsLinkAuthoring(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsLoginOK(username As String, password As String, Optional ErrorMessage As Variant, Optional ErrorCode As Variant) As Boolean
'        '        IsLoginOK = cmc.main_IsLoginOK(username, password, ErrorMessage, ErrorCode)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsMember() As Boolean
'        '        IsMember = cmc.main_IsMember()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsNewLoginOK(username As String, password As String, Optional ErrorMessage As Variant, Optional ErrorCode As Variant) As Boolean
'        '        IsNewLoginOK = cmc.main_IsNewLoginOK(username, password, ErrorMessage, ErrorCode)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsOrderOK() As Boolean
'        '        'IsOrderOK = cmc.main_IsOrderOK()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsPresentationAuthoring(ContentName As Variant) As Boolean
'        '        IsPresentationAuthoring = cmc.main_IsPresentationAuthoring(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsPromotionExpired(PromotionKey As Variant) As Boolean
'        '        IsPromotionExpired = cmc.main_IsPromotionExpired(PromotionKey)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsPromotionUsed(PromotionKey As Variant) As Boolean
'        '        IsPromotionUsed = cmc.main_IsPromotionUsed(PromotionKey)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsQuickEditing(ContentName As Variant) As Boolean
'        '        IsQuickEditing = cmc.main_IsQuickEditing(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsRecognized() As Boolean
'        '        IsRecognized = cmc.main_IsRecognized()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsRecordLocked(ContentName As Variant, recordId As Variant) As Boolean
'        '        IsRecordLocked = cmc.main_IsRecordLocked(ContentName, recordId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function isSectionBlocked(SectionID As Long, AllowSectionBlocking As Boolean) As Boolean
'        '        isSectionBlocked = cmc.main_isSectionBlocked(SectionID, AllowSectionBlocking)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsSQLTable(DataSourcename As String, tableName As String) As Boolean
'        '        IsSQLTable = cmc.main_IsSQLTable(DataSourcename, tableName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsSQLTableField(DataSourcename As String, tableName As String, FieldName As String) As Boolean
'        '        IsSQLTableField = cmc.main_IsSQLTableField(DataSourcename, tableName, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsUserError() As Boolean
'        '        IsUserError = cmc.main_IsUserError()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsViewingProperty(PropertyName As String) As Boolean
'        '        IsViewingProperty = cmc.main_IsViewingProperty(PropertyName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsWithinContent(ChildContentID As Variant, ParentContentID As Variant) As Boolean
'        '        IsWithinContent = cmc.main_IsWithinContent(ChildContentID, ParentContentID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsWorkflowAuthoringCompatible(ContentName As String) As Boolean
'        '        IsWorkflowAuthoringCompatible = cmc.main_IsWorkflowAuthoringCompatible(ContentName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function IsWorkflowRendering() As Boolean
'        '        IsWorkflowRendering = cmc.main_IsWorkflowRendering()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function LoginMember(loginFieldValue As Variant, password As Variant, Optional AllowAutoLogin As Variant) As Boolean
'        '        LoginMember = cmc.main_LoginMember(loginFieldValue, password, AllowAutoLogin)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function LoginMemberByID(recordId As Variant, Optional AllowAutoLogin As Variant) As Boolean
'        '        LoginMemberByID = cmc.main_LoginMemberByID(recordId, AllowAutoLogin)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenContent(ContentName As Variant, Optional Criteria As Variant, Optional SortFieldList As Variant, Optional ActiveOnly As Variant) As Long
'        '        OpenContent = cmc.main_OpenContent(ContentName, Criteria, SortFieldList, ActiveOnly)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        Public Function OpenCSContent(ContentName As String, Optional Criteria As String = "", Optional SortFieldList As String = "", Optional ActiveOnly As Boolean = False, Optional ignore0 As Boolean = False, Optional ignore1 As Boolean = False, Optional SelectFieldList As String = "", Optional PageSize As Integer = 9999, Optional PageNumber As Integer = 1) As Integer
'            Dim result As Integer = 0
'            Do While (csArray(result) IsNot Nothing)
'                result += 1
'            Loop
'            Dim cs As CPCSBaseClass = cp.CSNew()
'            If (cs.Open(ContentName, Criteria, SortFieldList, ActiveOnly, SelectFieldList, PageSize, PageNumber)) Then
'                csArray(result) = cs
'            Else
'                result = -1
'            End If
'            Return result
'        End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSContentRecord(ContentName As Variant, recordId As Variant, Optional WorkflowAuthoringMode As Variant, Optional WorkflowEditingMode As Variant, Optional SelectFieldList As Variant) As Long
'        '        OpenCSContentRecord = cmc.main_OpenCSContentRecord(ContentName, recordId, WorkflowAuthoringMode, WorkflowEditingMode, SelectFieldList)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSContentWatchList(ListName As Variant, Optional SortFieldList As Variant, Optional ActiveOnly As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
'        '        OpenCSContentWatchList = cmc.main_OpenCSContentWatchList(ListName, SortFieldList, ActiveOnly, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSContent_Internal(ContentName As String, Optional Criteria As Variant, Optional SortFieldList As Variant, Optional ActiveOnly As Variant, Optional WorkflowRenderingMode As Variant, Optional WorkflowEditingMode As Variant, Optional SelectFieldList As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
'        '        OpenCSContent_Internal = cmc.main_OpenCSContent_Internal(ContentName, Criteria, SortFieldList, ActiveOnly, WorkflowRenderingMode, WorkflowEditingMode, SelectFieldList, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSGroupList(GroupNameList As Variant, Optional Criteria As Variant, Optional SortFieldList As Variant, Optional ActiveOnly As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
'        '        OpenCSGroupList = cmc.main_OpenCSGroupList(GroupNameList, Criteria, SortFieldList, ActiveOnly, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSGroupMembers(GroupName As Variant, Optional Criteria As Variant, Optional SortFieldList As Variant, Optional ActiveOnly As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
'        '        OpenCSGroupMembers = cmc.main_OpenCSGroupMembers(GroupName, Criteria, SortFieldList, ActiveOnly, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSJoin(CSPointer As Variant, FieldName As Variant) As Long
'        '        OpenCSJoin = cmc.main_OpenCSJoin(CSPointer, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSSQL(DataSourcename As Variant, SQL As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
'        '        OpenCSSQL = cmc.main_OpenCSSQL(DataSourcename, SQL, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSTextSearch(KeywordList As Variant, Optional TopicIDList As Variant, Optional ContentIDList As Variant, Optional PageSize As Variant, Optional PageNumber As Variant, Optional LanguageID As Variant) As Long
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [OpenCSTextSearch]")
'        '        'OpenCSTextSearch = cmc.main_OpenCSTextSearch(KeywordList, TopicIDList, ContentIDList, PageSize, PageNumber, LanguageID)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSWhatsNew(Optional SortFieldList As Variant, Optional ActiveOnly As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
'        '        OpenCSWhatsNew = cmc.main_OpenCSWhatsNew(SortFieldList, ActiveOnly, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenCSWhatsRelated(ContentName As Variant, recordId As Variant, Optional SortFieldList As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
'        '        'OpenCSWhatsRelated = cmc.main_OpenCSWhatsRelated(ContentName, recordId, SortFieldList, PageSize, PageNumber)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenRSSQL(DataSourcename As Variant, SQL As Variant, Optional CommandTimeout, Optional PageSize, Optional PageNumber, Optional UseCompatibleCursor As Variant, Optional UseServerCursor As Variant) As Recordset
'        '    Set OpenRSSQL = cmc.main_OpenRSSQL(DataSourcename, SQL, CommandTimeout, PageSize, PageNumber, UseCompatibleCursor, UseServerCursor)
'        'End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenRSTable(DataSourcename As Variant, tableName As Variant, Criteria As Variant, SortFieldList As Variant, Optional SelectFieldList As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Recordset
'        '    Set OpenRSTable = cmc.main_OpenRSTable(DataSourcename, tableName, Criteria, SortFieldList, SelectFieldList, PageSize, PageNumber)
'        'End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenSQL(DataSourcename As Variant, SQL As Variant)
'        '        OpenSQL = cmc.main_OpenSQL(DataSourcename, SQL)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OpenTable(tableName As Variant, Optional Criteria As Variant, Optional SortFieldList As Variant, Optional ActiveOnly As Variant) As Recordset
'        '    Set OpenTable = cmc.main_OpenTable(tableName, Criteria, SortFieldList, ActiveOnly)
'        'End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function OrderOK() As Boolean
'        '        'OrderOK = cmc.main_OrderOK()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function parseJSON(Source As String) As Object
'        '    Set parseJSON = cmc.main_parseJSON(Source)
'        'End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ProcessFormInputFile(TagName As Variant, Optional VirtualFilePath As Variant) As String
'        '        ProcessFormInputFile = cmc.main_ProcessFormInputFile(TagName, VirtualFilePath)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ProcessReplacement(NameValueLines As Variant, Source As Variant) As String
'        '        ProcessReplacement = cmc.main_ProcessReplacement(NameValueLines, Source)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadBake(Name As Variant) As String
'        '        ReadBake = cmc.main_ReadBake(Name)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadCache(Name As Variant) As String
'        '        ReadCache = cmc.main_ReadCache(Name)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadFile(Filename As Variant) As String
'        '        ReadFile = cmc.main_ReadFile(Filename)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadStreamBoolean(Key As Variant) As Boolean
'        '        ReadStreamBoolean = cmc.main_ReadStreamBoolean(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadStreamDate(Key As Variant) As Variant
'        '        ReadStreamDate = cmc.main_ReadStreamDate(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadStreamEncodeBoolean(Key As Variant) As Boolean
'        '        ReadStreamEncodeBoolean = cmc.main_ReadStreamEncodeBoolean(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadStreamEncodeDate(Key As Variant) As Date
'        '        ReadStreamEncodeDate = cmc.main_ReadStreamEncodeDate(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadStreamEncodeInteger(Key As Variant) As Long
'        '        ReadStreamEncodeInteger = cmc.main_ReadStreamEncodeInteger(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadStreamEncodeNumber(Key As Variant) As Double
'        '        ReadStreamEncodeNumber = cmc.main_ReadStreamEncodeNumber(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadStreamEncodeText(Key As Variant) As String
'        '        ReadStreamEncodeText = cmc.main_ReadStreamEncodeText(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadStreamNumber(Key As Variant) As Variant
'        '        ReadStreamNumber = cmc.main_ReadStreamNumber(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadStreamText(Key As Variant) As Variant
'        '        ReadStreamText = cmc.main_ReadStreamText(Key)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function ReadVirtualFile(Filename As Variant) As String
'        '        ReadVirtualFile = cmc.main_ReadVirtualFile(Filename)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function RecognizeMemberByID(recordId As Long) As Boolean
'        '        RecognizeMemberByID = cmc.main_RecognizeMemberByID(recordId)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function RedirectByLink(Link As Variant)
'        '        RedirectByLink = cmc.main_RedirectByLink(Link)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function RedirectByRecord_ReturnStatus(ContentName As Variant, recordId As Variant, Optional FieldName As Variant) As Boolean
'        '        RedirectByRecord_ReturnStatus = cmc.main_RedirectByRecord_ReturnStatus(ContentName, recordId, FieldName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function RemoveControlCharacters(DirtyText As Variant) As String
'        '        RemoveControlCharacters = cmc.main_RemoveControlCharacters(DirtyText)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function RunSQL(DataSourcename As Variant, SQL As Variant, Optional Retries As Variant) As Recordset
'        '    Set RunSQL = cmc.main_RunSQL(DataSourcename, SQL, Retries)
'        'End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function SaveCache(Name As Variant, ValueText As Variant) As String
'        '        SaveCache = cmc.main_SaveCache(Name, ValueText)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function SaveStreamFile(TagName As Variant, Optional VirtualFilePath As Variant) As String
'        '        SaveStreamFile = cmc.main_SaveStreamFile(TagName, VirtualFilePath)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function SendEmail(toAddress As Variant, fromAddress As Variant, SubjectMessage As Variant, BodyMessage As Variant, Optional optionalEmailIdForLog As Variant, Optional Immediate As Variant, Optional HTML As Variant) As String
'        '        SendEmail = cmc.main_SendEmail(toAddress, fromAddress, SubjectMessage, BodyMessage, optionalEmailIdForLog, Immediate, HTML)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function SendMemberEmail(ToMemberID As Variant, From As Variant, subject As Variant, Body As Variant, Immediate As Variant, HTML As Variant) As String
'        '        SendMemberEmail = cmc.main_SendMemberEmail(ToMemberID, From, subject, Body, Immediate, HTML)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function SendMemberEmail2(ToMemberID As Variant, From As Variant, subject As Variant, Body As Variant, Immediate As Variant, HTML As Variant, Optional emailIdForLog As Long) As String
'        '        SendMemberEmail2 = cmc.main_SendMemberEmail2(ToMemberID, From, subject, Body, Immediate, HTML, emailIdForLog)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function SendMemberEmail_Fast(ToMemberID As Long, From As String, subject As String, Body As String, Immediate As Boolean, HTML As Boolean) As String
'        '        SendMemberEmail_Fast = cmc.main_SendMemberEmail_Fast(ToMemberID, From, subject, Body, Immediate, HTML)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function SendMemberPassword(email As Variant) As Boolean
'        '        SendMemberPassword = cmc.main_SendMemberPassword(email)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function SetMemberIdentity(Criteria As Variant) As Boolean
'        '        SetMemberIdentity = cmc.main_SetMemberIdentity(Criteria)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function Testing() As Boolean
'        '        Testing = cmc.main_Testing()
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function VerifyDynamicMenu(MenuName As Variant) As String
'        '        VerifyDynamicMenu = cmc.main_VerifyDynamicMenu(MenuName)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Function WrapContent(ByVal Content As String, WrapperID As Long, Optional WrapperSourceForComment As String) As String
'        '        WrapContent = cmc.main_WrapContent(Content, WrapperID, WrapperSourceForComment)
'        '    End Function
'        '    '
'        '    '
'        '    '
'        '    Public Sub AbortEdit(ContentName As Variant, recordId As Variant)
'        '        Call cmc.main_AbortEdit(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddEndOfBodyJavascript(NewCode As String)
'        '        Call cmc.main_AddEndOfBodyJavascript(NewCode)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddEndOfBodyJavascript2(NewCode As String, addedByMessage As String)
'        '        Call cmc.main_AddEndOfBodyJavascript2(NewCode, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddErrorMessage(Message As Variant)
'        '        Call cmc.main_AddErrorMessage(Message)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddGroupMember(GroupName As Variant, Optional NewMemberID As Variant, Optional DateExpires As Variant)
'        '        Call cmc.main_AddGroupMember(GroupName, NewMemberID, DateExpires)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddHeadJavascript(NewCode As String)
'        '        Call cmc.main_AddHeadJavascript(NewCode)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddHeadScriptCode(NewCode As String, addedByMessage As String)
'        '        Call cmc.main_AddHeadScriptCode(NewCode, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddHeadScriptLink(Filename As String, addedByMessage As String)
'        '        Call cmc.main_AddHeadScriptLink(Filename, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddHeadTag(HeadTag As String)
'        '        Call cmc.main_AddHeadTag(HeadTag)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddHeadTag2(HeadTag As String, addedByMessage As String)
'        '        Call cmc.main_AddHeadTag2(HeadTag, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddInStream(QS As String)
'        '        Call cmc.main_AddInStream(QS)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddLinkAlias(linkAlias As String, pageId As Long, QueryStringSuffix As String, Optional OverRideDuplicate As Boolean, Optional DupCausesWarning As Boolean)
'        '        Call cmc.main_AddLinkAlias(linkAlias, pageId, QueryStringSuffix, OverRideDuplicate, DupCausesWarning)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddLiveTabEntry(Caption As Variant, LiveBody As Variant, Optional StylePrefix As Variant)
'        '        Call cmc.main_AddLiveTabEntry(Caption, LiveBody, StylePrefix)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddMenuEntry(Name As Variant, Optional ParentName As Variant, Optional ImageLink As Variant, Optional ImageOverLink As Variant, Optional Link As Variant, Optional Caption As Variant, Optional styleSheet As Variant, Optional StyleSheetHover As Variant, Optional NewWindow As Boolean)
'        '        Call cmc.main_AddMenuEntry(Name, ParentName, ImageLink, ImageOverLink, Link, Caption, styleSheet, StyleSheetHover, NewWindow)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub addMeta(metaName As String, metaContent As String, addedByMessage As String)
'        '        Call cmc.main_addMeta(metaName, metaContent, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddMetaDescription(MetaDescription As String)
'        '        Call cmc.main_addMetaDescription(MetaDescription)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddMetaDescription2(MetaDescription As String, addedByMessage As String)
'        '        Call cmc.main_addMetaDescription2(MetaDescription, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddMetaKeywordList(MetaKeywordList As String)
'        '        Call cmc.main_addMetaKeywordList(MetaKeywordList)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddMetaKeywordList2(MetaKeywordList As String, addedByMessage As String)
'        '        Call cmc.main_addMetaKeywordList2(MetaKeywordList, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub addMetaProperty(metaProperty As String, metaContent As String, addedByMessage As String)
'        '        Call cmc.main_addMetaProperty(metaProperty, metaContent, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddOnLoadJavascript(NewCode As String)
'        '        Call cmc.main_AddOnLoadJavascript(NewCode)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddOnLoadJavascript2(NewCode As String, addedByMessage As String)
'        '        Call cmc.main_AddOnLoadJavascript2(NewCode, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddPageTitle(PageTitle As String)
'        '        Call cmc.main_AddPagetitle(PageTitle)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddPageTitle2(PageTitle As String, addedByMessage As String)
'        '        Call cmc.main_AddPagetitle2(PageTitle, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddSharedStyleID(styleid As Long)
'        '        Call cmc.main_AddSharedStyleID(styleid)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddSharedStyleID2(styleid As Long, Optional addedByMessage As String)
'        '        Call cmc.main_AddSharedStyleID2(styleid, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddStylesheetLink(StyleSheetLink As String)
'        '        Call cmc.main_AddStylesheetLink(StyleSheetLink)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddStylesheetLink2(StyleSheetLink As String, addedByMessage As String)
'        '        Call cmc.main_AddStylesheetLink2(StyleSheetLink, addedByMessage)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddTabEntry(Caption As Variant, Link As Variant, IsHit As Variant, Optional StylePrefix As Variant, Optional LiveBody As Variant)
'        '        Call cmc.main_AddTabEntry(Caption, Link, IsHit, StylePrefix, LiveBody)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AddUserError(Message As Variant)
'        '        Call cmc.main_AddUserError(Message)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub AppendVirtualFile(Filename As Variant, fileContent As Variant)
'        '        Call cmc.main_AppendVirtualFile(Filename, fileContent)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ApproveEdit(ContentName As Variant, recordId As Variant)
'        '        Call cmc.main_ApproveEdit(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CacheSQL(DataSourcename As Variant, SQL As Variant, Optional Retries As Variant)
'        '        Call cmc.main_CacheSQL(DataSourcename, SQL, Retries)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CheckMember()
'        '        Call cmc.main_CheckMember
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ClearBake(ContentNameList As Variant)
'        '        Call cmc.main_ClearBake(ContentNameList)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub clearCache()
'        '        Call cmc.main_clearCache
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ClearEditLock(ContentName As Variant, recordId As Variant)
'        '        Call cmc.main_ClearEditLock(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ClearPageContentCache()
'        '        Call cmc.main_ClearPageContentCache
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ClearPageTemplateCache()
'        '        Call cmc.main_ClearPagetemplateCache
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ClearSiteSectionCache()
'        '        Call cmc.main_ClearSiteSectionCache
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ClearStream()
'        '        Call cmc.main_ClearStream
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        Public Sub CloseCS(CSPointer As Integer, Optional AsyncSave As Boolean = False)
'            If (CSPointer <= 100) Then
'                If (csArray(CSPointer) IsNot Nothing) Then
'                    If (csArray(CSPointer).OK()) Then
'                        csArray(CSPointer).Close()
'                    End If
'                End If
'            End If
'        End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ClosePage(Optional AllowLogin As Variant, Optional AllowTools As Variant)
'        '        'Call cmc.main_ClosePage(AllowLogin, AllowTools)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CloseStream()
'        '        Call cmc.main_CloseStream
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ContentWatch(ContentName As Variant, recordId As Variant)
'        '        Call cmc.main_ContentWatch(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CopyCSRecord(CSSource As Variant, CSDestination As Variant)
'        '        Call cmc.main_CopyCSRecord(CSSource, CSDestination)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub copyFile(sourcePathFilename As Variant, destinationPathFilename As Variant)
'        '        Call cmc.main_copyFile(sourcePathFilename, destinationPathFilename)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CopyVirtualFile(SourceFilename As Variant, destinationFilename As Variant)
'        '        Call cmc.main_CopyVirtualFile(SourceFilename, destinationFilename)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateAdminMenu(ParentName As Variant, EntryName As Variant, ContentName As Variant, LinkPage As Variant, SortOrder As Variant, Optional AdminOnly As Variant, Optional DeveloperOnly As Variant, Optional NewWindow As Variant)
'        '        Call cmc.main_CreateAdminMenu(ParentName, EntryName, ContentName, LinkPage, SortOrder, AdminOnly, DeveloperOnly, NewWindow)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateContent(DataSourcename As Variant, tableName As Variant, ContentName As Variant)
'        '        Call cmc.main_CreateContent(DataSourcename, tableName, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateContentChild(ChildContentName As Variant, ParentContentName As Variant)
'        '        Call cmc.main_CreateContentChild(ChildContentName, ParentContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateContentField(ContentName As String, FieldName As String, fieldType As Long, Optional FieldSortOrder As Variant, Optional FieldAuthorable As Variant, Optional FieldCaption As Variant, Optional LookupContentName As Variant, Optional DefaultValue As Variant, Optional NotEditable As Variant, Optional AdminIndexColumn As Variant, Optional AdminIndexWidth As Variant, Optional AdminIndexSort As Variant, Optional RedirectContentName As Variant, Optional RedirectIDField As Variant, Optional RedirectPath As Variant, Optional HTMLContent As Variant, Optional UniqueName As Variant, Optional password As Variant, Optional AdminOnly As Boolean, Optional DeveloperOnly As Boolean, Optional ReadOnly As Boolean, Optional FieldRequired As Boolean)
'        '        Call cmc.main_CreateContentField(ContentName, FieldName, fieldType, FieldSortOrder, FieldAuthorable, FieldCaption, LookupContentName, DefaultValue, NotEditable, AdminIndexColumn, AdminIndexWidth, AdminIndexSort, RedirectContentName, RedirectIDField, RedirectPath, HTMLContent, UniqueName, password, AdminOnly, DeveloperOnly, readOnly, FieldRequired)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateContentFieldsFromSQLTable(DataSourcename As Variant, tableName As Variant)
'        '        Call cmc.main_CreateContentFieldsFromSQLTable(DataSourcename, tableName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateContentFromSQLTable(DataSourcename As Variant, tableName As Variant, ContentName As Variant)
'        '        Call cmc.main_CreateContentFromSQLTable(DataSourcename, tableName, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateFileFolder(FolderPath As Variant)
'        '        Call cmc.main_CreateFileFolder(FolderPath)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreatePeople()
'        '        Call cmc.main_CreatePeople
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateSQLIndex(DataSourcename As Variant, tableName As Variant, IndexName As Variant, FieldNames As Variant)
'        '        Call cmc.main_CreateSQLIndex(DataSourcename, tableName, IndexName, FieldNames)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateSQLTable(DataSourcename As Variant, tableName As Variant, Optional AllowAutoIncrement As Variant)
'        '        Call cmc.main_CreateSQLTable(DataSourcename, tableName, AllowAutoIncrement)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateSQLTableField(DataSourcename As String, tableName As String, FieldName As String, fieldType As Long)
'        '        Call cmc.main_CreateSQLTableField(DataSourcename, tableName, FieldName, fieldType)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub CreateUser()
'        '        Call cmc.main_CreateUser
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteContentRecord(ContentName As Variant, recordId As Variant)
'        '        Call cmc.main_DeleteContentRecord(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteContentRecords(ContentName As Variant, Criteria As Variant)
'        '        Call cmc.main_DeleteContentRecords(ContentName, Criteria)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteContentRecordsByPointer(ContentPointer As Variant, Criteria As Variant)
'        '        Call cmc.main_DeleteContentRecordsByPointer(ContentPointer, Criteria)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteContentTracking(ContentName As Variant, recordId As Variant, Permanent As Variant)
'        '        Call cmc.main_DeleteContentTracking(ContentName, recordId, Permanent)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteCSRecord(CSPointer As Variant)
'        '        Call cmc.main_DeleteCSRecord(CSPointer)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteFile(Filename As Variant)
'        '        Call cmc.main_DeleteFile(Filename)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteGroup(GroupName As Variant)
'        '        Call cmc.main_DeleteGroup(GroupName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteGroupMember(GroupName As Variant, Optional NewMemberID As Variant)
'        '        Call cmc.main_DeleteGroupMember(GroupName, NewMemberID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteSQLIndex(DataSourcename As Variant, tableName As Variant, IndexName As Variant)
'        '        Call cmc.main_DeleteSQLIndex(DataSourcename, tableName, IndexName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteTableRecord(DataSourcename As String, tableName As String, recordId As Long)
'        '        Call cmc.main_DeleteTableRecord(DataSourcename, tableName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub DeleteVirtualFile(Filename As Variant)
'        '        Call cmc.main_DeleteVirtualFile(Filename)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub dispose()
'        '        'Call cmc.dispose
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub EncodePage_SplitBody(PageSource As String, PageSourceBody As String, PageSourcePreBody As String, PageSourcePostBody As String)
'        '        Call cmc.main_EncodePage_SplitBody(PageSource, PageSourceBody, PageSourcePreBody, PageSourcePostBody)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ExportXML(Filename As Variant)
'        '        Call cmc.main_ExportXML(Filename)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub FirstCSRecord(CSPointer As Variant, Optional AsyncSave As Variant)
'        '        Call cmc.main_FirstCSRecord(CSPointer, AsyncSave)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub FlushStream()
'        '        Call cmc.main_FlushStream
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub GetAuthoringPermissions(ContentName As String, recordId As Long, AllowInsert As Boolean, AllowCancel As Boolean, allowSave As Boolean, AllowDelete As Boolean, AllowPublish As Boolean, AllowAbort As Boolean, AllowSubmit As Boolean, AllowApprove As Boolean, ReadOnly As Boolean)
'        '        Call cmc.main_GetAuthoringPermissions(ContentName, recordId, AllowInsert, AllowCancel, allowSave, AllowDelete, AllowPublish, AllowAbort, AllowSubmit, AllowApprove, readOnly)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub GetAuthoringStatus(ContentName As String, recordId As Long, IsSubmitted As Boolean, IsApproved As Boolean, SubmittedName As String, ApprovedName As String, IsInserted As Boolean, IsDeleted As Boolean, IsModified As Boolean, ModifiedName As String, ModifiedDate As Date, SubmittedDate As Date, ApprovedDate As Date)
'        '        Call cmc.main_GetAuthoringStatus(ContentName, recordId, IsSubmitted, IsApproved, SubmittedName, ApprovedName, IsInserted, IsDeleted, IsModified, ModifiedName, ModifiedDate, SubmittedDate, ApprovedDate)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub GetBrowserLanguage(LanguageID As Long, LanguageName As String)
'        '        Call cmc.main_GetBrowserLanguage(LanguageID, LanguageName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub HandleError_TrapPatch(MethodName As String, Optional ErrorCause As Variant)
'        '        Call cmc.main_HandleError_TrapPatch(MethodName, ErrorCause)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub HandleWCCSError(MethodName As String)
'        '        Call cmc.main_HandleWCCSError(MethodName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub HandleWCInternalError(MethodName As String, cause As String)
'        '        Call cmc.main_HandleWCInternalError(MethodName, cause)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub HandleWCTrapError(MethodName As String)
'        '        Call cmc.main_HandleWCTrapError(MethodName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ImportXML(Filename As Variant)
'        '        Call cmc.main_ImportXML(Filename)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub IncrementContentField(ContentName As Variant, recordId As Variant, FieldName As Variant)
'        '        Call cmc.main_IncrementContentField(ContentName, recordId, FieldName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub IncrementErrorCount()
'        '        Call cmc.main_IncrementErrorCount
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub IncrementTableField(tableName As Variant, recordId As Variant, FieldName As Variant, Optional DataSourcename As Variant)
'        '        Call cmc.main_IncrementTableField(tableName, recordId, FieldName, DataSourcename)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub Init(InitApplicationName As Variant, Optional ignore0 As Variant, Optional ignore1 As Variant, Optional Ignore2 As Variant, Optional Ignore3 As Variant)
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [Init], Use cp object to build page.")
'        '        ' deprecated - mainclass is now only a shell for addons and not part of the page rendering process
'        '        'Call cmc.main_Init(InitApplicationName, ignore0, ignore1, Ignore2, Ignore3)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub InitASPEnvironment(ASPResponse As Response, ignore_ASPRequest As Request, ignore_ASPServer As Server)
'        '        Call Err.Raise(KmaErrorBase, "mainClass", "Method Deprecated, [InitASPEnvironment]")
'        '        'Call cmc.main_InitASPEnvironment(ASPResponse, ignore_ASPRequest, ignore_ASPServer)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub InsertCSContentWatchListRecord(CSPointer As Variant, ContentWatchListName As Variant, Optional Link As Variant, Optional LinkLabel As Variant, Optional DateExpires As Variant)
'        '        Call cmc.main_InsertCSContentWatchListRecord(CSPointer, ContentWatchListName, Link, LinkLabel, DateExpires)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub LoadContentDefinition(ContentID As Variant)
'        '        Call cmc.main_LoadContentDefinition(ContentID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub LoadContentDefinitions()
'        '        Call cmc.main_LoadContentDefinitions
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub LogActivity(Message As String)
'        '        Call cmc.main_LogActivity(Message)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub LogActivity2(Message As String, SubjectMemberID As Long, SubjectOrganizationID As Long)
'        '        Call cmc.main_LogActivity2(Message, SubjectMemberID, SubjectOrganizationID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub LogoutMember()
'        '        Call cmc.main_LogoutMember
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub NextCSRecord(CSPointer As Variant, Optional AsyncSave As Variant)
'        '        Call cmc.main_NextCSRecord(CSPointer, AsyncSave)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub OnEndPage()
'        '        'Call cmc.main_OnEndPage
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub OnStartPage(myScriptingContext As ScriptingContext)
'        '        'Call cmc.main_OnStartPage(myScriptingContext)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub OpenMember(recordId As Variant)
'        '        Call cmc.main_OpenMember(recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PreloadContentPage(RootPageName As Variant, Optional RootContentName As Variant)
'        '        Call cmc.main_PreloadContentPage(RootPageName, RootContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintAccount()
'        '        ''Call cmc.main_PrintAccount
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintAdminForm()
'        '        ''Call cmc.main_PrintAdminForm
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintAdminPage()
'        '        'Call cmc.main_PrintAdminPage
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintAdminPageBottom(Optional AllowTools As Boolean)
'        '        'Call cmc.main_PrintAdminPageBottom(AllowTools)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintAdminPageTop(PageTitle As Variant)
'        '        'Call cmc.main_PrintAdminPageTop(PageTitle)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintBanner()
'        '        'Call cmc.main_PrintBanner
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintBookList(ListTitle As Variant, BookContentName As Variant, WhereClause As Variant)
'        '        'Call cmc.main_PrintBookList(ListTitle, BookContentName, WhereClause)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintCatalog(Optional ItemContentName As Variant, Optional CategoryContentName As Variant)
'        '        'Call cmc.main_PrintCatalog(ItemContentName, CategoryContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintCatalogCategoryListing(CategoryName As Variant, Optional ItemContentName, Optional CategoryContentName)
'        '        'Call cmc.main_PrintCatalogCategoryListing(CategoryName, ItemContentName, CategoryContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintCatalogFeaturedListing(Optional ItemContentName As Variant, Optional CategoryContentName As Variant)
'        '        'Call cmc.main_PrintCatalogFeaturedListing(ItemContentName, CategoryContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintCatalogNewListing(Optional ItemContentName As Variant, Optional CategoryContentName As Variant)
'        '        'Call cmc.main_PrintCatalogNewListing(ItemContentName, CategoryContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintContentBlock(BlockName As Variant, Optional ContentName As Variant)
'        '        'Call cmc.main_PrintContentBlock(BlockName, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintContentPage(RootPageName As Variant, Optional ContentName As Variant, Optional OrderByClause As Variant, Optional AllowChildPageList As Variant)
'        '        'Call cmc.main_PrintContentPage(RootPageName, ContentName, OrderByClause, AllowChildPageList)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintCSField(CSPointer As Variant, FieldName As Variant)
'        '        'Call cmc.main_PrintCSField(CSPointer, FieldName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintCSRecordAddLink(CSPointer As Variant, Optional Criteria As Variant)
'        '        'Call cmc.main_PrintCSRecordAddLink(CSPointer, Criteria)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintCSRecordEditLink(CSPointer As Variant)
'        '        'Call cmc.main_PrintCSRecordEditLink(CSPointer)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintDownloadList(ListTitle As Variant, ContentName As Variant, Optional Criteria As Variant, Optional SortField As Variant)
'        '        'Call cmc.main_PrintDownloadList(ListTitle, ContentName, Criteria, SortField)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintErrorMessage()
'        '        'Call cmc.main_PrintErrorMessage
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFeedbackForm(ContentName As Variant, recordId As Variant, ToMemberID As Variant, Optional Headline As Variant)
'        '        'Call cmc.main_PrintFeedbackForm(ContentName, recordId, ToMemberID, Headline)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormButton(ButtonLabel As Variant)
'        '        'Call cmc.main_PrintFormButton(ButtonLabel)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormCSHidden(CSPointer As Variant, FieldName As Variant)
'        '        'Call cmc.main_PrintFormCSHidden(CSPointer, FieldName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormCSInput(CSPointer As Variant, FieldName As Variant)
'        '        'Call cmc.main_PrintFormCSInput(CSPointer, FieldName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormEnd()
'        '        'Call cmc.main_PrintFormEnd
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormHidden(TagName As Variant, TagValue As Variant)
'        '        'Call cmc.main_PrintFormHidden(TagName, TagValue)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormInputCheckBox(TagName As Variant, Optional DefaultValue As Variant)
'        '        'Call cmc.main_PrintFormInputCheckBox(TagName, DefaultValue)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormInputFile(TagName As Variant)
'        '        'Call cmc.main_PrintFormInputFile(TagName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormInputRadioBox(TagName As Variant, TagValue As Variant, CurrentValue As Variant)
'        '        'Call cmc.main_PrintFormInputRadioBox(TagName, TagValue, CurrentValue)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormInputSelect(MenuName As Variant, CurrentValue As Variant, ContentName As Variant, Criteria As Variant, Optional NoneCaption As Variant)
'        '        'Call cmc.main_PrintFormInputSelect(MenuName, CurrentValue, ContentName, Criteria, NoneCaption)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormInputText(TagName As Variant, Optional DefaultValue As Variant, Optional Height As Variant, Optional Width As Variant, Optional IDElement As Variant)
'        '        'Call cmc.main_PrintFormInputText(TagName, DefaultValue, Height, Width, IDElement)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormMyAccount()
'        '        'Call cmc.main_PrintFormMyAccount
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintFormStart(Optional ActionQueryString As Variant)
'        '        'Call cmc.main_PrintFormStart(ActionQueryString)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintForum(ForumName As Variant, Optional ContentName As Variant)
'        '        'Call cmc.main_PrintForum(ForumName, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintForumList(PageTitle As Variant, Optional ContentName As Variant)
'        '        'Call cmc.main_PrintForumList(PageTitle, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintGroupMemberList(ListTitle As Variant, GroupName As Variant)
'        '        'Call cmc.main_PrintGroupMemberList(ListTitle, GroupName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintJoinForm()
'        '        'Call cmc.main_PrintJoinForm
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintLinkList(ListTitle As Variant, ContentName As Variant, Optional Criteria As Variant, Optional SortFieldList As Variant)
'        '        'Call cmc.main_PrintLinkList(ListTitle, ContentName, Criteria, SortFieldList)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintLoginForm()
'        '        'Call cmc.main_PrintLoginForm
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintLoginLink()
'        '        'Call cmc.main_PrintLoginLink
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintLoginPage()
'        '        'Call cmc.main_PrintLoginPage
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintMeetingPage(PageTitle As Variant, ContentName As Variant)
'        '        'Call cmc.main_PrintMeetingPage(PageTitle, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintMember(memberID As Variant)
'        '        'Call cmc.main_PrintMember(memberID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintMemberProfileForm()
'        '        'Call cmc.main_PrintMemberProfileForm
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintMoreInfo(contactMemberID As Variant)
'        '        'Call cmc.main_PrintMoreInfo(contactMemberID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintNewsletterList(ListTitle As Variant, Optional ContentName As Variant)
'        '        'Call cmc.main_PrintNewsletterList(ListTitle, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintOrderAccount()
'        '        'Call cmc.main_PrintOrderAccount
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintOrderForms()
'        '        'Call cmc.main_PrintOrderForms
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintOrderProcess()
'        '        'Call cmc.main_PrintOrderProcess
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintOrganization(OrganizationID As Variant)
'        '        'Call cmc.main_PrintOrganization(OrganizationID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintOrganizationList(ListTitle As Variant, Optional ContentName As Variant, Optional Criteria As Variant)
'        '        'Call cmc.main_PrintOrganizationList(ListTitle, ContentName, Criteria)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintPage(tableName As Variant, Criteria As Variant)
'        '        'Call cmc.main_PrintPage(tableName, Criteria)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintPageByName(PageName As Variant)
'        '        'Call cmc.main_PrintPageByName(PageName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintPageList(tableName As Variant, Criteria As Variant, RootPageID As Variant)
'        '        'Call cmc.main_PrintPageList(tableName, Criteria, RootPageID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintPageListByName(PageName As Variant)
'        '        'Call cmc.main_PrintPageListByName(PageName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintPanelBottom(ColorBase As Variant, ColorHilite As Variant, ColorShadow As Variant, Width As Variant, Padding As Variant)
'        '        'Call cmc.main_PrintPanelBottom(ColorBase, ColorHilite, ColorShadow, Width, Padding)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintPanelTop(ColorBase As Variant, ColorHilite As Variant, ColorShadow As Variant, Width As Variant, Padding As Variant)
'        '        'Call cmc.main_PrintPanelTop(ColorBase, ColorHilite, ColorShadow, Width, Padding)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintPopupMessage(Message As Variant, Optional WindowWidth As Variant, Optional WindowHeight As Variant, Optional WindowScrollBars As Variant, Optional WindowResizable As Variant, Optional WindowName As Variant)
'        '        'Call cmc.main_PrintPopupMessage(Message, WindowWidth, WindowHeight, WindowScrollBars, WindowResizable, WindowName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintPopupPage(Message As Variant, Optional WindowWidth As Variant, Optional WindowHeight As Variant, Optional WindowScrollBars As Variant, Optional WindowResizable As Variant, Optional WindowName As Variant)
'        '        'Call cmc.main_PrintPopupPage(Message, WindowWidth, WindowHeight, WindowScrollBars, WindowResizable, WindowName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintPromotionRegistration(Optional PasswordRequired As Variant)
'        '        'Call cmc.main_PrintPromotionRegistration(PasswordRequired)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintRecordAddLink(ContentName As Variant, PresetNameValueList As Variant)
'        '        'Call cmc.main_PrintRecordAddLink(ContentName, PresetNameValueList)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintRecordAddLinkByTable(tableName As Variant, Criteria As Variant)
'        '        'Call cmc.main_PrintRecordAddLinkByTable(tableName, Criteria)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintRecordEditLink(ContentName As Variant, recordId As Variant)
'        '        'Call cmc.main_PrintRecordEditLink(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintRecordEditLinkByContent(ContentID As Variant, recordId As Variant, Criteria As Variant)
'        '        'Call cmc.main_PrintRecordEditLinkByContent(ContentID, recordId, Criteria)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintRecordEditLinkByTable(tableName As Variant, RecordIDVariant As Variant, Criteria As Variant)
'        '        'Call cmc.main_PrintRecordEditLinkByTable(tableName, RecordIDVariant, Criteria)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintRecordList(ListTitle As Variant, ContentName As Variant, Optional Criteria As Variant, Optional SortFieldList As Variant)
'        '        'Call cmc.main_PrintRecordList(ListTitle, ContentName, Criteria, SortFieldList)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintReportsPage()
'        '        'Call cmc.main_PrintReportsPage
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintSeeAlso(ContentName As Variant, recordId As Variant)
'        '        'Call cmc.main_PrintSeeAlso(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintSurveyForm(SurveyName As Variant, Optional ContentName As Variant)
'        '        'Call cmc.main_PrintSurveyForm(SurveyName, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintSurveyList(ListTitle As Variant, Optional ContentName As Variant)
'        '        'Call cmc.main_PrintSurveyList(ListTitle, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintTitle(Title As Variant)
'        '        'Call cmc.main_PrintTitle(Title)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintToolsForm()
'        '        'Call cmc.main_PrintToolsForm
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintToolsPanel()
'        '        'Call cmc.main_PrintToolsPanel
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PrintWhatsNew(Optional SortFieldList As Variant)
'        '        'Call cmc.main_PrintWhatsNew(SortFieldList)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ProcessCheckList(TagName As String, PrimaryContentName As String, PrimaryRecordID As String, SecondaryContentName As String, RulesContentName As String, RulesPrimaryFieldname As String, RulesSecondaryFieldName As String)
'        '        Call cmc.main_ProcessCheckList(TagName, PrimaryContentName, PrimaryRecordID, SecondaryContentName, RulesContentName, RulesPrimaryFieldname, RulesSecondaryFieldName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ProcessFormInputTopics(TagName As Variant, TopicContentName As Variant, ContentName As Variant, recordId As Variant)
'        '        'Call cmc.main_ProcessFormInputTopics(TagName, TopicContentName, ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ProcessOrderForms()
'        '        'Call cmc.main_ProcessOrderForms
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ProcessRSS()
'        '        'Call cmc.main_ProcessRSS
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ProcessSpecialCaseAfterSave(IsDelete As Boolean, ContentName As String, recordId As Long, RecordName As String, RecordParentID As Long, UseContentWatchLink As Boolean)
'        '        Call cmc.main_ProcessSpecialCaseAfterSave(IsDelete, ContentName, recordId, RecordName, RecordParentID, UseContentWatchLink)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub PublishEdit(ContentName As Variant, recordId As Variant)
'        '        Call cmc.main_PublishEdit(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub QueueEmail(toAddress As Variant, fromAddress As Variant, subject As Variant, Body As Variant)
'        '        Call cmc.main_QueueEmail(toAddress, fromAddress, subject, Body)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub QueueSQL(DataSourcename As Variant, SQL As Variant)
'        '        Call cmc.main_QueueSQL(DataSourcename, SQL)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub Redirect(Link As Variant)
'        '        Call cmc.main_Redirect(Link)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub RedirectByRecord(ContentName As Variant, recordId As Variant, Optional FieldName As Variant)
'        '        Call cmc.main_RedirectByRecord(ContentName, recordId, FieldName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub RedirectHTTP(Link As Variant)
'        '        Call cmc.main_RedirectHTTP(Link)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub RemovePCCRow(pageId As Long, IsWorkflowRendering As Boolean, IsQuickEditing As Boolean)
'        '        Call cmc.main_RemovePCCRow(pageId, IsWorkflowRendering, IsQuickEditing)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub renameFile(sourcePathFilename As Variant, destinationFilename As Variant)
'        '        Call cmc.main_renameFile(sourcePathFilename, destinationFilename)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ReportError(cause As Variant)
'        '        Call cmc.main_ReportError(cause)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ReportError2(ErrorObject As Object, cause As Variant)
'        '        Call cmc.main_ReportError2(ErrorObject, cause)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub ReportError3(AppName As String, ClassName As String, MethodName As String, cause As String, Err_Number As Long, Err_Source As String, Err_Description As String, WillResumeAfterLogging As Boolean)
'        '        Call cmc.main_ReportError3(AppName, ClassName, MethodName, cause, Err_Number, Err_Source, Err_Description, WillResumeAfterLogging)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub RequestTask(Command As Variant, SQL As Variant, ExportName As Variant, Filename As Variant)
'        '        Call cmc.main_RequestTask(Command, SQL, ExportName, Filename)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub RollBackCS(CSPointer As Variant)
'        '        Call cmc.main_RollBackCS(CSPointer)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub RunContentDiagnostics_X()
'        '        'Call cmc.main_RunContentDiagnostics_X
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SaveBake(Name As Variant, Value As Variant, ContentNameList As Variant, Optional DateExpires As Variant)
'        '        Call cmc.main_SaveBake(Name, Value, ContentNameList, DateExpires)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SaveCSRecord(CSPointer As Variant, Optional AsyncSave As Variant)
'        '        Call cmc.main_SaveCSRecord(CSPointer, AsyncSave)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SaveFile(Filename As Variant, fileContent As Variant)
'        '        Call cmc.main_SaveFile(Filename, fileContent)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SaveMember()
'        '        Call cmc.main_SaveMember
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SaveMemberBase()
'        '        Call cmc.main_SaveMemberBase
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SaveVirtualFile(Filename As Variant, fileContent As Variant)
'        '        Call cmc.main_SaveVirtualFile(Filename, fileContent)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SaveVisit()
'        '        Call cmc.main_SaveVisit
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SaveVisitor()
'        '        Call cmc.main_SaveVisitor
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SendEmailConfirmation(emailId As Long, ConfirmationMemberID As Long)
'        '        Call cmc.main_SendEmailConfirmation(emailId, ConfirmationMemberID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SendFormEmail(SendTo As Variant, SendFrom As Variant, SendSubject As Variant)
'        '        Call cmc.main_SendFormEmail(SendTo, SendFrom, SendSubject)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SendGroupEmail(GroupList As Variant, fromAddress As Variant, subject As Variant, Body As Variant, Immediate As Variant, HTML As Variant)
'        '        Call cmc.main_SendGroupEmail(GroupList, fromAddress, subject, Body, Immediate, HTML)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SendNote(ToMemberID As Variant, FromName As Variant, FromEmail As Variant, subject As Variant, Body As Variant, DateExpires As Variant)
'        '        'Call cmc.main_SendNote(ToMemberID, FromName, FromEmail, subject, Body, DateExpires)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SendNote_Fast(ToMemberID As Long, FromName As String, FromEmail As String, subject As String, Body As String, DateExpires As Date)
'        '        'Call cmc.main_SendNote_Fast(ToMemberID, FromName, FromEmail, subject, Body, DateExpires)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SendPublishSubmitNotice(ContentName As String, recordId As Long, RecordName As String)
'        '        Call cmc.main_SendPublishSubmitNotice(ContentName, recordId, RecordName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SendSystemEmail(emailName As Variant, Optional AdditionalCopy As Variant, Optional AdditionalMemberID As Variant)
'        '        Call cmc.main_SendSystemEmail(emailName, AdditionalCopy, AdditionalMemberID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetContentCopy(CopyName As Variant, Content As Variant)
'        '        Call cmc.main_SetContentCopy(CopyName, Content)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetCS(CSPointer As Variant, FieldName As Variant, FieldValue As Variant)
'        '        Call cmc.main_SetCS(CSPointer, FieldName, FieldValue)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetCSField(CSPointer As Variant, FieldName As Variant, FieldValue As Variant)
'        '        Call cmc.main_SetCSField(CSPointer, FieldName, FieldValue)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetCSFormInput(CSPointer As Long, FieldName As String, Optional RequestName As String)
'        '        Call cmc.main_SetCSFormInput(CSPointer, FieldName, RequestName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetCSModified(CSPointer As Variant, Optional ModifiedByID As Variant)
'        '        Call cmc.main_SetCSModified(CSPointer, ModifiedByID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetCSTextFile(CSPointer As Variant, FieldName As Variant, Copy As Variant, ContentName As Variant)
'        '        Call cmc.main_SetCSTextFile(CSPointer, FieldName, Copy, ContentName)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetEditLock(ContentName As Variant, recordId As Variant)
'        '        Call cmc.main_SetEditLock(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetMember(PeopleID As Variant)
'        '        Call cmc.main_SetMember(PeopleID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetMemberProperty(PropertyName As Variant, Value As Variant, Optional TargetMemberID As Variant)
'        '        Call cmc.main_SetMemberProperty(PropertyName, Value, TargetMemberID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetMetaContent(ContentID As Variant, recordId As Variant)
'        '        Call cmc.main_SetMetaContent(ContentID, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetPromotion(PromotionKey As Variant)
'        '        'Call cmc.main_SetPromotion(PromotionKey)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetSiteProperty(FieldName As Variant, FieldValue As Variant, Optional AllowAdminAccess As Variant)
'        '        Call cmc.main_SetSiteProperty(FieldName, FieldValue, AllowAdminAccess)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetStreamBuffer(BufferOn As Variant)
'        '        Call cmc.main_SetStreamBuffer(BufferOn)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetStreamHeader(HeaderName As Variant, HeaderValue As Variant)
'        '        Call cmc.main_SetStreamHeader(HeaderName, HeaderValue)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetStreamStatus(status As String)
'        '        Call cmc.main_SetStreamStatus(status)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetStreamTimeout(TimeoutSeconds As Variant)
'        '        'Call cmc.main_SetStreamTimeout(TimeoutSeconds)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetStreamType(contentType As Variant)
'        '        Call cmc.main_SetStreamType(contentType)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetViewingProperty(PropertyName As String, Value As String)
'        '        Call cmc.main_SetViewingProperty(PropertyName, Value)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetVisitorProperty(PropertyName As Variant, Value As Variant, Optional TargetVisitorid As Variant)
'        '        Call cmc.main_SetVisitorProperty(PropertyName, Value, TargetVisitorid)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SetVisitProperty(PropertyName As Variant, Value As Variant, Optional TargetVisitID As Variant)
'        '        Call cmc.main_SetVisitProperty(PropertyName, Value, TargetVisitID)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub SubmitEdit(ContentName As Variant, recordId As Variant)
'        '        Call cmc.main_SubmitEdit(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        Public Sub TestPoint(Message As String)
'            cp.Site.TestPoint(Message)
'        End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub TestPointEnter(Message As Variant)
'        '        Call cmc.main_TestPointEnter(Message)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub TestPointExit(Optional Message)
'        '        Call cmc.main_TestPointExit(Message)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub TrackContent(ContentName As Variant, recordId As Variant)
'        '        Call cmc.main_TrackContent(ContentName, recordId)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub TrackContentSet(CSPointer As Variant, Optional Link As Variant)
'        '        Call cmc.main_TrackContentSet(CSPointer, Link)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub UpdatePCCRow(pageId As Long, IsWorkflowRendering As Boolean, IsQuickEditing As Boolean)
'        '        Call cmc.main_UpdatePCCRow(pageId, IsWorkflowRendering, IsQuickEditing)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub WritePleaseWaitEnd()
'        '        Call cmc.main_WritePleaseWaitEnd
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub WritePleaseWaitStart()
'        '        Call cmc.main_WritePleaseWaitStart
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub WriteStream(Message As Variant)
'        '        Call cmc.main_WriteStream(Message)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub WriteStreamComment(Message As Variant)
'        '        Call cmc.main_WriteStreamComment(Message)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub WriteStreamCookie(CookieName As Variant, CookieValue As Variant, Optional DateExpires As Variant, Optional domain As Variant, Optional Path As Variant, Optional Secure As Variant)
'        '        Call cmc.main_WriteStreamCookie(CookieName, CookieValue, DateExpires, domain, Path, Secure)
'        '    End Sub
'        '    '
'        '    '
'        '    '
'        '    Public Sub WriteStreamLine(Message As Variant)
'        '        Call cmc.main_WriteStreamLine(Message)
'        '    End Sub


'    End Class

'End Namespace




