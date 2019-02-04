Option Explicit On
Option Strict On

Imports Contensive.BaseClasses

Namespace Contensive.VbConversion
    Public Class ContentServerClass
        Private cp As CPBaseClass
        Public Sub New(cp As CPBaseClass)
            Me.cp = cp
        End Sub

        '        Private cmc As cpCom.comMainCsvClass
        '        '
        '        '
        '        '
        '        Public Enum contentTypeEnum
        '            contentTypeWeb = 1
        '            contentTypeEmail = 2
        '            contentTypeWebTemplate = 3
        '            contentTypeEmailTemplate = 4
        '        End Enum

        '        '
        '        '
        '        '
        '        Public Enum AddonContextEnum
        '            ' should have been addonContextPage, etc
        '            ContextPage = 1
        '            ContextAdmin = 2
        '            ContextTemplate = 3
        '            contextEmail = 4
        '            ContextRemoteMethod = 5
        '            ContextOnNewVisit = 6
        '            ContextOnPageEnd = 7
        '            ContextOnPageStart = 8
        '            ContextEditor = 9
        '            ContextHelpUser = 10
        '            ContextHelpAdmin = 11
        '            ContextHelpDeveloper = 12
        '            ContextOnContentChange = 13
        '            ContextFilter = 14
        '            ContextSimple = 15
        '            'ContextFunction = 15
        '            ContextOnBodyStart = 16
        '            ContextOnBodyEnd = 17
        '        End Enum

        '        '
        '        '
        '        '
        '        Public Enum SfImageResizeAlgorithms
        '            Box = 0
        '            Triangle = 1
        '            Hermite = 2
        '            Bell = 3
        '            BSpline = 4
        '            Lanczos3 = 5
        '            Mitchell = 6
        '            Stretch = 7
        '        End Enum

        '        '
        '        '
        '        '
        '        Public Type DataSourceConnType
        '    Conn As Connection
        '    IsOpen As Boolean
        '    Type As Long
        'End Type

        ''
        ''
        ''
        'Public Type CSConnectionType
        '    PeopleCID As Long
        '    ContentCID As Long
        '    ContentFieldsCID As Long
        '    OrganizationsCID As Long
        '    MemberRulesCID As Long
        '    GroupRulesCID As Long
        '    TopicRulesCID As Long
        '    MembersCID As Long
        '    ContentServerVersion As String
        '    OptionKeyString As String
        '    URLEncoder As String
        '    DomainName As String
        '    ConnectionHandle As Long
        '    RootPath As String
        '    ApplicationStatus As Long
        '    ServerPageDefault As String
        'End Type

        ''
        ''
        ''
        'Public Type VisitCacheType
        '    InUse As Boolean
        '    DateLastUsed As Date
        '    VisitID As Long
        '    VisitName As String
        '    VisitMemberID As Long
        '    VisitAuthenticated As Boolean
        '    VisitStartTime As Long
        '    VisitStartDateValue As Double
        '    VisitLastVisitTime As Date
        '    VisitStopTime As Date
        '    VisitPageVisits As Long
        '    VisitCookieSupport As Boolean
        '    VisitLoginAttempts As Long
        '    VisitorNew As Boolean
        '    VisitReferer As String
        '    VisitRemoteIP As String
        '    VisitMemberNew As Boolean
        '    VisitStateOK As Boolean
        '    VisitorID As Long
        '    VisitorName As String
        '    VisitorMemberID As Long
        '    VisitorOrderID As Long
        '    memberID As Long
        '    MemberName As String
        '    MemberAdmin As Boolean
        '    MemberDeveloper As Boolean
        '    MemberContentControlID As Long
        '    MemberAllowBulkEmail As String
        '    MemberAllowToolsPanel As String
        '    MemberAdminMenuModeID As Long
        '    MemberAutoLogin As String
        '    MemberSendNotes As Boolean
        '    MemberUsername As String
        '    MemberPassword As String
        '    MemberOrganizationID As Long
        '    MemberLanguage As String
        '    MemberLanguageID As Long
        '    MemberActive As Boolean
        '    MemberVisits As Long
        '    MemberLastVisit As Date
        '    MemberCompany As String
        '    MemberEmail As String
        '    MemberNew As Boolean
        'End Type

        ''
        ''
        ''
        'Public Type VisitEnvironmentType
        '    PageBrowser As String
        '    PageHTTPVia As String
        '    PageHTTPFrom As String
        '    PageRemoteIP As String
        'End Type

        ''
        ''
        ''
        'Public Type ContentSetRowCacheType
        '    Name As String
        '    Caption As String
        '    ValueVariant As Variant
        '    fieldType As Long
        '    Changed As Boolean                  ' If true, the next SaveCSRecord will save this field
        'End Type

        ''
        ''
        ''
        'Public Type ContentSetType
        '    IsOpen As Boolean                   ' If true, it is in use
        '    LastUsed As Date                    ' The date/time this ContentSet was last used
        '    Updateable As Boolean               ' Can not update an OpenCSSQL because Fields are not accessable
        '    NewRecord As Boolean                ' true if it was created here
        '    'ContentPointer As Long              ' Pointer to the content for this Set
        '    ContentName As String
        '    cDef As CDefType
        '    OwnerMemberID As Long               ' ID of the member who opened the ContentSet
        '    '
        '    ' Workflow editing modes
        '    '
        '    WorkflowAuthoringMode As Boolean    ' if true, these records came from the AuthoringTable, else ContentTable
        '    WorkflowEditingRequested As Boolean ' if true, the CS was opened requesting WorkflowEditingMode
        '    WorkflowEditingMode As Boolean      ' if true, the current record can be edited, else just rendered (effects EditBlank and SaveCSRecord)
        '    '
        '    ' ----- Write Cache
        '    '
        '    RowCacheChanged As Boolean          ' if true, RowCache contains changes
        '    RowCache() As ContentSetRowCacheType ' array of fields buffered for this set
        '    RowCacheSize As Long                ' the total number of fields in the row
        '    RowCacheCount As Long               ' the number of field() values to write
        '    IsModified As Boolean               ' Set when CS is opened and if a save happens
        '    '
        '    ' ----- Recordset used to retrieve the results
        '    '
        '    RS As Recordset                        ' The Recordset
        '    'RSOpen As Boolean                   ' true if the recordset is open
        '    'EOF As Boolean                      ' if true, Row is empty and at end of records
        '    ' ##### new way 4/19/2004
        '    '   ResultCache stores only the current row
        '    '   RS holds all other rows
        '    '   GetCSRow returns the ResultCache
        '    '   NextCSRecord saves the difference between the ResultCache and the RowCache, and movesnext, inc ResultachePointer
        '    '   LoadResultCache stores the current RS row to the ResultCache
        '    '
        '    '
        '    ' ##### old way
        '    ' Storage for the RecordSet results (future)
        '    '       Result - refers to the entire set of rows the the SQL (Source) returns
        '    '       ResultCache - the block of records currently stored in member (ResultCacheTop to ResultCacheTop+PageSize-1)
        '    '       ResultCache is initially loaded with PageSize records, starting on page PageNumber
        '    '       NextCSRecord increments ResultCachePointer
        '    '           If ResultCachePointer > ResultCacheRowCount-1 then LoadResultCache
        '    '       EOF true if ( ResultCachePointer > ResultCacheRowCount-1 ) and ( ResultCacheRowCount < PageSize )
        '    '
        '    Source As String                    ' Holds the SQL that created the result set
        '    DataSource As String                ' The Datasource of the SQL that created the result set
        '    PageSize As Long                    ' Number of records in a cache page
        '    PageNumber As Long                  ' The Page that this result starts with
        '    '
        '    ' ----- Read Cache
        '    '
        '    ResultColumnNames() As String       ' 1-D array of the result field names
        '    ResultColumnCount As Long           ' number of columns in the ResultColumnNames and ResultCacheValues
        '    ResultEOF As Boolean                ' Resultcache is at the last record
        '    ResultCacheValues() As Variant      ' 2-D array of the result rows/columns
        '    ResultCacheRowCount As Long         ' number of rows in the ResultCacheValues
        '    ResultCachePointer As Long          ' Pointer to the current result row, if 0, this is BOF
        '    '
        '    FieldPointer As Long                ' Used for GetFirstField, GetNextField, etc
        '    '
        '    SelectTableFieldList As String      ' comma delimited list of all fields selected, in the form table.field
        'End Type

        ''
        '' DebugMode
        ''
        'Public Property Let DebugMode(ByVal vNewValue As Boolean)
        '    cmc.csv_DebugMode = vNewValue
        'End Property
        '        '
        '        ' DebugMode
        '        '
        '        Public Property Get DebugMode() As Boolean
        '    DebugMode = cmc.csv_DebugMode
        'End Property
        '        '
        '        ' HostServiceProcessID
        '        '
        '        Public Property Let HostServiceProcessID(ByVal vNewValue As Long)
        '    cmc.csv_HostServiceProcessID = vNewValue
        'End Property
        '        '
        '        ' HostServiceProcessID
        '        '
        '        Public Property Get HostServiceProcessID() As Long
        '    HostServiceProcessID = cmc.csv_HostServiceProcessID
        'End Property
        '        '
        '        ' ApplicationNameLocal
        '        '
        '        Public Property Let ApplicationNameLocal(ByVal vNewValue As String)
        '    cmc.cmc_applicationName = vNewValue
        'End Property
        '        '
        '        ' ApplicationNameLocal
        '        '
        '        Public Property Get ApplicationNameLocal() As String
        '    ApplicationNameLocal = cmc.cmc_applicationName
        'End Property
        '        '
        '        ' ConnectionHandleLocal
        '        '
        '        Public Property Let ConnectionHandleLocal(ByVal vNewValue As Long)
        '    cmc.csv_ConnectionHandleLocal = vNewValue
        'End Property
        '        '
        '        ' ConnectionHandleLocal
        '        '
        '        Public Property Get ConnectionHandleLocal() As Long
        '    ConnectionHandleLocal = cmc.csv_ConnectionHandleLocal
        'End Property
        '        '
        '        ' ConnectionID
        '        '
        '        Public Property Let ConnectionID(ByVal vNewValue As Long)
        '    cmc.csv_ConnectionID = vNewValue
        'End Property
        '        '
        '        ' ConnectionID
        '        '
        '        Public Property Get ConnectionID() As Long
        '    ConnectionID = cmc.csv_ConnectionID
        'End Property
        '        '
        '        ' ContentSetCount
        '        '
        '        Public Property Let ContentSetCount(ByVal vNewValue As Long)
        '    cmc.csv_ContentSetCount = vNewValue
        'End Property
        '        '
        '        ' ContentSetCount
        '        '
        '        Public Property Get ContentSetCount() As Long
        '    ContentSetCount = cmc.csv_ContentSetCount
        'End Property
        '        '
        '        ' ContentSetSize
        '        '
        '        Public Property Let ContentSetSize(ByVal vNewValue As Long)
        '    cmc.csv_ContentSetSize = vNewValue
        'End Property
        '        '
        '        ' ContentSetSize
        '        '
        '        Public Property Get ContentSetSize() As Long
        '    ContentSetSize = cmc.csv_ContentSetSize
        'End Property
        '        '
        '        ' UpgradeInProgress
        '        '
        '        Public Property Get UpgradeInProgress() As Boolean
        '    UpgradeInProgress = cmc.csv_UpgradeInProgress
        'End Property
        '        '
        '        ' UpgradeInProgress
        '        '
        '        Public Property Let UpgradeInProgress(vNewValue As Boolean)
        '    cmc.csv_UpgradeInProgress = vNewValue
        'End Property
        '        '
        '        ' PhysicalFilePath
        '        '
        '        Public Property Get PhysicalFilePath() As String
        '    PhysicalFilePath = cmc.csv_PhysicalFilePath
        'End Property
        '        '
        '        ' PhysicalWWWPath
        '        '
        '        Public Property Get PhysicalWWWPath() As String
        '    PhysicalWWWPath = cmc.csv_PhysicalWWWPath
        'End Property
        '        '
        '        ' PhysicalWWWPath
        '        '
        '        Public Property Let PhysicalWWWPath(ByVal vNewValue As String)
        '    cmc.csv_PhysicalWWWPath = vNewValue
        'End Property
        '        '
        '        ' RootPath
        '        '
        '        Public Property Get RootPath() As String
        '    RootPath = cmc.csv_RootPath
        'End Property
        '        '
        '        ' DomainName
        '        '
        '        Public Property Get DomainName() As String
        '    DomainName = cmc.csv_DomainName
        'End Property
        '        '
        '        ' ErrorCount
        '        '
        '        Public Property Get ErrorCount() As Long
        '    ErrorCount = cmc.csv_ErrorCount
        'End Property
        '        '
        '        ' AppServicesProgress
        '        '
        '        Public Property Get AppServicesProgress() As String
        '    AppServicesProgress = cmc.csv_AppServicesProgress
        'End Property
        '        '
        '        ' AppServicesProgress
        '        '
        '        Public Property Let AppServicesProgress(ByVal vNewValue As String)
        '    cmc.csv_AppServicesProgress = vNewValue
        'End Property
        '        '
        '        ' AbortActivity
        '        '
        '        Public Property Get AbortActivity() As Boolean
        '    AbortActivity = cmc.csv_AbortActivity
        'End Property
        '        '
        '        ' KernelServicesVersion
        '        '
        '        Public Property Get KernelServicesVersion() As String
        '    KernelServicesVersion = cmc.csv_KernelServicesVersion
        'End Property
        '        '
        '        ' KernelServicesVersion
        '        '
        '        Public Property Let KernelServicesVersion(ByVal vNewValue As String)
        '    cmc.csv_KernelServicesVersion = vNewValue
        'End Property
        '        '
        '        ' SQLCommandTimeout
        '        '
        '        Public Property Get SQLCommandTimeout() As Long
        '    SQLCommandTimeout = cmc.csv_SQLCommandTimeout
        'End Property
        '        '
        '        ' SQLCommandTimeout
        '        '
        '        Public Property Let SQLCommandTimeout(ByVal vNewValue As Long)
        '    cmc.csv_SQLCommandTimeout = vNewValue
        'End Property
        '        '
        '        ' CDefConfigSaveNeeded
        '        '
        '        Public Property Get CDefConfigSaveNeeded() As Boolean
        '    CDefConfigSaveNeeded = cmc.csv_CDefConfigSaveNeeded
        'End Property
        '        '
        '        ' CDefConfigSaveNeeded
        '        '
        '        Public Property Let CDefConfigSaveNeeded(ByVal vNewValue As Boolean)
        '    cmc.csv_CDefConfigSaveNeeded = vNewValue
        'End Property
        '        '
        '        ' URLEncoder
        '        '
        '        Public Property Get URLEncoder() As String
        '    URLEncoder = cmc.csv_URLEncoder
        'End Property
        '        '
        '        ' URLEncoder
        '        '
        '        Public Property Let URLEncoder(ByVal vNewValue As String)
        '    cmc.csv_URLEncoder = vNewValue
        'End Property
        '        '
        '        ' ContentServerVersion
        '        '
        '        Public Property Get ContentServerVersion() As String
        '    ContentServerVersion = cmc.csv_ContentServerVersion
        'End Property
        '        '
        '        ' ContentServerVersion
        '        '
        '        Public Property Let ContentServerVersion(ByVal vNewValue As String)
        '    cmc.csv_ContentServerVersion = vNewValue
        'End Property
        '        '
        '        ' PhysicalContensivePath
        '        '
        '        Public Property Get PhysicalContensivePath() As String
        '    PhysicalContensivePath = cmc.csv_PhysicalContensivePath
        'End Property

        '        Public Property Set cmcObj(NewValue As Object)
        '            Set cmc = NewValue
        'End Property

        '        Public Sub Class_Initialize()
        '            ' no - cmc is created by and destroyed by cp
        '            'Call cmc.csv_initialize
        '        End Sub

        '        Public Sub Class_Terminate()
        '    'Call cmc.csv_terminate
        '    Set cmc = Nothing
        'End Sub
        '        '
        '        '
        '        '
        '        Public Function AbuseCheck(Key As String) As Boolean
        '            AbuseCheck = cmc.csv_AbuseCheck(Key)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function AddToEmailBlockList(EmailAddress As String)
        '            AddToEmailBlockList = cmc.csv_AddToEmailBlockList(EmailAddress)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function CheckFile(PathFilename As String) As Boolean
        '            CheckFile = cmc.csv_CheckFile(PathFilename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function CheckFileFolder(FolderPath As String) As Boolean
        '            CheckFileFolder = cmc.csv_CheckFileFolder(FolderPath)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ClearContentTimeStamp(ContentName As String) As String
        '            ClearContentTimeStamp = cmc.csv_ClearContentTimeStamp(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function CreateCSO() As ContentSetClass
        '            Call Err.Raise(KmaErrorInternal, "contentServerClass", "method deprecated")
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function DecodeActiveContent(SourceCopy As String) As String
        '            DecodeActiveContent = cmc.csv_DecodeActiveContent(SourceCopy)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function DecodeAddonOptionArgument(EncodedArg As String) As String
        '            DecodeAddonOptionArgument = cmc.csv_DecodeAddonOptionArgument(EncodedArg)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function DecodeContent(Source As String) As String
        '            DecodeContent = cmc.csv_DecodeContent(Source)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function DecodeKeyNumber(EncodedKey As String, URLEncoder As String) As Long
        '            DecodeKeyNumber = cmc.csv_DecodeKeyNumber(EncodedKey, URLEncoder)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function DecodeKeyTime(EncodedKey As String, URLEncoder As String) As Variant
        '            DecodeKeyTime = cmc.csv_DecodeKeyTime(EncodedKey, URLEncoder)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function decodeNvaArgument(EncodedArg As String) As String
        '            decodeNvaArgument = cmc.csv_decodeNvaArgument(EncodedArg)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeActiveContent(Source As String, PeopleID As Long, Optional CSFormattingContext As Variant, Optional AddLinkEID As Boolean, Optional EncodeCachableTags As Boolean, Optional EncodeImages As Boolean, Optional EncodeEditIcons As Boolean, Optional EncodeNonCachableTags As Boolean, Optional AddAnchorQuery As String, Optional ProtocolHostString As String) As String
        '            EncodeActiveContent = cmc.csv_EncodeActiveContent(Source, PeopleID, CSFormattingContext, AddLinkEID, EncodeCachableTags, EncodeImages, EncodeEditIcons, EncodeNonCachableTags, AddAnchorQuery, ProtocolHostString)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeActiveContent2(Source As String, PeopleID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, AddLinkEID As Boolean, EncodeCachableTags As Boolean, EncodeImages As Boolean, EncodeEditIcons As Boolean, EncodeNonCachableTags As Boolean, AddAnchorQuery As String, ProtocolHostString As String) As String
        '            EncodeActiveContent2 = cmc.csv_EncodeActiveContent2(Source, PeopleID, ContextContentName, ContextRecordID, ContextContactPeopleID, AddLinkEID, EncodeCachableTags, EncodeImages, EncodeEditIcons, EncodeNonCachableTags, AddAnchorQuery, ProtocolHostString)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeActiveContent3(Source As String, PeopleID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, AddLinkEID As Boolean, EncodeCachableTags As Boolean, EncodeImages As Boolean, EncodeEditIcons As Boolean, EncodeNonCachableTags As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean) As String
        '            EncodeActiveContent3 = cmc.csv_EncodeActiveContent3(Source, PeopleID, ContextContentName, ContextRecordID, ContextContactPeopleID, AddLinkEID, EncodeCachableTags, EncodeImages, EncodeEditIcons, EncodeNonCachableTags, AddAnchorQuery, ProtocolHostString, IsEmailContent)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeActiveContent4(Source As String, PeopleID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, AddLinkEID As Boolean, EncodeCachableTags As Boolean, EncodeImages As Boolean, EncodeEditIcons As Boolean, EncodeNonCachableTags As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean, AdminURL As String) As String
        '            EncodeActiveContent4 = cmc.csv_EncodeActiveContent4(Source, PeopleID, ContextContentName, ContextRecordID, ContextContactPeopleID, AddLinkEID, EncodeCachableTags, EncodeImages, EncodeEditIcons, EncodeNonCachableTags, AddAnchorQuery, ProtocolHostString, IsEmailContent, AdminURL)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeActiveContent5(Source As String, PeopleID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, AddLinkEID As Boolean, EncodeCachableTags As Boolean, EncodeImages As Boolean, EncodeEditIcons As Boolean, EncodeNonCachableTags As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean, AdminURL As String, personalizationIsAuthenticated As Boolean, Context As AddonContextEnum) As String
        '            EncodeActiveContent5 = cmc.csv_EncodeActiveContent5(Source, PeopleID, ContextContentName, ContextRecordID, ContextContactPeopleID, AddLinkEID, EncodeCachableTags, EncodeImages, EncodeEditIcons, EncodeNonCachableTags, AddAnchorQuery, ProtocolHostString, IsEmailContent, AdminURL, personalizationIsAuthenticated, Context)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeAddonOptionArgument(Arg As String) As String
        '            EncodeAddonOptionArgument = cmc.csv_EncodeAddonOptionArgument(Arg)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContent(Source As String, Optional memberID As Long, Optional CSFormattingContext As Variant, Optional PlainText As Boolean, Optional AddLinkEID As Boolean, Optional EncodeActiveFormatting As Boolean, Optional EncodeActiveImages As Boolean, Optional EncodeActiveEditIcons As Boolean, Optional EncodeActivePersonalization As Boolean, Optional AddAnchorQuery As String, Optional ProtocolHostString As String) As String
        '            EncodeContent = cmc.csv_EncodeContent(Source, memberID, CSFormattingContext, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContent2(Source As String, memberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String) As String
        '            EncodeContent2 = cmc.csv_EncodeContent2(Source, memberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContent3(Source As String, memberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean) As String
        '            EncodeContent3 = cmc.csv_EncodeContent3(Source, memberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString, IsEmailContent)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContent4(Source As String, memberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean, AdminURL As String) As String
        '            EncodeContent4 = cmc.csv_EncodeContent4(Source, memberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString, IsEmailContent, AdminURL)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContent5(Source As String, memberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean, AdminURL As String, ignore_DefaultWrapperID As Long) As String
        '            EncodeContent5 = cmc.csv_EncodeContent5(Source, memberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString, IsEmailContent, AdminURL, ignore_DefaultWrapperID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContent6(Source As String, memberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, AddAnchorQuery As String, ProtocolHostString As String, IsEmailContent As Boolean, AdminURL As String, ignore_DefaultWrapperID As Long, TemplateCaseOnly_Content As String) As String
        '            EncodeContent6 = cmc.csv_EncodeContent6(Source, memberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, AddAnchorQuery, ProtocolHostString, IsEmailContent, AdminURL, ignore_DefaultWrapperID, TemplateCaseOnly_Content)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContent7(Source As String, memberID As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, queryStringForLinkAppend As String, ProtocolHostString As String, IsEmailContent As Boolean, AdminURL As String, ignore_DefaultWrapperID As Long, TemplateCaseOnly_Content As String, Optional isAuthenticated As Boolean) As String
        '            EncodeContent7 = cmc.csv_EncodeContent7(Source, memberID, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, queryStringForLinkAppend, ProtocolHostString, IsEmailContent, AdminURL, ignore_DefaultWrapperID, TemplateCaseOnly_Content, isAuthenticated)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContent8(mainOrNothing As Object, Source As String, personalizationPeopleId As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, queryStringForLinkAppend As String, ProtocolHostString As String, IsEmailContent As Boolean, AdminURL As String, ignore_DefaultWrapperID As Long, ignore_TemplateCaseOnly_Content As String, isAuthenticated As Boolean, addonContext As AddonContextEnum) As String
        '            EncodeContent8 = cmc.csv_EncodeContent8(mainOrNothing, Source, personalizationPeopleId, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, queryStringForLinkAppend, ProtocolHostString, IsEmailContent, AdminURL, ignore_DefaultWrapperID, ignore_TemplateCaseOnly_Content, isAuthenticated, addonContext)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContent9(Source As String, personalizationPeopleId As Long, ContextContentName As String, ContextRecordID As Long, ContextContactPeopleID As Long, PlainText As Boolean, AddLinkEID As Boolean, EncodeActiveFormatting As Boolean, EncodeActiveImages As Boolean, EncodeActiveEditIcons As Boolean, EncodeActivePersonalization As Boolean, queryStringForLinkAppend As String, ProtocolHostString As String, IsEmailContent As Boolean, ignore_DefaultWrapperID As Long, ignore_TemplateCaseOnly_Content As String, Context As AddonContextEnum, personalizationIsAuthenticated As Boolean, mainOrNothing As Object, isEditingAnything As Boolean) As String
        '            EncodeContent9 = cmc.csv_EncodeContent9(Source, personalizationPeopleId, ContextContentName, ContextRecordID, ContextContactPeopleID, PlainText, AddLinkEID, EncodeActiveFormatting, EncodeActiveImages, EncodeActiveEditIcons, EncodeActivePersonalization, queryStringForLinkAppend, ProtocolHostString, IsEmailContent, ignore_DefaultWrapperID, ignore_TemplateCaseOnly_Content, Context, personalizationIsAuthenticated, mainOrNothing, isEditingAnything)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeContentUpgrades(Source As String) As String
        '            EncodeContentUpgrades = cmc.csv_EncodeContentUpgrades(Source)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeKeyNumber(Key As Long, EncodeTime As Date, URLEncoder As String) As String
        '            EncodeKeyNumber = cmc.csv_EncodeKeyNumber(Key, EncodeTime, URLEncoder)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function encodeNvaArgument(Arg As String) As String
        '            encodeNvaArgument = cmc.csv_encodeNvaArgument(Arg)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeSQLBoolean(Source As Variant) As String
        '            EncodeSQLBoolean = cmc.csv_EncodeSQLBoolean(Source)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeSQLDate(Source As Variant) As String
        '            EncodeSQLDate = cmc.csv_EncodeSQLDate(Source)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeSQLNumber(Source As Variant) As String
        '            EncodeSQLNumber = cmc.csv_EncodeSQLNumber(Source)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function EncodeSQLText(Source As Variant) As String
        '            EncodeSQLText = cmc.csv_EncodeSQLText(Source)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteActiveX(ProgramID As String, AddonCaption As String, mainObjOrNothing As Object, OptionString_ForObjectCall As String, OptionStringForDisplay As String, Return_AddonErrorMessage As String) As String
        '            ExecuteActiveX = cmc.csv_ExecuteActiveX(ProgramID, AddonCaption, OptionString_ForObjectCall, OptionStringForDisplay, Return_AddonErrorMessage)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteAddon(addonid As Long, AddonNameOrGuid As String, OptionString As String, Context As AddonContextEnum, HostContentName As String, HostRecordID As Long, HostFieldName As String, ACInstanceID As String, IsIncludeAddon As Boolean, DefaultWrapperID As Long, ignore_TemplateCaseOnly_PageContent As String, Return_StatusOK As Boolean, ignore_SetNothingObject As Object, ignore_addonCallingItselfIdList As String, mainObjOrNothing As Object, ignore_AddonsRunOnThisPageIdList As String) As String
        '            ExecuteAddon = cmc.csv_ExecuteAddon(addonid, AddonNameOrGuid, OptionString, Context, HostContentName, HostRecordID, HostFieldName, ACInstanceID, IsIncludeAddon, DefaultWrapperID, ignore_TemplateCaseOnly_PageContent, Return_StatusOK, ignore_SetNothingObject, ignore_addonCallingItselfIdList, mainObjOrNothing, ignore_AddonsRunOnThisPageIdList)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteAddon2(addonid As Long, AddonNameOrGuid As String, OptionString As String, Context As AddonContextEnum, HostContentName As String, HostRecordID As Long, HostFieldName As String, ACInstanceID As String, IsIncludeAddon As Boolean, DefaultWrapperID As Long, ignore_TemplateCaseOnly_PageContent As String, Return_StatusOK As Boolean, ignore_SetNothingObject As Object, ignore_addonCallingItselfIdList As String, mainObjOrNothing As Object, ignore_AddonsRunOnThisPageIdList As String, personalizationPeopleId As Long, personalizationIsAuthenticated As Boolean) As String
        '            ExecuteAddon2 = cmc.csv_ExecuteAddon2(addonid, AddonNameOrGuid, OptionString, Context, HostContentName, HostRecordID, HostFieldName, ACInstanceID, IsIncludeAddon, DefaultWrapperID, ignore_TemplateCaseOnly_PageContent, Return_StatusOK, ignore_SetNothingObject, ignore_addonCallingItselfIdList, mainObjOrNothing, ignore_AddonsRunOnThisPageIdList, personalizationPeopleId, personalizationIsAuthenticated)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteAddonAsProcess(AddonIDGuidOrName As String, Optional OptionString As String, Optional ignore_SetNothingObject As Object, Optional WaitForResults As Boolean) As String
        '            ExecuteAddonAsProcess = cmc.csv_ExecuteAddonAsProcess(AddonIDGuidOrName, OptionString, ignore_SetNothingObject, WaitForResults)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteAssembly(addonid As Long, AddonCaption As String, AssemblyFullName As String, CollectionGuid As String, cp As Object, Return_AddonErrorMessage As String) As String
        '            ExecuteAssembly = cmc.csv_ExecuteAssembly(addonid, AddonCaption, AssemblyFullName, CollectionGuid, cp, Return_AddonErrorMessage)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function executeContentCommands(mainOrNothing As Object, Source As String, Context As AddonContextEnum, personalizationPeopleId As Long, personalizationIsAuthenticated As Boolean, ByRef Return_ErrorMessage As String) As String
        '            executeContentCommands = cmc.csv_executeContentCommands(mainOrNothing, Source, Context, personalizationPeopleId, personalizationIsAuthenticated, Return_ErrorMessage)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteScript(Language As String, Code As String, EntryPoint As String, MainObj As Object, ignore_SetNothingObject As Object, Return_AddonErrorMessage As String) As String
        '            ExecuteScript = cmc.csv_ExecuteScript(Language, Code, EntryPoint, MainObj, ignore_SetNothingObject, Return_AddonErrorMessage)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteScript2(Language As String, Code As String, EntryPoint As String, MainObj As Object, ignore_SetNothingObject As Object, Return_AddonErrorMessage As String, ScriptingTimeout As Long, ScriptName As String) As String
        '            ExecuteScript2 = cmc.csv_ExecuteScript2(Language, Code, EntryPoint, MainObj, ignore_SetNothingObject, Return_AddonErrorMessage, ScriptingTimeout, ScriptName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteScript3(Language As String, Code As String, EntryPoint As String, mainOrNothing As Object, ignore_SetNothingObject As Object, Return_AddonErrorMessage As String, ScriptingTimeout As Long, ScriptName As String, ReplaceCnt As Long, ReplaceNames() As String, ReplaceValues() As String) As String
        '            ExecuteScript3 = cmc.csv_ExecuteScript3(Language, Code, EntryPoint, mainOrNothing, ignore_SetNothingObject, Return_AddonErrorMessage, ScriptingTimeout, ScriptName, ReplaceCnt, ReplaceNames(), ReplaceValues())
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteSQL(DataSourcename As String, SQL As String, Optional Retries As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Recordset
        '    Set ExecuteSQL = cmc.csv_ExecuteSQL(DataSourcename, SQL, Retries, PageSize, PageNumber)
        'End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteSQLCommand(DataSourcename As String, SQL As String, Optional CommandTimeout As Long, Optional PageSize As Variant, Optional PageNumber As Variant) As Recordset
        '    Set ExecuteSQLCommand = cmc.csv_ExecuteSQLCommand(DataSourcename, SQL, CommandTimeout, PageSize, PageNumber)
        'End Function
        '        '
        '        '
        '        '
        '        Public Function ExecuteTrapLessSQL(DataSourcename As String, SQL As String, Optional Retries As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Recordset
        '    Set ExecuteTrapLessSQL = cmc.csv_ExecuteTrapLessSQL(DataSourcename, SQL, Retries, PageSize, PageNumber)
        'End Function
        '        '
        '        '
        '        '
        '        Public Function FilterDomainName(Link As String) As String
        '            FilterDomainName = cmc.csv_FilterDomainName(Link)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function FilterQueryString(Link As String) As String
        '            FilterQueryString = cmc.csv_FilterQueryString(Link)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetAddonOption(OptionName As String, OptionString As String) As String
        '            GetAddonOption = cmc.csv_GetAddonOption(OptionName, OptionString)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetAddonOptionStringValue(OptionName As String, addonOptionString As String) As String
        '            GetAddonOptionStringValue = cmc.csv_GetAddonOptionStringValue(OptionName, addonOptionString)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetAddonSelector(SrcOptionName As String, InstanceOptionValue_AddonEncoded As String, SrcOptionValueSelector As String) As String
        '            GetAddonSelector = cmc.csv_GetAddonSelector(SrcOptionName, InstanceOptionValue_AddonEncoded, SrcOptionValueSelector)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCacheFilename(Name As String) As String
        '            GetCacheFilename = cmc.csv_GetCacheFilename(Name)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCDefAdminColumns(ContentName As String) As Variant
        '            GetCDefAdminColumns = cmc.csv_GetCDefAdminColumns(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetConnectionString(DataSourcename As String) As String
        '            GetConnectionString = cmc.csv_GetConnectionString(DataSourcename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentControlCriteria(ContentName As String) As String
        '            GetContentControlCriteria = cmc.csv_GetContentControlCriteria(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentCopy(CopyName As String, ContentName As String, DefaultContent As String, personalizationPeopleId As Long) As String
        '            GetContentCopy = cmc.csv_GetContentCopy(CopyName, ContentName, DefaultContent, personalizationPeopleId)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentCopy2(CopyName As String, ContentName As String, DefaultContent As String, personalizationPeopleId As Long, mainOrNothing As Object, AllowEditWrapper As Boolean) As String
        '            GetContentCopy2 = cmc.csv_GetContentCopy2(CopyName, ContentName, DefaultContent, personalizationPeopleId, mainOrNothing, AllowEditWrapper)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentCopy3(CopyName As String, ContentName As String, DefaultContent As String, personalizationPeopleId As Long, mainOrNothing As Object, AllowEditWrapper As Boolean, personalizationIsAuthenticated As Boolean) As String
        '            GetContentCopy3 = cmc.csv_GetContentCopy3(CopyName, ContentName, DefaultContent, personalizationPeopleId, mainOrNothing, AllowEditWrapper, personalizationIsAuthenticated)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentDataSource(ContentName As String) As String
        '            GetContentDataSource = cmc.csv_GetContentDataSource(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentDataSourceByPointer(ContentPointer As Long) As String
        '            GetContentDataSourceByPointer = cmc.csv_GetContentDataSourceByPointer(ContentPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentDefinition(ByVal ContentName As String) As CDefType
        '            GetContentDefinition = cmc.csv_GetContentDefinition(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentFieldCount(ContentPointer As Long) As Long
        '            GetContentFieldCount = cmc.csv_GetContentFieldCount(ContentPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentID(ContentName As String) As Long
        '            GetContentID = cmc.csv_GetContentID(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentIDByTablename(tableName As String) As Long
        '            GetContentIDByTablename = cmc.csv_GetContentIDByTablename(tableName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentNameByID(ContentID As Long) As String
        '            GetContentNameByID = cmc.csv_GetContentNameByID(ContentID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentPointer(ContentName As String) As Long
        '            GetContentPointer = cmc.csv_GetContentPointer(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentPointerByID(ContentID As Long) As Long
        '            GetContentPointerByID = cmc.csv_GetContentPointerByID(ContentID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentPointerByTablename(tableName As String) As Long
        '            GetContentPointerByTablename = cmc.csv_GetContentPointerByTablename(tableName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentRows(ContentName As String, Optional Criteria As Variant, Optional SortFieldList As Variant, Optional ActiveOnly As Variant, Optional memberID As Long, Optional WorkflowRenderingMode As Boolean, Optional WorkflowEditingMode As Boolean, Optional SelectFieldList As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Variant
        '            GetContentRows = cmc.csv_GetContentRows(ContentName, Criteria, SortFieldList, ActiveOnly, memberID, WorkflowRenderingMode, WorkflowEditingMode, SelectFieldList, PageSize, PageNumber)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentTableID(ContentName As String) As Long
        '            GetContentTableID = cmc.csv_GetContentTableID(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetContentTablename(ContentName As String) As String
        '            GetContentTablename = cmc.csv_GetContentTablename(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCS(CSPointer As Long, FieldName As String) As String
        '            GetCS = cmc.csv_GetCS(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSBoolean(CSPointer As Long, FieldName As String) As Boolean
        '            GetCSBoolean = cmc.csv_GetCSBoolean(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSContentName(CSPointer As Long) As String
        '            GetCSContentName = cmc.csv_GetCSContentName(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSDate(CSPointer As Long, FieldName As String) As Date
        '            GetCSDate = cmc.csv_GetCSDate(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSField(CSPointer As Long, FieldName As String) As Variant
        '            GetCSField = cmc.csv_GetCSField(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSFieldCaption(CSPointer As Long, FieldName As String) As String
        '            GetCSFieldCaption = cmc.csv_GetCSFieldCaption(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSFieldType(CSPointer As Long, FieldName As String) As Long
        '            GetCSFieldType = cmc.csv_GetCSFieldType(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSFilename(CSPointer As Long, FieldName As String, OriginalFilename As String, Optional ContentName As String) As String
        '            GetCSFilename = cmc.csv_GetCSFilename(CSPointer, FieldName, OriginalFilename, ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSFirstFieldName(CSPointer As Long) As String
        '            GetCSFirstFieldName = cmc.csv_GetCSFirstFieldName(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSInteger(CSPointer As Long, FieldName As String) As Long
        '            GetCSInteger = cmc.csv_GetCSInteger(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSLookup(CSPointer As Long, FieldName As String) As String
        '            GetCSLookup = cmc.csv_GetCSLookup(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSNextFieldName(CSPointer As Long) As String
        '            GetCSNextFieldName = cmc.csv_GetCSNextFieldName(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSNumber(CSPointer As Long, FieldName As String) As Double
        '            GetCSNumber = cmc.csv_GetCSNumber(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSRow(CSPointer As Long) As Variant
        '            GetCSRow = cmc.csv_GetCSRow(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSRowCount(CSPointer As Long) As Long
        '            GetCSRowCount = cmc.csv_GetCSRowCount(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSRowFields(CSPointer As Long) As Variant
        '            GetCSRowFields = cmc.csv_GetCSRowFields(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSRows(CSPointer As Long) As Variant
        '            GetCSRows = cmc.csv_GetCSRows(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSSelectFieldList(CSPointer As Long) As String
        '            GetCSSelectFieldList = cmc.csv_GetCSSelectFieldList(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSSource(CSPointer As Long) As String
        '            GetCSSource = cmc.csv_GetCSSource(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSText(CSPointer As Long, FieldName As String) As String
        '            GetCSText = cmc.csv_GetCSText(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetCSTextFile(CSPointer As Long, FieldName As String) As String
        '            GetCSTextFile = cmc.csv_GetCSTextFile(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDataBuildDbVersion() As String
        '            GetDataBuildDbVersion = cmc.csv_GetDataBuildDbVersion()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDataBuildVersion()
        '            GetDataBuildVersion = cmc.csv_GetDataBuildVersion()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDataSourceByID(DataSourceID As Long) As String
        '            GetDataSourceByID = cmc.csv_GetDataSourceByID(DataSourceID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDataSourceID(DataSourcename As String) As Long
        '            GetDataSourceID = cmc.csv_GetDataSourceID(DataSourcename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDataSourceList() As String
        '            GetDataSourceList = cmc.csv_GetDataSourceList()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDataSourcePointer(DataSourcename As String) As Long
        '            GetDataSourcePointer = cmc.csv_GetDataSourcePointer(DataSourcename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDataSourceType(DataSourcename As String) As Long
        '            GetDataSourceType = cmc.csv_GetDataSourceType(DataSourcename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDefaultValue(Key As String) As String
        '            GetDefaultValue = cmc.csv_GetDefaultValue(Key)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDriveFreeSpace(DriveLetter As String) As Double
        '            GetDriveFreeSpace = cmc.csv_GetDriveFreeSpace(DriveLetter)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDynamicFormACSelect() As String
        '            GetDynamicFormACSelect = cmc.csv_GetDynamicFormACSelect()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetDynamicMenuACSelect() As String
        '            GetDynamicMenuACSelect = cmc.csv_GetDynamicMenuACSelect()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetEditLock(ContentName As String, recordId As Long, ReturnMemberID As Long, ReturnDateExpires As Date) As Boolean
        '            GetEditLock = cmc.csv_GetEditLock(ContentName, recordId, ReturnMemberID, ReturnDateExpires)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getEmailStyles(emailId As Variant) As String
        '            getEmailStyles = cmc.csv_getEmailStyles(emailId)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetEncodeContent_HeadTags() As String
        '            GetEncodeContent_HeadTags = cmc.csv_GetEncodeContent_HeadTags()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetEncodeContent_JavascriptBodyEnd() As String
        '            GetEncodeContent_JavascriptBodyEnd = cmc.csv_GetEncodeContent_JavascriptBodyEnd()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetEncodeContent_JavascriptInHead() As String
        '            GetEncodeContent_JavascriptInHead = cmc.csv_GetEncodeContent_JavascriptInHead()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetEncodeContent_JavascriptOnLoad() As String
        '            GetEncodeContent_JavascriptOnLoad = cmc.csv_GetEncodeContent_JavascriptOnLoad()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetEncodeContent_JSFilename() As String
        '            GetEncodeContent_JSFilename = cmc.csv_GetEncodeContent_JSFilename()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetEncodeContent_StyleFilenames() As String
        '            GetEncodeContent_StyleFilenames = cmc.csv_GetEncodeContent_StyleFilenames()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetFieldDescriptorByType(fieldType As Long) As String
        '            GetFieldDescriptorByType = cmc.csv_GetFieldDescriptorByType(fieldType)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetFieldTypeByDescriptor(FieldDescriptor As String) As Long
        '            GetFieldTypeByDescriptor = cmc.csv_GetFieldTypeByDescriptor(FieldDescriptor)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetFileCount(FolderPath As String) As Long
        '            GetFileCount = cmc.csv_GetFileCount(FolderPath)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetFileList(FolderPath As String, Optional PageSize As Long, Optional PageNumber As Long) As String
        '            GetFileList = cmc.csv_GetFileList(FolderPath, PageSize, PageNumber)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetFolderList(FolderPath As String) As String
        '            GetFolderList = cmc.csv_GetFolderList(FolderPath)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetFolderList2(FolderPath As String) As String
        '            GetFolderList2 = cmc.csv_GetFolderList2(FolderPath)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getGroupEmailSQL(ToAll As Boolean, emailId As Long) As String
        '            getGroupEmailSQL = cmc.csv_getGroupEmailSQL(ToAll, emailId)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetLinkedText(AnchorTag As String, AnchorText As String) As String
        '            GetLinkedText = cmc.csv_GetLinkedText(AnchorTag, AnchorText)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getMoreInfoHtml(PeopleID As Long)
        '            getMoreInfoHtml = cmc.csv_getMoreInfoHtml(PeopleID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getNvaValue(Name As String, nvaEncodedString As String) As String
        '            getNvaValue = cmc.csv_getNvaValue(Name, nvaEncodedString)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetParentIDFromNameSpace(ContentName As String, NameSpace As String) As Long
        '            GetParentIDFromNameSpace = cmc.csv_GetParentIDFromNameSpace(ContentName, NameSpace)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetPersistentVariant(Key As String) As Variant
        '            GetPersistentVariant = cmc.csv_GetPersistentVariant(Key)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getPhysicalFilename(VirtualFilename As String) As String
        '            getPhysicalFilename = cmc.csv_getPhysicalFilename(VirtualFilename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetProductCode(LicenseKey As String) As Long
        '            GetProductCode = cmc.csv_GetProductCode(LicenseKey)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetProperty(typeId As Long, KeyID As Long, Name As String, memberID As Long, Optional DefaultValue As Variant) As String
        '            GetProperty = cmc.csv_GetProperty(typeId, KeyID, Name, memberID, DefaultValue)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetRandomLong() As Long
        '            GetRandomLong = cmc.csv_GetRandomLong()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetRecordID(ContentName As String, RecordName As String) As Long
        '            GetRecordID = cmc.csv_GetRecordID(ContentName, RecordName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetRecordIDByGuid(ContentName As String, RecordGuid As String) As Long
        '            GetRecordIDByGuid = cmc.csv_GetRecordIDByGuid(ContentName, RecordGuid)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetRecordName(ContentName As String, recordId As Long) As String
        '            GetRecordName = cmc.csv_GetRecordName(ContentName, recordId)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetRSField(RS As Recordset, FieldName As String) As Variant
        '            GetRSField = cmc.csv_GetRSField(RS, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetSiteProperty(PropertyName As String, DefaultValue As String, memberID As Long, Optional AllowAdminAccess As Boolean) As String
        '            GetSiteProperty = cmc.csv_GetSiteProperty(PropertyName, DefaultValue, memberID, AllowAdminAccess)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetSQLAlterColumnType(DataSourcename As String, fieldType As Long) As String
        '            GetSQLAlterColumnType = cmc.csv_GetSQLAlterColumnType(DataSourcename, fieldType)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetSQLIndexList(DataSourcename As String, tableName As String) As String
        '            GetSQLIndexList = cmc.csv_GetSQLIndexList(DataSourcename, tableName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetSQLSelect(DataSourcename As String, From As String, Optional FieldList As String, Optional Where As String, Optional OrderBy As String, Optional GroupBy As String, Optional RecordLimit As Long) As String
        '            GetSQLSelect = cmc.csv_GetSQLSelect(DataSourcename, From, FieldList, Where, OrderBy, GroupBy, RecordLimit)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetSQLTableList(DataSourcename As String) As String
        '            GetSQLTableList = cmc.csv_GetSQLTableList(DataSourcename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getStyleSheet() As String
        '            getStyleSheet = cmc.csv_getStyleSheet()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getStyleSheet2(contentType As contentTypeEnum, Optional templateId As Long, Optional emailId As Long) As String
        '            getStyleSheet2 = cmc.csv_getStyleSheet2(contentType, templateId, emailId)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getStyleSheetDefault() As String
        '            getStyleSheetDefault = cmc.csv_getStyleSheetDefault()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getStyleSheetProcessed() As String
        '            getStyleSheetProcessed = cmc.csv_getStyleSheetProcessed()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetVirtualFileCount(FolderPath As String) As Long
        '            GetVirtualFileCount = cmc.csv_GetVirtualFileCount(FolderPath)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function getVirtualFileLink(serverFilePath As String, virtualFile As String) As String
        '            getVirtualFileLink = cmc.csv_getVirtualFileLink(serverFilePath, virtualFile)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetVirtualFileList(FolderPath As String, Optional PageSize As Long, Optional PageNumber As Long) As String
        '            GetVirtualFileList = cmc.csv_GetVirtualFileList(FolderPath, PageSize, PageNumber)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetVirtualFilename(ContentName As String, FieldName As String, recordId As Long, Optional OriginalFilename As Variant) As String
        '            GetVirtualFilename = cmc.csv_GetVirtualFilename(ContentName, FieldName, recordId, OriginalFilename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetVirtualFilenameByTable(tableName As String, FieldName As String, recordId As Long, OriginalFilename As String, fieldType As Long) As String
        '            GetVirtualFilenameByTable = cmc.csv_GetVirtualFilenameByTable(tableName, FieldName, recordId, OriginalFilename, fieldType)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetVirtualFolderList(FolderPath As String) As String
        '            GetVirtualFolderList = cmc.csv_GetVirtualFolderList(FolderPath)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetWordSearchExcludeList() As String
        '            GetWordSearchExcludeList = cmc.csv_GetWordSearchExcludeList()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetXMLContentDefinition(Optional ContentName As String) As String
        '            GetXMLContentDefinition = cmc.csv_GetXMLContentDefinition(ContentName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetXMLContentDefinition3(Optional ContentName As String, Optional IncludeBaseFields As Boolean) As String
        '            GetXMLContentDefinition3 = cmc.csv_GetXMLContentDefinition3(ContentName, IncludeBaseFields)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function InsertContentRecordByPointer(ContentPointer As Long, memberID As Long) As Long
        '            InsertContentRecordByPointer = cmc.csv_InsertContentRecordByPointer(ContentPointer, memberID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function InsertContentRecordGetID(ContentName As String, memberID As Long) As Long
        '            InsertContentRecordGetID = cmc.csv_InsertContentRecordGetID(ContentName, memberID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function InsertCSRecord(ContentName As String, memberID As Long) As Long
        '            InsertCSRecord = cmc.csv_InsertCSRecord(ContentName, memberID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function InsertTableRecordGetID(DataSourcename As String, tableName As String, memberID As Long) As Long
        '            InsertTableRecordGetID = cmc.csv_InsertTableRecordGetID(DataSourcename, tableName, memberID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function InsertTableRecordGetRS(DataSourcename As String, tableName As String, memberID As Long) As Recordset
        '    Set InsertTableRecordGetRS = cmc.csv_InsertTableRecordGetRS(DataSourcename, tableName, memberID)
        'End Function
        '        '
        '        '
        '        '
        '        Public Function IsContentFieldSupported(ContentName As String, FieldName As String) As Boolean
        '            IsContentFieldSupported = cmc.csv_IsContentFieldSupported(ContentName, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function IsCSFieldSupported(CSPointer As Long, FieldName As String) As Boolean
        '            IsCSFieldSupported = cmc.csv_IsCSFieldSupported(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function IsCSOK(CSPointer As Long) As Boolean
        '            IsCSOK = cmc.csv_IsCSOK(CSPointer)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function IsGroupIDListMember(memberID As Long, isAuthenticated As Boolean, GroupIDList As String) As Boolean
        '            IsGroupIDListMember = cmc.csv_IsGroupIDListMember(memberID, isAuthenticated, GroupIDList)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function IsGroupIDListMember2(memberID As Long, isAuthenticated As Boolean, GroupIDList As String, adminReturnsTrue As Boolean) As Boolean
        '            IsGroupIDListMember2 = cmc.csv_IsGroupIDListMember2(memberID, isAuthenticated, GroupIDList, adminReturnsTrue)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function IsLicenseOK(LicenseKey As String, DomainName As String) As Boolean
        '            IsLicenseOK = cmc.csv_IsLicenseOK(LicenseKey, DomainName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function IsRecordLocked(ContentName As String, recordId As Long, memberID As Long) As Boolean
        '            IsRecordLocked = cmc.csv_IsRecordLocked(ContentName, recordId, memberID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function IsSQLTable(DataSourcename As String, tableName As String) As Boolean
        '            IsSQLTable = cmc.csv_IsSQLTable(DataSourcename, tableName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function IsSQLTableField(DataSourcename As String, tableName As String, FieldName As String) As Boolean
        '            IsSQLTableField = cmc.csv_IsSQLTableField(DataSourcename, tableName, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function IsWithinContent(ChildContentID As Long, ParentContentID As Long) As Boolean
        '            IsWithinContent = cmc.csv_IsWithinContent(ChildContentID, ParentContentID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function MergeTemplate(EncodedTemplateHTML As String, EncodedContentHTML As String, memberID As Long) As String
        '            MergeTemplate = cmc.csv_MergeTemplate(EncodedTemplateHTML, EncodedContentHTML, memberID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenConnection(Applicationname As String) As CSConnectionType
        '            OpenConnection = OpenConnection2(Applicationname, Nothing)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenConnection2(Applicationname As String, cpOrNothing As Object) As CSConnectionType
        '            Dim c As csv_CSConnectionType
        '            c = cmc.csv_OpenConnection2(Applicationname, cpOrNothing)
        '            OpenConnection2.ApplicationStatus = c.ApplicationStatus
        '            OpenConnection2.ConnectionHandle = c.ConnectionHandle
        '            OpenConnection2.ContentCID = c.ContentCID
        '            OpenConnection2.ContentFieldsCID = c.ContentFieldsCID
        '            OpenConnection2.ContentServerVersion = c.ContentServerVersion
        '            OpenConnection2.DomainName = c.DomainName
        '            OpenConnection2.GroupRulesCID = c.GroupRulesCID
        '            OpenConnection2.MemberRulesCID = c.MemberRulesCID
        '            OpenConnection2.MembersCID = c.MembersCID
        '            OpenConnection2.OptionKeyString = c.OptionKeyString
        '            OpenConnection2.OrganizationsCID = c.OrganizationsCID
        '            OpenConnection2.PeopleCID = c.PeopleCID
        '            OpenConnection2.RootPath = c.RootPath
        '            OpenConnection2.ServerPageDefault = c.ServerPageDefault
        '            OpenConnection2.TopicRulesCID = c.TopicRulesCID
        '            OpenConnection2.URLEncoder = c.URLEncoder
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenCSContent(ContentName As String, Optional Criteria As Variant, Optional SortFieldList As Variant, Optional ActiveOnly As Variant, Optional memberID As Long, Optional WorkflowRenderingMode As Boolean, Optional WorkflowEditingMode As Boolean, Optional SelectFieldList As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
        '            OpenCSContent = cmc.csv_OpenCSContent(ContentName, Criteria, SortFieldList, ActiveOnly, memberID, WorkflowRenderingMode, WorkflowEditingMode, SelectFieldList, PageSize, PageNumber)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenCSContentRecord(ContentName As String, recordId As Long, Optional memberID As Long, Optional WorkflowAuthoringMode As Boolean, Optional WorkflowEditingMode As Boolean, Optional SelectFieldList As Variant) As Long
        '            OpenCSContentRecord = cmc.csv_OpenCSContentRecord(ContentName, recordId, memberID, WorkflowAuthoringMode, WorkflowEditingMode, SelectFieldList)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenCSContentWatchList(ListName As Variant, Optional SortFieldList As Variant, Optional ActiveOnly As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
        '            OpenCSContentWatchList = cmc.csv_OpenCSContentWatchList(ListName, SortFieldList, ActiveOnly, PageSize, PageNumber)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function openCSGroupUsers(GroupName As String, IsList As Boolean, Optional sqlCriteria As String, Optional SortFieldList As String, Optional ActiveOnly As Variant, Optional PageSize As Long, Optional PageNumber As Long) As Long
        '            openCSGroupUsers = cmc.csv_OpenCSGroupUsers(GroupName, IsList, sqlCriteria, SortFieldList, ActiveOnly, PageSize, PageNumber)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenCSJoin(CSPointer As Long, FieldName As String) As Long
        '            OpenCSJoin = cmc.csv_OpenCSJoin(CSPointer, FieldName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenCSSQL(ByVal DataSourcename As String, ByVal SQL As String, memberID As Long, Optional PageSize As Variant, Optional PageNumber As Variant) As Long
        '            OpenCSSQL = cmc.csv_OpenCSSQL(DataSourcename, SQL, memberID, PageSize, PageNumber)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenCSTextSearch(KeywordList As String, TopicIDList As String, ContentIDList As String, VisitID As Long, Optional PageSize As Variant, Optional PageNumber As Variant, Optional LanguageID As Variant) As String
        '            OpenCSTextSearch = cmc.csv_OpenCSTextSearch(KeywordList, TopicIDList, ContentIDList, VisitID, PageSize, PageNumber, LanguageID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenRSSQL(DataSourcename As String, SQL As String, Optional CommandTimeout As Long, Optional PageSize As Variant, Optional PageNumber As Variant, Optional UseCompatibleCursor As Variant, Optional UseServerCursor As Variant) As Recordset
        '    Set OpenRSSQL = cmc.csv_OpenRSSQL(DataSourcename, SQL, CommandTimeout, PageSize, PageNumber, UseCompatibleCursor, UseServerCursor)
        'End Function
        '        '
        '        '
        '        '
        '        Public Function OpenRSTable(DataSourcename As String, tableName As String, Criteria As String, SortFieldList As String, Optional SelectFieldList As Variant, Optional PageSize As Variant, Optional PageNumber As Variant) As Recordset
        '    Set OpenRSTable = cmc.csv_OpenRSTable(DataSourcename, tableName, Criteria, SortFieldList, SelectFieldList, PageSize, PageNumber)
        'End Function
        '        '
        '        '
        '        '
        '        Public Function OpenWebConnection(Applicationname As String) As CSConnectionType
        '            OpenWebConnection = OpenWebConnection2(Applicationname, Nothing)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function OpenWebConnection2(Applicationname As String, cpOrNothing As Object) As CSConnectionType
        '            Dim c As csv_CSConnectionType
        '            c = cmc.csv_OpenWebConnection2(Applicationname, cpOrNothing)
        '            OpenWebConnection2.ApplicationStatus = c.ApplicationStatus
        '            OpenWebConnection2.ConnectionHandle = c.ConnectionHandle
        '            OpenWebConnection2.ContentCID = c.ContentCID
        '            OpenWebConnection2.ContentFieldsCID = c.ContentFieldsCID
        '            OpenWebConnection2.ContentServerVersion = c.ContentServerVersion
        '            OpenWebConnection2.DomainName = c.DomainName
        '            OpenWebConnection2.GroupRulesCID = c.GroupRulesCID
        '            OpenWebConnection2.MemberRulesCID = c.MemberRulesCID
        '            OpenWebConnection2.MembersCID = c.MembersCID
        '            OpenWebConnection2.OptionKeyString = c.OptionKeyString
        '            OpenWebConnection2.OrganizationsCID = c.OrganizationsCID
        '            OpenWebConnection2.PeopleCID = c.PeopleCID
        '            OpenWebConnection2.RootPath = c.RootPath
        '            OpenWebConnection2.ServerPageDefault = c.ServerPageDefault
        '            OpenWebConnection2.TopicRulesCID = c.TopicRulesCID
        '            OpenWebConnection2.URLEncoder = c.URLEncoder
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function parseJSON(Source As String) As Object
        '    Set parseJSON = cmc.csv_parseJSON(Source)
        'End Function
        '        '
        '        '
        '        '
        '        Public Function ProcessReplacement(NameValueLines As Variant, Source As Variant) As String
        '            ProcessReplacement = cmc.csv_ProcessReplacement(NameValueLines, Source)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ProcessStyleSheet(Source As String) As String
        '            ProcessStyleSheet = cmc.csv_ProcessStyleSheet(Source)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ReadBake(Name As String) As String
        '            ReadBake = cmc.csv_ReadBake(Name)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ReadCache(Name As String) As String
        '            ReadCache = cmc.csv_ReadCache(Name)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ReadFile(PathFilename As String) As String
        '            ReadFile = cmc.csv_ReadFile(PathFilename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ReadVirtualFile(Filename As String) As String
        '            ReadVirtualFile = cmc.csv_ReadVirtualFile(Filename)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function reportAlarm(cause As String)
        '            reportAlarm = cmc.csv_reportAlarm(cause)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function reportError(cause As String, Optional Source As String, Optional ResumeNext As Boolean)
        '            reportError = cmc.csv_reportError(cause, Source, ResumeNext)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function reportWarning(Name As String, Description As String, generalKey As String, specificKey As String)
        '            reportWarning = cmc.csv_reportWarning(Name, Description, generalKey, specificKey)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function reportWarning2(Name As String, shortDescription As String, location As String, pageId As Long, Description As String, generalKey As String, specificKey As String)
        '            reportWarning2 = cmc.csv_reportWarning2(Name, shortDescription, location, pageId, Description, generalKey, specificKey)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function SendEmail(ByVal toAddress As String, ByVal fromAddress As String, ByVal SubjectMessage As String, ByVal BodyMessage As String, Optional ResultLogFilename As Variant, Optional Immediate As Variant, Optional HTML As Variant) As String
        '            SendEmail = cmc.csv_SendEmail(toAddress, fromAddress, SubjectMessage, BodyMessage, ResultLogFilename, Immediate, HTML)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function SendEmail2(ByVal toAddress As String, ByVal fromAddress As String, ByVal SubjectMessage As String, ByVal BodyMessage As String, BounceAddress As String, ReplyToAddress As String, Optional ResultLogFilename As Variant, Optional Immediate As Variant, Optional HTML As Variant) As String
        '            SendEmail2 = cmc.csv_SendEmail2(toAddress, fromAddress, SubjectMessage, BodyMessage, BounceAddress, ReplyToAddress, ResultLogFilename, Immediate, HTML)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function SendEmail3(ByVal toAddress As String, ByVal fromAddress As String, ByVal SubjectMessage As String, ByVal BodyMessage As String, BounceAddress As String, ReplyToAddress As String, ResultLogFilename As String, isImmediate As Boolean, isHTML As Boolean, emailIdOrZeroForLog As Long) As String
        '            SendEmail3 = cmc.csv_SendEmail3(toAddress, fromAddress, SubjectMessage, BodyMessage, BounceAddress, ReplyToAddress, ResultLogFilename, isImmediate, isHTML, emailIdOrZeroForLog)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function SendMemberEmail3(ToMemberID As Long, fromAddress As String, subject As String, Body As String, Immediate As Boolean, HTML As Boolean, emailIdOrZeroForLog As Long, template As String, EmailAllowLinkEID As Boolean) As String
        '            SendMemberEmail3 = cmc.csv_SendMemberEmail3(ToMemberID, fromAddress, subject, Body, Immediate, HTML, emailIdOrZeroForLog, template, EmailAllowLinkEID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function SendSystemEmail(emailName As String, AdditionalCopy As String, AdditionalMemberIDOrZero As Long) As String
        '            SendSystemEmail = cmc.csv_SendSystemEmail(emailName, AdditionalCopy, AdditionalMemberIDOrZero)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function SyncDataSources()
        '            SyncDataSources = cmc.csv_SyncDataSources()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function TextDeScramble(Copy As String) As String
        '            TextDeScramble = cmc.csv_TextDeScramble(Copy)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function TextScramble(Copy As String) As String
        '            TextScramble = cmc.csv_TextScramble(Copy)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function TimerTraceStart(TraceName As String) As Long
        '            TimerTraceStart = cmc.csv_TimerTraceStart(TraceName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function VerifyCDefField_ReturnID(ContentName As String, FieldName As String, Args As String, Delimiter As String) As Long
        '            VerifyCDefField_ReturnID = cmc.csv_VerifyCDefField_ReturnID(ContentName, FieldName, Args, Delimiter)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function VerifyDynamicMenu(MenuName As String) As Long
        '            VerifyDynamicMenu = cmc.csv_VerifyDynamicMenu(MenuName)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function VerifyNavigatorEntry4(ccGuid As String, NameSpace As String, EntryName As String, ContentName As String, LinkPage As String, SortOrder As String, adminOnly As Boolean, DeveloperOnly As Boolean, NewWindow As Boolean, Active As Boolean, MenuContentName As String, AddonName As String, NavIconType As String, NavIconTitle As String, InstalledByCollectionID As Long) As Long
        '            VerifyNavigatorEntry4 = cmc.csv_VerifyNavigatorEntry4(ccGuid, NameSpace, EntryName, ContentName, LinkPage, SortOrder, adminOnly, DeveloperOnly, NewWindow, Active, MenuContentName, AddonName, NavIconType, NavIconTitle, InstalledByCollectionID)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Sub AbortEdit(ContentName As String, recordId As Long, memberID As Long)
        '            Call cmc.csv_AbortEdit(ContentName, recordId, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub addHeadTags(headTags As String)
        '            Call cmc.csv_addHeadTags(headTags)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub addLinkAlias(linkAlias As String, pageId As Long, QueryStringSuffix As String)
        '            Call cmc.csv_addLinkAlias(linkAlias, pageId, QueryStringSuffix)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub addLinkAlias2(linkAlias As String, pageId As Long, QueryStringSuffix As String, OverRideDuplicate As Boolean, DupCausesWarning As Boolean, ByRef return_WarningMessage As String)
        '            Call cmc.csv_addLinkAlias2(linkAlias, pageId, QueryStringSuffix, OverRideDuplicate, DupCausesWarning, return_WarningMessage)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub AppendFile(Filename As String, fileContent As String)
        '            Call cmc.csv_AppendFile(Filename, fileContent)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub AppendVirtualFile(Filename As String, fileContent As String)
        '            Call cmc.csv_AppendVirtualFile(Filename, fileContent)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ApproveEdit(ContentName As String, recordId As Long, memberID As Long)
        '            Call cmc.csv_ApproveEdit(ContentName, recordId, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub BuildAddonOptionLists(OptionString_ForObjectCall As String, AddonOptionExpandedConstructor As String, AddonOptionConstructor As String, addonOptionString As String, InstanceID As String, IncludeSettingsBubbleOptions As Boolean)
        '            Call cmc.csv_BuildAddonOptionLists(OptionString_ForObjectCall, AddonOptionExpandedConstructor, AddonOptionConstructor, addonOptionString, InstanceID, IncludeSettingsBubbleOptions)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ClearAuthoringControl(ContentName As String, recordId As Long, AuthoringControl As Long, memberID As Long)
        '            Call cmc.csv_ClearAuthoringControl(ContentName, recordId, AuthoringControl, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ClearBake(ContentNameList As String)
        '            Call cmc.csv_ClearBake(ContentNameList)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub clearCache()
        '            Call cmc.csv_clearCache
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ClearEditLock(ContentName As String, recordId As Long, memberID As Long)
        '            Call cmc.csv_ClearEditLock(ContentName, recordId, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub clearPageContentCache()
        '            Call cmc.csv_ClearPageContentCache
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub clearPageTemplateCache()
        '            Call cmc.csv_clearPageTemplateCache
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub clearSiteSectionCache()
        '            Call cmc.csv_clearSiteSectionCache
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CloseConnection(Ignore As Long)
        '            Call cmc.csv_CloseConnection(Ignore)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CloseCS(CSPointer As Long, Optional AsyncSave As Boolean)
        '            Call cmc.csv_CloseCS(CSPointer, AsyncSave)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CloseStream()
        '            Call cmc.csv_CloseStream
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CopyCSRecord(CSSource As Long, CSDestination As Long)
        '            Call cmc.csv_CopyCSRecord(CSSource, CSDestination)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CopyFile(SourceFilename As String, destinationFilename As String)
        '            Call cmc.csv_CopyFile(SourceFilename, destinationFilename)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CopyVirtualFile(SourceFilename As String, destinationFilename As String)
        '            Call cmc.csv_CopyVirtualFile(SourceFilename, destinationFilename)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateAdminMenu(ParentName As String, EntryName As String, ContentName As String, LinkPage As String, SortOrder As String, Optional adminOnly As Boolean, Optional DeveloperOnly As Boolean, Optional NewWindow As Boolean)
        '            Call cmc.csv_CreateAdminMenu(ParentName, EntryName, ContentName, LinkPage, SortOrder, adminOnly, DeveloperOnly, NewWindow)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContent(Active As Boolean, DataSourcename As String, tableName As String, ContentName As String, Optional adminOnly As Variant, Optional DeveloperOnly As Variant, Optional AllowAdd As Variant, Optional AllowDelete As Variant, Optional ParentContentName As Variant, Optional DefaultSortMethod As Variant, Optional DropDownFieldList As Variant, Optional AllowWorkflowAuthoring As Variant, Optional AllowCalendarEvents As Variant, Optional AllowContentTracking As Variant, Optional AllowTopicRules As Variant, Optional AllowContentChildTool As Variant, Optional AllowMetaContent As Variant)
        '            Call cmc.csv_CreateContent(Active, DataSourcename, tableName, ContentName, adminOnly, DeveloperOnly, AllowAdd, AllowDelete, ParentContentName, DefaultSortMethod, DropDownFieldList, AllowWorkflowAuthoring, AllowCalendarEvents, AllowContentTracking, AllowTopicRules, AllowContentChildTool, AllowMetaContent)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContent2(Active As Boolean, DataSourcename As String, tableName As String, ContentName As String, Optional adminOnly As Variant, Optional DeveloperOnly As Variant, Optional AllowAdd As Variant, Optional AllowDelete As Variant, Optional ParentContentName As Variant, Optional DefaultSortMethod As Variant, Optional DropDownFieldList As Variant, Optional AllowWorkflowAuthoring As Variant, Optional AllowCalendarEvents As Variant, Optional AllowContentTracking As Variant, Optional AllowTopicRules As Variant, Optional AllowContentChildTool As Variant, Optional AllowMetaContent As Variant, Optional IconLink As String, Optional IconWidth As Long, Optional IconHeight As Long, Optional IconSprites As Long, Optional ccGuid As String)
        '            Call cmc.csv_CreateContent2(Active, DataSourcename, tableName, ContentName, adminOnly, DeveloperOnly, AllowAdd, AllowDelete, ParentContentName, DefaultSortMethod, DropDownFieldList, AllowWorkflowAuthoring, AllowCalendarEvents, AllowContentTracking, AllowTopicRules, AllowContentChildTool, AllowMetaContent, IconLink, IconWidth, IconHeight, IconSprites, ccGuid)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContent3(Active As Boolean, DataSourcename As String, tableName As String, ContentName As String, Optional adminOnly As Variant, Optional DeveloperOnly As Variant, Optional AllowAdd As Variant, Optional AllowDelete As Variant, Optional ParentContentName As Variant, Optional DefaultSortMethod As Variant, Optional DropDownFieldList As Variant, Optional AllowWorkflowAuthoring As Variant, Optional AllowCalendarEvents As Variant, Optional AllowContentTracking As Variant, Optional AllowTopicRules As Variant, Optional AllowContentChildTool As Variant, Optional AllowMetaContent As Variant, Optional IconLink As String, Optional IconWidth As Long, Optional IconHeight As Long, Optional IconSprites As Long, Optional ccGuid As String, Optional isBaseContent As Boolean)
        '            Call cmc.csv_CreateContent3(Active, DataSourcename, tableName, ContentName, adminOnly, DeveloperOnly, AllowAdd, AllowDelete, ParentContentName, DefaultSortMethod, DropDownFieldList, AllowWorkflowAuthoring, AllowCalendarEvents, AllowContentTracking, AllowTopicRules, AllowContentChildTool, AllowMetaContent, IconLink, IconWidth, IconHeight, IconSprites, ccGuid, isBaseContent)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContent4(Active As Boolean, DataSourcename As String, tableName As String, ContentName As String, Optional adminOnly As Variant, Optional DeveloperOnly As Variant, Optional AllowAdd As Variant, Optional AllowDelete As Variant, Optional ParentContentName As Variant, Optional DefaultSortMethod As Variant, Optional DropDownFieldList As Variant, Optional AllowWorkflowAuthoring As Variant, Optional AllowCalendarEvents As Variant, Optional AllowContentTracking As Variant, Optional AllowTopicRules As Variant, Optional AllowContentChildTool As Variant, Optional AllowMetaContent As Variant, Optional IconLink As String, Optional IconWidth As Long, Optional IconHeight As Long, Optional IconSprites As Long, Optional ccGuid As String, Optional isBaseContent As Boolean, Optional installedByCollectionGuid As String)
        '            Call cmc.csv_CreateContent4(Active, DataSourcename, tableName, ContentName, adminOnly, DeveloperOnly, AllowAdd, AllowDelete, ParentContentName, DefaultSortMethod, DropDownFieldList, AllowWorkflowAuthoring, AllowCalendarEvents, AllowContentTracking, AllowTopicRules, AllowContentChildTool, AllowMetaContent, IconLink, IconWidth, IconHeight, IconSprites, ccGuid, isBaseContent, installedByCollectionGuid)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContentChild(ChildContentName As String, ParentContentName As String, memberID As Long)
        '            Call cmc.csv_CreateContentChild(ChildContentName, ParentContentName, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContentField(Active As Boolean, ContentName As String, FieldName As String, fieldType As Long, Optional FieldSortOrder As Variant, Optional FieldAuthorable As Variant, Optional FieldCaption As Variant, Optional LookupContentName As Variant, Optional DefaultValue As Variant, Optional NotEditable As Variant, Optional AdminIndexColumn As Variant, Optional AdminIndexWidth As Variant, Optional AdminIndexSort As Variant, Optional RedirectContentName As Variant, Optional RedirectIDField As Variant, Optional RedirectPath As Variant, Optional HTMLContent As Variant, Optional UniqueName As Variant, Optional password As Variant, Optional adminOnly As Boolean, Optional DeveloperOnly As Boolean, Optional ReadOnly As Boolean, Optional FieldRequired As Boolean)
        '            Call cmc.csv_CreateContentField(Active, ContentName, FieldName, fieldType, FieldSortOrder, FieldAuthorable, FieldCaption, LookupContentName, DefaultValue, NotEditable, AdminIndexColumn, AdminIndexWidth, AdminIndexSort, RedirectContentName, RedirectIDField, RedirectPath, HTMLContent, UniqueName, password, adminOnly, DeveloperOnly, readOnly, FieldRequired)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContentField2(Active As Boolean, ContentName As String, FieldName As String, fieldType As Long, Optional FieldSortOrder As Variant, Optional FieldAuthorable As Variant, Optional FieldCaption As Variant, Optional LookupContentName As Variant, Optional DefaultValue As Variant, Optional NotEditable As Variant, Optional AdminIndexColumn As Variant, Optional AdminIndexWidth As Variant, Optional AdminIndexSort As Variant, Optional RedirectContentName As Variant, Optional RedirectIDField As Variant, Optional RedirectPath As Variant, Optional HTMLContent As Variant, Optional UniqueName As Variant, Optional password As Variant, Optional adminOnly As Boolean, Optional DeveloperOnly As Boolean, Optional ReadOnly As Boolean, Optional FieldRequired As Boolean, Optional IsBaseField As Boolean)
        '            Call cmc.csv_CreateContentField2(Active, ContentName, FieldName, fieldType, FieldSortOrder, FieldAuthorable, FieldCaption, LookupContentName, DefaultValue, NotEditable, AdminIndexColumn, AdminIndexWidth, AdminIndexSort, RedirectContentName, RedirectIDField, RedirectPath, HTMLContent, UniqueName, password, adminOnly, DeveloperOnly, readOnly, FieldRequired, IsBaseField)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContentFieldByTable(Active As Boolean, tableName As String, FieldName As String, fieldType As Long, Optional FieldSortOrder As Variant, Optional FieldAuthorable As Variant, Optional FieldCaption As Variant, Optional LookupContentName As Variant, Optional DefaultValue As Variant, Optional NotEditable As Variant, Optional AdminIndexColumn As Variant, Optional AdminIndexWidth As Variant, Optional AdminIndexSort As Variant, Optional RedirectContentName As Variant, Optional RedirectIDField As Variant, Optional RedirectPath As Variant, Optional HTMLContent As Variant, Optional UniqueName As Variant, Optional password As Variant, Optional adminOnly As Boolean, Optional DeveloperOnly As Boolean, Optional ReadOnly As Boolean, Optional FieldRequired As Boolean)
        '            Call cmc.csv_CreateContentFieldByTable(Active, tableName, FieldName, fieldType, FieldSortOrder, FieldAuthorable, FieldCaption, LookupContentName, DefaultValue, NotEditable, AdminIndexColumn, AdminIndexWidth, AdminIndexSort, RedirectContentName, RedirectIDField, RedirectPath, HTMLContent, UniqueName, password, adminOnly, DeveloperOnly, readOnly, FieldRequired)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContentFieldFromTableField(ContentName As String, FieldName As String, ADOFieldType As Long)
        '            Call cmc.csv_CreateContentFieldFromTableField(ContentName, FieldName, ADOFieldType)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContentFieldsFromSQLTable(DataSourcename As String, tableName As String)
        '            Call cmc.csv_CreateContentFieldsFromSQLTable(DataSourcename, tableName)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateContentFromSQLTable(DataSourcename As String, tableName As String, ContentName As String, memberID As Long)
        '            Call cmc.csv_CreateContentFromSQLTable(DataSourcename, tableName, ContentName, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateFileFolder(FolderPath As String)
        '            Call cmc.csv_CreateFileFolder(FolderPath)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateSQLIndex(DataSourcename As String, tableName As String, IndexName As String, FieldNames As String)
        '            Call cmc.csv_CreateSQLIndex(DataSourcename, tableName, IndexName, FieldNames)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateSQLTable(DataSourcename As String, tableName As String, Optional AllowAutoIncrement As Variant)
        '            Call cmc.csv_CreateSQLTable(DataSourcename, tableName, AllowAutoIncrement)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub CreateSQLTableField(DataSourcename As String, tableName As String, FieldName As String, fieldType As Long)
        '            Call cmc.csv_CreateSQLTableField(DataSourcename, tableName, FieldName, fieldType)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteContentRecord(ContentName As String, recordId As Long, Optional memberID As Long)
        '            Call cmc.csv_DeleteContentRecord(ContentName, recordId, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteContentRecords(ContentName As String, Criteria As String, Optional memberID As Long)
        '            Call cmc.csv_DeleteContentRecords(ContentName, Criteria, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteContentRules(ContentID As Long, recordId As Long)
        '            Call cmc.csv_DeleteContentRules(ContentID, recordId)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteContentTracking(ContentName As String, recordId As Long, Permanent As Boolean)
        '            Call cmc.csv_DeleteContentTracking(ContentName, recordId, Permanent)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteCSRecord(CSPointer As Long)
        '            Call cmc.csv_DeleteCSRecord(CSPointer)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteFile(PathFilename As String)
        '            Call cmc.csv_DeleteFile(PathFilename)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteSQLTableField(DataSourcename As String, tableName As String, FieldName As String)
        '            Call cmc.csv_DeleteSQLTableField(DataSourcename, tableName, FieldName)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteTable(DataSourcename As String, tableName As String)
        '            Call cmc.csv_DeleteTable(DataSourcename, tableName)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteTableField(DataSourcename As String, tableName As String, FieldName As String)
        '            Call cmc.csv_DeleteTableField(DataSourcename, tableName, FieldName)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteTableIndex(DataSourcename As String, tableName As String, IndexName As String)
        '            Call cmc.csv_DeleteTableIndex(DataSourcename, tableName, IndexName)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteTableRecord(DataSourcename As String, tableName As String, recordId As Long)
        '            Call cmc.csv_DeleteTableRecord(DataSourcename, tableName, recordId)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteTableRecordChunks(DataSourcename As String, tableName As String, Criteria As String, Optional ChunkSize As Long, Optional MaxChunkCount As Long)
        '            Call cmc.csv_DeleteTableRecordChunks(DataSourcename, tableName, Criteria, ChunkSize, MaxChunkCount)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteTableRecords(DataSourcename As String, tableName As String, Criteria As String)
        '            Call cmc.csv_DeleteTableRecords(DataSourcename, tableName, Criteria)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub DeleteVirtualFile(Filename As String)
        '            Call cmc.csv_DeleteVirtualFile(Filename)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub dispose()
        '            'Call cmc.dispose
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ExportCDef()
        '            Call cmc.csv_ExportCDef
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ExportCDef2(IncludeBaseFields As Boolean)
        '            Call cmc.csv_ExportCDef2(IncludeBaseFields)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ExportXML(Filename As String)
        '            Call cmc.csv_ExportXML(Filename)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ExportXML2(Filename As String, IncludeBaseFields As Boolean)
        '            Call cmc.csv_ExportXML2(Filename, IncludeBaseFields)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub FirstCSRecord(CSPointer As Long, Optional AsyncSave As Boolean)
        '            Call cmc.csv_FirstCSRecord(CSPointer, AsyncSave)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub GetAuthoringStatus(ContentName As String, recordId As Long, memberID As Long, IsSubmitted As Boolean, IsApproved As Boolean, SubmittedName As String, ApprovedName As String, IsInserted As Boolean, IsDeleted As Boolean, IsModified As Boolean, ModifiedName As String, ModifiedDate As Date, SubmittedDate As Date, ApprovedDate As Date)
        '            Call cmc.csv_GetAuthoringStatus(ContentName, recordId, memberID, IsSubmitted, IsApproved, SubmittedName, ApprovedName, IsInserted, IsDeleted, IsModified, ModifiedName, ModifiedDate, SubmittedDate, ApprovedDate)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ImportXML(Filename As String)
        '            Call cmc.csv_ImportXML(Filename)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub IncrementErrorCount()
        '            Call cmc.csv_IncrementErrorCount
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub InsertTableRecord(DataSourcename As String, tableName As String, SQLNameArray() As String, SQLValueArray() As String)
        '            Call cmc.csv_InsertTableRecord(DataSourcename, tableName, SQLNameArray(), SQLValueArray())
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub LoadContentDefinition(ContentID As Long)
        '            Call cmc.csv_LoadContentDefinition(ContentID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub LoadContentDefinitions()
        '            Call cmc.csv_LoadContentDefinitions
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub LoadContentEngine()
        '            Call cmc.csv_LoadContentEngine
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub LoadDataSources()
        '            Call cmc.csv_LoadDataSources
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub LoadResultCache(CSPointer As Long)
        '            Call cmc.csv_LoadResultCache(CSPointer)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub LoadSiteProperties()
        '            Call cmc.csv_LoadSiteProperties
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub LogActivity(Message As String)
        '            Call cmc.csv_LogActivity(Message)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub LogActivity2(Message As String, ByMemberID As Long, SubjectMemberID As Long, SubjectOrganizationID As Long)
        '            Call cmc.csv_LogActivity2(Message, ByMemberID, SubjectMemberID, SubjectOrganizationID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub LogActivity3(Message As String, ByMemberID As Long, SubjectMemberID As Long, SubjectOrganizationID As Long, Optional Link As String, Optional VisitorID As Long, Optional VisitID As Long)
        '            Call cmc.csv_LogActivity3(Message, ByMemberID, SubjectMemberID, SubjectOrganizationID, Link, VisitorID, VisitID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub NextCSRecord(CSPointer As Long, Optional AsyncSave As Boolean)
        '            Call cmc.csv_NextCSRecord(CSPointer, AsyncSave)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub OpenStream(Filename As String)
        '            Call cmc.csv_OpenStream(Filename)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub PublishEdit(ContentName As String, recordId As Long, memberID As Long)
        '            Call cmc.csv_PublishEdit(ContentName, recordId, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub renameFile(sourcePathFilename As String, destinationFilename As String)
        '            Call cmc.csv_renameFile(sourcePathFilename, destinationFilename)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ReportClientError(ErrorDescription As String)
        '            Call cmc.csv_ReportClientError(ErrorDescription)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub RequestTask(Command As String, SQL As String, ExportName As String, Filename As String, RequestedByMemberID As Long)
        '            Call cmc.csv_RequestTask(Command, SQL, ExportName, Filename, RequestedByMemberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ResizeImage(SrcFilename As String, DstFilename As String, Width As Long, Height As Long)
        '            Call cmc.csv_ResizeImage(SrcFilename, DstFilename, Width, Height)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub ResizeImage2(SrcFilename As String, DstFilename As String, Width As Long, Height As Long, Algorithm As SfImageResizeAlgorithms)
        '            Call cmc.csv_ResizeImage2(SrcFilename, DstFilename, Width, Height, Algorithm)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub RollBackCS(CSPointer As Long)
        '            Call cmc.csv_RollBackCS(CSPointer)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SaveBake(Name As String, Value As String, Optional ContentNameList As String, Optional DateExpires As Date)
        '            Call cmc.csv_SaveBake(Name, Value, ContentNameList, DateExpires)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SaveCache(Name As String, Value As String)
        '            Call cmc.csv_SaveCache(Name, Value)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SaveCS(CSPointer As Long, Optional AsyncSave As Boolean, Optional BlockClearBake As Boolean)
        '            Call cmc.csv_SaveCS(CSPointer, AsyncSave, BlockClearBake)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SaveCSRecord(CSPointer As Long, Optional AsyncSave As Boolean, Optional BlockClearBake As Boolean)
        '            Call cmc.csv_SaveCSRecord(CSPointer, AsyncSave, BlockClearBake)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SaveFile(Filename As String, fileContent As String)
        '            Call cmc.csv_SaveFile(Filename, fileContent)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SavePersistentVariant(Var As Variant, Key As String)
        '            Call cmc.csv_SavePersistentVariant(Var, Key)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SaveVirtualFile(Filename As String, fileContent As String)
        '            Call cmc.csv_SaveVirtualFile(Filename, fileContent)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SetAuthoringControl(ContentName As String, recordId As Long, AuthoringControl As Long, memberID As Long)
        '            Call cmc.csv_SetAuthoringControl(ContentName, recordId, AuthoringControl, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SetCS(CSPointer As Long, FieldName As String, FieldValue As Variant)
        '            Call cmc.csv_SetCS(CSPointer, FieldName, FieldValue)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SetCSField(CSPointer As Long, FieldName As String, FieldValue As Variant)
        '            Call cmc.csv_SetCSField(CSPointer, FieldName, FieldValue)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SetCSRecordDefaults(CS As Long)
        '            Call cmc.csv_SetCSRecordDefaults(CS)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SetCSTextFile(CSPointer As Long, FieldName As String, Copy As String, ContentName As String)
        '            Call cmc.csv_SetCSTextFile(CSPointer, FieldName, Copy, ContentName)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SetEditLock(ContentName As String, recordId As Long, memberID As Long)
        '            Call cmc.csv_SetEditLock(ContentName, recordId, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SetProperty(typeId As Long, KeyID As Long, Name As String, Value As String, memberID As Long, Optional ForceInsert As Boolean)
        '            Call cmc.csv_SetProperty(typeId, KeyID, Name, Value, memberID, ForceInsert)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SetSiteProperty(Name As String, Value As String, memberID As Long, Optional AllowAdminAccess As Boolean)
        '            Call cmc.csv_SetSiteProperty(Name, Value, memberID, AllowAdminAccess)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub SubmitEdit(ContentName As String, recordId As Long, memberID As Long)
        '            Call cmc.csv_SubmitEdit(ContentName, recordId, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub TestPoint(Message)
        '            Call cmc.csv_TestPoint(Message)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub TrackContentSet(CSPointer As Long, pathPage As String, memberID As Long)
        '            Call cmc.csv_TrackContentSet(CSPointer, pathPage, memberID)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub UpdateTableRecord(DataSourcename As String, tableName As String, Criteria As String, SQLNameArray() As String, SQLValueArray() As String)
        '            Call cmc.csv_UpdateTableRecord(DataSourcename, tableName, Criteria, SQLNameArray(), SQLValueArray())
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyAdminMenu(ParentName As String, EntryName As String, ContentName As String, LinkPage As String, SortOrder As String, Optional adminOnly As Boolean, Optional DeveloperOnly As Boolean, Optional NewWindow As Boolean, Optional Active As Variant)
        '            Call cmc.csv_VerifyAdminMenu(ParentName, EntryName, ContentName, LinkPage, SortOrder, adminOnly, DeveloperOnly, NewWindow, Active)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyAggregateFunction(Name As String, Link As String, ObjectProgramID As String, ArgumentList As String, SortOrder As String)
        '            Call cmc.csv_VerifyAggregateFunction(Name, Link, ObjectProgramID, ArgumentList, SortOrder)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyAggregateObject(Name As String, ObjectProgramID As String, ArgumentList As String, SortOrder As String)
        '            Call cmc.csv_VerifyAggregateObject(Name, ObjectProgramID, ArgumentList, SortOrder)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyAggregateReplacement(Name As String, Copy As String, SortOrder As String)
        '            Call cmc.csv_VerifyAggregateReplacement(Name, Copy, SortOrder)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyAggregateReplacement2(Name As String, Copy As String, ArgumentList As String, SortOrder As String)
        '            Call cmc.csv_VerifyAggregateReplacement2(Name, Copy, ArgumentList, SortOrder)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyAggregateScript(Name As String, Link As String, ArgumentList As String, SortOrder As String)
        '            Call cmc.csv_VerifyAggregateScript(Name, Link, ArgumentList, SortOrder)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyCDefField(ContentName As String, FieldName As String, Args As String, Delimiter As String)
        '            Call cmc.csv_VerifyCDefField(ContentName, FieldName, Args, Delimiter)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyMenuEntry(ParentName As String, EntryName As String, ContentName As String, LinkPage As String, SortOrder As String, adminOnly As Boolean, DeveloperOnly As Boolean, NewWindow As Boolean, Active As Boolean, MenuContentName As String, AddonName As String)
        '            Call cmc.csv_VerifyMenuEntry(ParentName, EntryName, ContentName, LinkPage, SortOrder, adminOnly, DeveloperOnly, NewWindow, Active, MenuContentName, AddonName)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyNavigatorEntry(NameSpace As String, EntryName As String, ContentName As String, LinkPage As String, SortOrder As String, adminOnly As Boolean, DeveloperOnly As Boolean, NewWindow As Boolean, Active As Boolean, MenuContentName As String, AddonName As String)
        '            Call cmc.csv_VerifyNavigatorEntry(NameSpace, EntryName, ContentName, LinkPage, SortOrder, adminOnly, DeveloperOnly, NewWindow, Active, MenuContentName, AddonName)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyNavigatorEntry2(ccGuid As String, NameSpace As String, EntryName As String, ContentName As String, LinkPage As String, SortOrder As String, adminOnly As Boolean, DeveloperOnly As Boolean, NewWindow As Boolean, Active As Boolean, MenuContentName As String, AddonName As String)
        '            Call cmc.csv_VerifyNavigatorEntry2(ccGuid, NameSpace, EntryName, ContentName, LinkPage, SortOrder, adminOnly, DeveloperOnly, NewWindow, Active, MenuContentName, AddonName)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub VerifyNavigatorEntry3(ccGuid As String, NameSpace As String, EntryName As String, ContentName As String, LinkPage As String, SortOrder As String, adminOnly As Boolean, DeveloperOnly As Boolean, NewWindow As Boolean, Active As Boolean, MenuContentName As String, AddonName As String, NavIconType As String, NavIconTitle As String)
        '            Call cmc.csv_VerifyNavigatorEntry3(ccGuid, NameSpace, EntryName, ContentName, LinkPage, SortOrder, adminOnly, DeveloperOnly, NewWindow, Active, MenuContentName, AddonName, NavIconType, NavIconTitle)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub WriteStream(Message)
        '            Call cmc.csv_WriteStream(Message)
        '        End Sub

    End Class

End Namespace




