Option Explicit On
Option Strict On



Namespace BTA_OSG
    Public Class AppStartup
        Public Shared ConnectionFactory As IDbConnectionFactory
        Public Shared Settings As AppSettings
        
        ' Repositories
        Public Shared UserRepo As UserRepository
        Public Shared DocumentRepo As DocumentRepository
        Public Shared DirectiveRepo As DirectiveRepository
        Public Shared RoutingRepo As RoutingRepository
        Public Shared StorageRepo As StorageRepository
        Public Shared AuditRepo As AuditRepository
        Public Shared SequenceRepo As SequenceRepository
        Public Shared ReferenceDataRepo As ReferenceDataRepository
        
        ' Services
        Public Shared AuthService As AuthenticationService
        Public Shared DocService As DocumentService
        Public Shared DirectiveService As DirectiveService
        Public Shared RoutingService As RoutingService  
        Public Shared StorageService As StorageService
        Public Shared SearchService As SearchService
        Public Shared PdfService As PdfLinkService
        Public Shared UserService As UserService
        Public Shared CardService As RfidCardService
        Public Shared AuditService As AuditService
        Public Shared PortalBridgeClient As IPortalBridge
        
        Public Shared Sub Initialize()
            Settings = AppSettings.Instance
            ConnectionFactory = New SqlConnectionFactory(Settings.DatabaseSettings.ConnectionString)
            
            ' Wire repos
            UserRepo = New UserRepository(ConnectionFactory)
            DocumentRepo = New DocumentRepository(ConnectionFactory)
            DirectiveRepo = New DirectiveRepository(ConnectionFactory)
            RoutingRepo = New RoutingRepository(ConnectionFactory)
            StorageRepo = New StorageRepository(ConnectionFactory)
            AuditRepo = New AuditRepository(ConnectionFactory)
            SequenceRepo = New SequenceRepository(ConnectionFactory)
            ReferenceDataRepo = New ReferenceDataRepository(ConnectionFactory)
            
            ' Wire services
            AuditService = New AuditService(AuditRepo)
            PortalBridgeClient = New PortalBridge(Settings.PortalSettings, auditService:=AuditService)
            AuthService = New AuthenticationService(UserRepo, AuditRepo, Settings.RfidSettings)
            DocService = New DocumentService(DocumentRepo, SequenceRepo, ReferenceDataRepo, AuditService, PortalBridgeClient, RoutingRepo, StorageRepo)
            DirectiveService = New DirectiveService(DirectiveRepo, DocumentRepo, ReferenceDataRepo, AuditService, PortalBridgeClient)
            RoutingService = New RoutingService(RoutingRepo, DocumentRepo, ReferenceDataRepo, AuditService, PortalBridgeClient)
            StorageService = New StorageService(StorageRepo, DocumentRepo, AuditService)
            SearchService = New SearchService(DocumentRepo)
            PdfService = New PdfLinkService(Settings.PdfLinkSettings, AuditService)
            UserService = New UserService(UserRepo, AuditService)
            CardService = New RfidCardService(ConnectionFactory, AuditService)
        End Sub
    End Class
End Namespace
