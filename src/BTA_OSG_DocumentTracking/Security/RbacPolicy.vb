Option Explicit On
Option Strict On

Namespace BTA_OSG
    Public Class RbacPolicy
        ' App codes (spec) + DB seed aliases (007) both accepted - seed uses DOC_* variants
        Public Const DOCUMENT_VIEW_ALL As String = "DOCUMENT_VIEW_ALL"
        Public Const DOC_VIEW As String = "DOC_VIEW"
        Public Const DOCUMENT_VIEW_ASSIGNED As String = "DOCUMENT_VIEW_ASSIGNED"
        Public Const DOCUMENT_CREATE As String = "DOCUMENT_CREATE"
        Public Const DOC_CREATE As String = "DOC_CREATE"
        Public Const DOCUMENT_UPDATE As String = "DOCUMENT_UPDATE"
        Public Const DOC_EDIT As String = "DOC_EDIT"
        Public Const DOCUMENT_DELETE_OR_DISABLE As String = "DOCUMENT_DELETE_OR_DISABLE"
        Public Const DOC_DELETE As String = "DOC_DELETE"
        Public Const DIRECTIVE_ISSUE As String = "DIRECTIVE_ISSUE"
        Public Const DOC_DIRECTIVE As String = "DOC_DIRECTIVE"
        Public Const ROUTING_CREATE As String = "ROUTING_CREATE"
        Public Const DOC_ROUTE As String = "DOC_ROUTE"
        Public Const STORAGE_UPDATE As String = "STORAGE_UPDATE"
        Public Const STORAGE_MANAGE As String = "STORAGE_MANAGE"
        Public Const PDF_OPEN As String = "PDF_OPEN"
        Public Const AUDIT_VIEW As String = "AUDIT_VIEW"
        Public Const USER_MANAGE As String = "USER_MANAGE"
        Public Const RFID_MANAGE As String = "RFID_MANAGE"
        Public Const ROLE_MANAGE As String = "ROLE_MANAGE"
        Public Const REFERENCE_MANAGE As String = "REFERENCE_MANAGE"
        Public Const SYS_CONFIG As String = "SYS_CONFIG"
        Public Const REPORT_EXPORT As String = "REPORT_EXPORT"
        Public Const REP_VIEW As String = "REP_VIEW"
        Public Const DASHBOARD_VIEW As String = "DASHBOARD_VIEW"

        ' Standard System Role Codes (matches db/scripts/007 seed)
        Public Const ROLE_SG As String = "SG"
        Public Const ROLE_SYSADMIN As String = "SYSADMIN"
        Public Const ROLE_OSG_CHIEF As String = "OSG_CHIEF"
        Public Const ROLE_RECORDS As String = "RECORDS"
        Public Const ROLE_SECRETARIAT As String = "SECRETARIAT"
        Public Const ROLE_LEGISLATIVE As String = "LEGISLATIVE"
        Public Const ROLE_FINANCE As String = "FINANCE"
        Public Const ROLE_TRAVEL As String = "TRAVEL"
        Public Const ROLE_ADMIN_STAFF As String = "ADMIN_STAFF"
    End Class
End Namespace
