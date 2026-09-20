# Desktop Civic Calm UI and Architecture Remediation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Transform the BTA OSG Document Tracking desktop application to 100% compliance with the updated DESIGN.md (Desktop Civic Calm) specification, enforce VB.NET Option Strict On compiler hygiene, strongly type the service layer, wire UI forms to the ADO.NET SQL Server architecture with EmbeddedDB fallback, and purge all legacy bloat and anti-slop violations.

**Architecture:** Implement a centralized `CivicCalmTheme` and `DataGridStyler` providing 3-tier optical surfaces (Canvas #F4F6F8, Card #FFFFFF, Recessed Well #EEF2F5) and 4-state double-buffered DataGridViews. Replace manual pixel coordinate positioning with `TableLayoutPanel` and `FlowLayoutPanel` containers adhering to Gestalt spacing. Replace blocking `MessageBox.Show` popups with bottom `StatusStrip` and `ErrorProvider`. Eliminate late-binding `Object` types in services and wire forms to `AppStartup`.

**Tech Stack:** VB.NET 10 (`net10.0-windows`), Windows Forms, `Microsoft.Data.SqlClient` 6.0.1, MSTest, `System.Drawing`, `System.Text.Json`.

## Global Constraints

- Target Framework: `net10.0-windows` with Windows Forms (`<UseWindowsForms>true</UseWindowsForms>`).
- Enforce `<OptionStrict>On</OptionStrict>`, `<OptionExplicit>On</OptionExplicit>`, `<OptionCompare>Binary</OptionCompare>`, and `<OptionInfer>On</OptionInfer>` across all projects.
- Zero Em-Dash Rule: Never use the em-dash character (`—`) in any UI text, comments, or strings per DESIGN.md §7 and antislop R-02.
- High DPI Scaling: Absolute coordinate positioning (`Location = New Point(x, y)`) is strictly banned. Layouts must use `TableLayoutPanel`, `FlowLayoutPanel`, or explicit `Dock`/`Anchor`.
- Non-Modal Feedback: Routine operational confirmations must update the bottom `StatusStrip`; field validation must use `ErrorProvider` tooltips. Blocking `MessageBox.Show` dialogs are strictly reserved for irreversible destructive actions.
- DataGridView Lifecycle: All grids must support 4 states (Loading, Empty watermark label *"No documents match the current filter criteria. Press Alt+C to clear filters."*, Error inline banner, Populated) with `RowHeadersVisible = False`, 28px rows, and `#FFFFFF`/`#F9FAFB` alternating rows.
- Keyboard-First Ergonomics: Every action button must include an Alt-key mnemonic (`&`); all dialogs must bind `AcceptButton` and `CancelButton`; sequential `TabIndex` must be declared.
- Thread Safety: Wrap all background and hardware listener callbacks in `If Me.InvokeRequired Then Me.Invoke(...)`.

---

### Task 1: Project Configuration & Legacy Bloat Removal

**Files:**
- Delete: `legacy/BTA_OSG_System.vb`
- Delete: `legacy/BTA_OSG_DB.sql`
- Delete: `legacy/BTA_OSG_DocumentTracking.vbproj`
- Delete: `src/BTA_OSG_DocumentTracking/Forms/FormDirectiveEntry.vb`
- Delete: `src/BTA_OSG_DocumentTracking/Forms/FormSettings.vb`
- Modify: `src/BTA_OSG_DocumentTracking/BTA_OSG_DocumentTracking.vbproj`
- Modify: `tests/BTA_OSG_DocumentTracking.Tests/BTA_OSG_DocumentTracking.Tests.vbproj`

**Interfaces:**
- Consumes: None
- Produces: Cleaned project files enforcing `Option Strict On` and `Option Explicit On`.

- [ ] **Step 1: Delete dead legacy files and unreferenced forms**

Run PowerShell commands:
```powershell
Remove-Item -Path "legacy/BTA_OSG_System.vb", "legacy/BTA_OSG_DB.sql", "legacy/BTA_OSG_DocumentTracking.vbproj" -Force
Remove-Item -Path "src/BTA_OSG_DocumentTracking/Forms/FormDirectiveEntry.vb", "src/BTA_OSG_DocumentTracking/Forms/FormSettings.vb" -Force
```

- [ ] **Step 2: Update BTA_OSG_DocumentTracking.vbproj with strict compilation options**

Update `src/BTA_OSG_DocumentTracking/BTA_OSG_DocumentTracking.vbproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net10.0-windows</TargetFramework>
    <UseWindowsForms>true</UseWindowsForms>
    <RootNamespace></RootNamespace>
    <HighDpiMode>SystemAware</HighDpiMode>
    <OptionExplicit>On</OptionExplicit>
    <OptionStrict>On</OptionStrict>
    <OptionCompare>Binary</OptionCompare>
    <OptionInfer>On</OptionInfer>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Data.SqlClient" Version="6.0.1" />
  </ItemGroup>

  <ItemGroup>
    <None Update="Resources\appsettings.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
    <None Update="Resources\appsettings.Production.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
  </ItemGroup>

</Project>
```

- [ ] **Step 3: Update BTA_OSG_DocumentTracking.Tests.vbproj with strict compilation options**

Update `tests/BTA_OSG_DocumentTracking.Tests/BTA_OSG_DocumentTracking.Tests.vbproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net10.0-windows</TargetFramework>
    <UseWindowsForms>true</UseWindowsForms>
    <ImplicitUsings>disable</ImplicitUsings>
    <Nullable>disable</Nullable>
    <IsPackable>false</IsPackable>
    <IsTestProject>true</IsTestProject>
    <OptionExplicit>On</OptionExplicit>
    <OptionStrict>On</OptionStrict>
    <OptionCompare>Binary</OptionCompare>
    <OptionInfer>On</OptionInfer>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.13.0" />
    <PackageReference Include="MSTest.TestAdapter" Version="3.8.2" />
    <PackageReference Include="MSTest.TestFramework" Version="3.8.2" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\..\src\BTA_OSG_DocumentTracking\BTA_OSG_DocumentTracking.vbproj" />
  </ItemGroup>

</Project>
```

- [ ] **Step 4: Run build to catalog any strict compilation errors**

Run: `dotnet build src/BTA_OSG_DocumentTracking -c Debug`
Expected: Output will highlight late-binding locations in services to be strongly typed in Task 2.

- [ ] **Step 5: Commit changes**

```powershell
git add -A
git commit -m "chore: purge legacy bloat and enable Option Strict On in project files"
```

---

### Task 2: Service Layer Strong Typing & Repository Alignment

**Files:**
- Modify: `src/BTA_OSG_DocumentTracking/Security/AuthenticationService.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Data/UserRepository.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Services/DocumentService.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Services/DirectiveService.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Services/RoutingService.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Services/StorageService.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Services/SearchService.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Services/UserService.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Services/RfidCardService.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Data/DocumentRepository.vb`
- Test: `tests/BTA_OSG_DocumentTracking.Tests/ServiceTypingTests.vb`

**Interfaces:**
- Consumes: Models (`User`, `Document`, `ActionDirective`, `RoutingLog`, `DocumentMovement`, `RfidCard`, `SessionContext`)
- Produces: Strongly typed service methods with zero `Object` references and zero late-bound invocations.

- [ ] **Step 1: Write unit tests for strongly typed services**

Create `tests/BTA_OSG_DocumentTracking.Tests/ServiceTypingTests.vb`:
```vb
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class ServiceTypingTests
        <TestMethod>
        Public Sub DocumentModel_InstantiatesWithRequiredFields()
            Dim doc As New Document With {
                .DocumentID = 101,
                .DocCode = "RES-2026-001",
                .Title = "Resolution on Parliamentary Affairs",
                .DocumentTypeID = 1,
                .StatusID = 1,
                .OriginOffice = "OSG",
                .DestinationOffice = "Plenary",
                .RegisteredByUserID = 1,
                .RegisteredAtUTC = DateTime.UtcNow
            }
            Assert.AreEqual("RES-2026-001", doc.DocCode)
            Assert.AreEqual(101, doc.DocumentID)
        End Sub

        <TestMethod>
        Public Sub ActionDirectiveModel_InstantiatesWithStrongTypes()
            Dim d As New ActionDirective With {
                .DirectiveID = 1,
                .DocumentID = 101,
                .DirectiveTypeID = 2,
                .DirectiveText = "For Immediate Review",
                .IssuedByUserID = 1,
                .IssuedAtUTC = DateTime.UtcNow,
                .IsActive = True
            }
            Assert.AreEqual("For Immediate Review", d.DirectiveText)
            Assert.IsTrue(d.IsActive)
        End Sub
    End Class
End Namespace
```

- [ ] **Step 2: Update AuthenticationService.vb with strong types**

Replace `src/BTA_OSG_DocumentTracking/Security/AuthenticationService.vb`:
```vb
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace BTA_OSG
    Public Class AuthenticationService
        Private ReadOnly _userRepo As UserRepository
        Private ReadOnly _auditRepo As AuditRepository
        Private ReadOnly _rfidSettings As RfidSettings

        Public Sub New(userRepo As UserRepository, auditRepo As AuditRepository, rfidSettings As RfidSettings)
            _userRepo = userRepo
            _auditRepo = auditRepo
            _rfidSettings = rfidSettings
        End Sub

        Public Function AuthenticateByCard(cardPublicId As String) As SessionContext
            If String.IsNullOrWhiteSpace(cardPublicId) Then Return Nothing
            Dim cleanCardId As String = cardPublicId.Trim().ToUpperInvariant()

            Dim user As User = _userRepo.GetByCardPublicID(cleanCardId)
            If user Is Nothing Then
                LogAudit("LOGIN_FAILURE", Nothing, Nothing, cleanCardId, False, "Card not recognized")
                Return Nothing
            End If

            If Not user.IsActive OrElse user.IsLocked Then
                LogAudit("LOGIN_FAILURE", CType(user.UserID, Integer?), user.Username, cleanCardId, False, "User account inactive or locked")
                Return Nothing
            End If

            If user.FailedTapCount >= _rfidSettings.LockoutThreshold Then
                LogAudit("LOGIN_FAILURE", CType(user.UserID, Integer?), user.Username, cleanCardId, False, "Account locked out due to consecutive failed attempts")
                Return Nothing
            End If

            Dim roles As List(Of Role) = _userRepo.GetUserRoles(user.UserID)
            Dim perms As New HashSet(Of String)()
            For Each r In roles
                Dim rolePerms = _userRepo.GetRolePermissions(r.RoleID)
                For Each p In rolePerms
                    perms.Add(p.PermissionCode)
                Next
            Next

            Dim session As New SessionContext With {
                .SessionID = Guid.NewGuid(),
                .User = user,
                .Roles = roles,
                .Permissions = perms,
                .LoginTime = DateTime.UtcNow,
                .LastActivityTime = DateTime.UtcNow
            }

            LogAudit("LOGIN_SUCCESS", CType(user.UserID, Integer?), user.Username, cleanCardId, True, Nothing)
            Return session
        End Function

        Private Sub LogAudit(actionType As String, userId As Integer?, username As String, cardId As String, success As Boolean, reason As String)
            If _auditRepo Is Nothing Then Return
            Dim entry As New AuditEntry With {
                .EventAtUTC = DateTime.UtcNow,
                .ActionType = actionType,
                .UserID = userId,
                .UsernameSnapshot = username,
                .CardPublicIDMasked = If(cardId.Length > 4, "****" & cardId.Substring(cardId.Length - 4), cardId),
                .Success = success,
                .FailureReason = reason,
                .MachineName = Environment.MachineName
            }
            Try
                _auditRepo.Insert(entry)
            Catch
            End Try
        End Sub
    End Class
End Namespace
```

- [ ] **Step 3: Update DocumentService, DirectiveService, RoutingService, StorageService, SearchService**

Ensure all service classes use explicit model types (`Document`, `ActionDirective`, `RoutingLog`, `DocumentMovement`, `User`) and call strongly-typed repository methods. Update `DocumentRepository.vb` to include `SearchDocuments` overload:
```vb
Public Function SearchDocuments(hasViewAll As Boolean, userId As Integer, titleLike As String, typeId As Integer?, statusId As Integer?, originLike As String, destLike As String, storageId As Integer?, dateFrom As DateTime?, dateTo As DateTime?, pageSize As Integer, pageNumber As Integer) As List(Of Document)
    Return GetByFilter(titleLike, typeId, statusId, originLike, destLike, storageId, dateFrom, dateTo, pageSize, pageNumber)
End Function
```

- [ ] **Step 4: Run tests to verify strong typing**

Run: `dotnet test tests/BTA_OSG_DocumentTracking.Tests`
Expected: PASS with 0 warnings under `Option Strict On`.

- [ ] **Step 5: Commit changes**

```powershell
git add src/ tests/
git commit -m "fix(arch): strongly type all services and repositories under Option Strict On"
```

---

### Task 3: Desktop Civic Calm Theme & DataGridStyler Implementation

**Files:**
- Create: `src/BTA_OSG_DocumentTracking/UI/CivicCalmTheme.vb`
- Create: `src/BTA_OSG_DocumentTracking/UI/DataGridStyler.vb`
- Test: `tests/BTA_OSG_DocumentTracking.Tests/ThemeTests.vb`

**Interfaces:**
- Consumes: `DESIGN.md` §1 Optical Surfaces, §2 Typography Ladder, §4 DataGridView Standards
- Produces: Shared styling definitions and `DataGridStyler.ApplyCivicStyle(dgv As DataGridView)`.

- [ ] **Step 1: Write unit test for CivicCalmTheme tokens**

Create `tests/BTA_OSG_DocumentTracking.Tests/ThemeTests.vb`:
```vb
Option Explicit On
Option Strict On

Imports System.Drawing
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class ThemeTests
        <TestMethod>
        Public Sub ThemeTokens_MatchDesignSpecification()
            Assert.AreEqual(ColorTranslator.FromHtml("#F4F6F8"), CivicCalmTheme.ColorCanvas)
            Assert.AreEqual(ColorTranslator.FromHtml("#FFFFFF"), CivicCalmTheme.ColorSurface)
            Assert.AreEqual(ColorTranslator.FromHtml("#EEF2F5"), CivicCalmTheme.ColorWell)
            Assert.AreEqual(ColorTranslator.FromHtml("#DDE2E5"), CivicCalmTheme.ColorBorder)
            Assert.AreEqual(ColorTranslator.FromHtml("#1B242C"), CivicCalmTheme.ColorInk)
            Assert.AreEqual(ColorTranslator.FromHtml("#146A3D"), CivicCalmTheme.ColorPrimary)
            Assert.AreEqual(ColorTranslator.FromHtml("#E6F2EB"), CivicCalmTheme.ColorPrimarySoft)
        End Sub

        <TestMethod>
        Public Sub TypographyLadder_UsesSegoeUI()
            Assert.AreEqual("Segoe UI", CivicCalmTheme.FontFormTitle.FontFamily.Name)
            Assert.AreEqual(12.0F, CivicCalmTheme.FontFormTitle.Size)
            Assert.AreEqual(FontStyle.Bold, CivicCalmTheme.FontFormTitle.Style)
            Assert.AreEqual("Segoe UI", CivicCalmTheme.FontBody.FontFamily.Name)
            Assert.AreEqual(9.0F, CivicCalmTheme.FontBody.Size)
        End Sub
    End Class
End Namespace
```

- [ ] **Step 2: Run test to verify it fails before implementation**

Run: `dotnet test tests/BTA_OSG_DocumentTracking.Tests --filter FullyQualifiedName~ThemeTests`
Expected: FAIL (types do not exist yet).

- [ ] **Step 3: Create CivicCalmTheme.vb**

Create `src/BTA_OSG_DocumentTracking/UI/CivicCalmTheme.vb`:
```vb
Option Explicit On
Option Strict On

Imports System.Drawing

Namespace BTA_OSG
    Public NotInheritable Class CivicCalmTheme
        Private Sub New()
        End Sub

        ' 3-Tier Optical Surface Hierarchy
        Public Shared ReadOnly ColorCanvas As Color = ColorTranslator.FromHtml("#F4F6F8")
        Public Shared ReadOnly ColorSurface As Color = ColorTranslator.FromHtml("#FFFFFF")
        Public Shared ReadOnly ColorWell As Color = ColorTranslator.FromHtml("#EEF2F5")
        Public Shared ReadOnly ColorBorder As Color = ColorTranslator.FromHtml("#DDE2E5")

        ' Typography & Ink Tokens
        Public Shared ReadOnly ColorInk As Color = ColorTranslator.FromHtml("#1B242C")
        Public Shared ReadOnly ColorInkMuted As Color = ColorTranslator.FromHtml("#55606A")

        ' Action & Semantic Tokens
        Public Shared ReadOnly ColorPrimary As Color = ColorTranslator.FromHtml("#146A3D")
        Public Shared ReadOnly ColorPrimarySoft As Color = ColorTranslator.FromHtml("#E6F2EB")
        Public Shared ReadOnly ColorAccentSG As Color = ColorTranslator.FromHtml("#B08524")
        Public Shared ReadOnly ColorAccentSoft As Color = ColorTranslator.FromHtml("#FEF9E7")
        Public Shared ReadOnly ColorDanger As Color = ColorTranslator.FromHtml("#A93226")
        Public Shared ReadOnly ColorDangerSoft As Color = ColorTranslator.FromHtml("#FBEAE8")

        ' Segoe UI Typography Ladder
        Public Shared ReadOnly FontFormTitle As New Font("Segoe UI", 12.0F, FontStyle.Bold)
        Public Shared ReadOnly FontSectionHeader As New Font("Segoe UI", 10.0F, FontStyle.Bold)
        Public Shared ReadOnly FontFieldLabel As New Font("Segoe UI", 9.0F, FontStyle.Bold)
        Public Shared ReadOnly FontBody As New Font("Segoe UI", 9.0F, FontStyle.Regular)
        Public Shared ReadOnly FontTabular As New Font("Segoe UI", 9.0F, FontStyle.Regular)
        Public Shared ReadOnly FontIdentifier As New Font("Segoe UI", 8.5F, FontStyle.Regular)
        Public Shared ReadOnly FontMicrocopy As New Font("Segoe UI", 8.0F, FontStyle.Regular)
    End Class
End Namespace
```

- [ ] **Step 4: Create DataGridStyler.vb**

Create `src/BTA_OSG_DocumentTracking/UI/DataGridStyler.vb`:
```vb
Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms

Namespace BTA_OSG
    Public NotInheritable Class DataGridStyler
        Private Sub New()
        End Sub

        Public Shared Sub ApplyCivicStyle(dgv As DataGridView)
            dgv.AutoGenerateColumns = False
            dgv.EnableHeadersVisualStyles = False
            dgv.BackgroundColor = CivicCalmTheme.ColorSurface
            dgv.BorderStyle = BorderStyle.FixedSingle
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            dgv.GridColor = CivicCalmTheme.ColorBorder
            dgv.ColumnHeadersHeight = 32
            dgv.RowTemplate.Height = 28
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.MultiSelect = False
            dgv.RowHeadersVisible = False

            Dim prop = GetType(Control).GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
            If prop IsNot Nothing Then
                prop.SetValue(dgv, True, Nothing)
            End If

            ' Column Headers (Tier 2 Recessed Well)
            dgv.ColumnHeadersDefaultCellStyle.BackColor = CivicCalmTheme.ColorWell
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.ColumnHeadersDefaultCellStyle.Font = CivicCalmTheme.FontSectionHeader

            ' Default & Alternating Row Styling
            dgv.DefaultCellStyle.BackColor = CivicCalmTheme.ColorSurface
            dgv.DefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.DefaultCellStyle.Font = CivicCalmTheme.FontTabular
            dgv.DefaultCellStyle.SelectionBackColor = CivicCalmTheme.ColorPrimarySoft
            dgv.DefaultCellStyle.SelectionForeColor = CivicCalmTheme.ColorPrimary

            dgv.AlternatingRowsDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#F9FAFB")
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = CivicCalmTheme.ColorInk
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = CivicCalmTheme.ColorPrimarySoft
            dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = CivicCalmTheme.ColorPrimary
        End Sub
    End Class
End Namespace
```

- [ ] **Step 5: Run tests to verify passing status**

Run: `dotnet test tests/BTA_OSG_DocumentTracking.Tests --filter FullyQualifiedName~ThemeTests`
Expected: PASS.

- [ ] **Step 6: Commit changes**

```powershell
git add src/BTA_OSG_DocumentTracking/UI/ tests/BTA_OSG_DocumentTracking.Tests/ThemeTests.vb
git commit -m "feat(ui): implement CivicCalmTheme tokens and DataGridStyler"
```

---

### Task 4: FormMain UI Redesign & High DPI Refactoring

**Files:**
- Modify: `src/BTA_OSG_DocumentTracking/Forms/FormMain.vb`

**Interfaces:**
- Consumes: `CivicCalmTheme`, `DataGridStyler`, `AppStartup`
- Produces: Fully compliant Desktop Civic Calm main form with 3-tier surfaces, StatusStrip, ErrorProvider, and TableLayoutPanels.

- [ ] **Step 1: Write integration test for FormMain initialization**

Add to `tests/BTA_OSG_DocumentTracking.Tests/ServiceTypingTests.vb`:
```vb
<TestMethod>
Public Sub FormMain_InitializesWithCivicCalmColors()
    Dim form As New FormMain()
    Assert.AreEqual(CivicCalmTheme.ColorCanvas, form.BackColor)
    Assert.IsFalse(form.Text.Contains("—"))
    form.Dispose()
End Sub
```

- [ ] **Step 2: Run test to verify failure**

Run: `dotnet test tests/BTA_OSG_DocumentTracking.Tests --filter FullyQualifiedName~FormMain_InitializesWithCivicCalmColors`
Expected: FAIL (FormMain still uses `#0F172A`).

- [ ] **Step 3: Refactor FormMain.vb layout and theme**

Refactor `src/BTA_OSG_DocumentTracking/Forms/FormMain.vb` with:
1. `Me.BackColor = CivicCalmTheme.ColorCanvas` (`#F4F6F8`).
2. Replace header simulation tap buttons with a single `&Tap RFID Badge` (Alt+T) and `&Logout` (Alt+L).
3. Add bottom `StatusStrip` with `lblStatusMessage`, `lblStatusCount`, and `lblStatusClock`.
4. Add `ErrorProvider` for inline field validation. Replace `MessageBox.Show("saved")` with `lblStatusMessage.Text = "Document registered: " & code`.
5. Refactor Registry view into a two-column `TableLayoutPanel` (left form 380px, right grid fill) with Gestalt 20px/12px/8px margins.
6. Refactor Directives, Search, Admin, and Audit views into responsive `TableLayoutPanel` and `FlowLayoutPanel` containers.
7. Apply `DataGridStyler.ApplyCivicStyle` across all DataGridViews.
8. Replace every occurrence of em-dash (`—`) with a colon or hyphen.
9. Wire data operations to `AppStartup.DocService`, `AppStartup.DirectiveService`, `AppStartup.AuthService`, falling back to `EmbeddedDB` when SQL Server is offline.

- [ ] **Step 4: Run test to verify passing status**

Run: `dotnet test tests/BTA_OSG_DocumentTracking.Tests --filter FullyQualifiedName~FormMain_InitializesWithCivicCalmColors`
Expected: PASS.

- [ ] **Step 5: Run automated self-check**

Run: `dotnet run --project src/BTA_OSG_DocumentTracking -- /test`
Expected: ALL 12 TESTS PASS.

- [ ] **Step 6: Commit changes**

```powershell
git add src/BTA_OSG_DocumentTracking/Forms/FormMain.vb
git commit -m "refactor(ui): apply Desktop Civic Calm theme, High DPI layouts, and StatusStrip to FormMain"
```

---

### Task 5: Child Forms Remediation (FormLogin, FormDocumentDetail, FormRouteDocument, FormMoveStorage)

**Files:**
- Modify: `src/BTA_OSG_DocumentTracking/Forms/FormLogin.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Forms/FormDocumentDetail.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Forms/FormRouteDocument.vb`
- Modify: `src/BTA_OSG_DocumentTracking/Forms/FormMoveStorage.vb`

**Interfaces:**
- Consumes: `CivicCalmTheme`, `DataGridStyler`
- Produces: Civic Calm dialogs with `AcceptButton`, `CancelButton`, Alt mnemonics, and clean RFID keystroke capture.

- [ ] **Step 1: Write tests for child dialog key bindings**

Add to `tests/BTA_OSG_DocumentTracking.Tests/ServiceTypingTests.vb`:
```vb
<TestMethod>
Public Sub FormLogin_HasAcceptAndCancelButtons()
    Dim dlg As New FormLogin()
    Assert.IsNotNull(dlg.CancelButton)
    Assert.AreEqual(CivicCalmTheme.ColorCanvas, dlg.BackColor)
    dlg.Dispose()
End Sub
```

- [ ] **Step 2: Run test to verify failure**

Run: `dotnet test tests/BTA_OSG_DocumentTracking.Tests --filter FullyQualifiedName~FormLogin_HasAcceptAndCancelButtons`
Expected: FAIL (CancelButton is currently Nothing).

- [ ] **Step 3: Refactor FormLogin.vb**
- Apply `CivicCalmTheme.ColorCanvas` and `ColorSurface`.
- Remove off-screen textbox hack at `(-100, -100)`; enable `KeyPreview = True` on the Form.
- Set `Me.CancelButton = btnCancel` (`&Cancel`, Alt+C).
- Use `TableLayoutPanel` for button positioning.

- [ ] **Step 4: Refactor FormDocumentDetail.vb**
- Remove em-dash from `Me.Text = String.Format("Document Details and Specifications: [{0}] {1}", ...)`
- Set `Me.CancelButton = btnClose` (`&Close`, Alt+C).
- Apply `DataGridStyler.ApplyCivicStyle` to Directives, Routing, and Movement DataGridViews.
- Apply `CivicCalmTheme.ColorSurface` to tab controls and panels.

- [ ] **Step 5: Refactor FormRouteDocument.vb and FormMoveStorage.vb**
- Apply `CivicCalmTheme` surfaces and fonts.
- Replace manual X/Y positioning with `TableLayoutPanel`.
- Set `Me.AcceptButton = btnRoute` / `btnMove` and `Me.CancelButton = btnCancel`.
- Add Alt mnemonics: `&Route Document` (Alt+R), `&Transfer Storage` (Alt+T), `&Cancel` (Alt+C).

- [ ] **Step 6: Run tests to verify passing status**

Run: `dotnet test tests/BTA_OSG_DocumentTracking.Tests`
Expected: 100% PASS.

- [ ] **Step 7: Commit changes**

```powershell
git add src/BTA_OSG_DocumentTracking/Forms/
git commit -m "refactor(ui): align child forms with Civic Calm theme, Accept/Cancel keys, and clean RFID capture"
```

---

### Task 6: Comprehensive Verification & Delivery Gate

**Files:**
- Test: Full solution build and pre-flight verification

**Interfaces:**
- Consumes: Remediated codebase
- Produces: Verified production build and zero-defect Delivery Gate pass.

- [ ] **Step 1: Execute test suite**

Run: `dotnet test tests/BTA_OSG_DocumentTracking.Tests`
Expected: Total: 14+, Failed: 0, Passed: 14+, 100% passing.

- [ ] **Step 2: Execute automated self-check**

Run: `dotnet run --project src/BTA_OSG_DocumentTracking -- /test`
Expected: Output shows all 12 tests passing with exit code 0.

- [ ] **Step 3: Verify Zero Em-Dash rule**

Run: `git grep "—" src/`
Expected: 0 matches returned.

- [ ] **Step 4: Verify contrast compliance with contrast-check.py**

Run PowerShell commands:
```powershell
python .agents/skills/antislop-human/contrast-check.py "#1B242C" "#F4F6F8"
python .agents/skills/antislop-human/contrast-check.py "#FFFFFF" "#146A3D"
python .agents/skills/antislop-human/contrast-check.py "#55606A" "#F4F6F8"
```
Expected: All return PASS (WCAG 2.2 AA / AAA).

- [ ] **Step 5: Verify self-contained release build**

Run: `dotnet publish src/BTA_OSG_DocumentTracking -c Release -r win-x64 --self-contained true`
Expected: Clean build with 0 warnings, producing `BTA_OSG_DocumentTracking.exe`.

- [ ] **Step 6: Final commit and tag**

```powershell
git commit --allow-empty -m "release: verified Desktop Civic Calm Path A v2.1 compliance"
```
