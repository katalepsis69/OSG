# BTA OSG Document Tracking : DESIGN.md (Desktop Civic Calm)

Desktop UI/UX specification for the Windows Forms (.NET 10) Document Status Tracking and Monitoring System.
Target Environment: Windows 10/11 x64 and Windows Server 2022 desktop workstations.
Design System Philosophy: Desktop Civic Calm (Derived from universal design engineering principles, filtered for desktop WinForms).

---

## 0. Executive Design Read and The 3 Dials

* **Domain:** Official parliamentary document tracking and chain of custody for the Office of the Secretary-General (OSG), Bangsamoro Parliament.
* **Target Users:** Administrative clerks, legislative docket officers, and tracking supervisors processing hundreds of incoming communications, legislative bills, and executive orders daily.
* **Design Read:** Institutional desktop administrative console engineered for high throughput, zero visual latency, high contrast readability under fluorescent lighting, and complete keyboard operability.

### The 3 Dials
* **`DESIGN_VARIANCE: 2` (Parliamentary Predictability)**
  - Fixed, predictable screen structures across all forms.
  - No asymmetric layouts, experimental navigation, or floating unanchored controls.
  - Consistent header banner, standardized search filter bar, central DataGridView, and docked bottom status bar.
* **`MOTION_INTENSITY: 1` (Instantaneous Native Response)**
  - Zero decorative animation, zero spring physics, zero transition delays.
  - State changes happen in 0 to 50ms.
  - All CPU cycles reserved for database queries, RFID processing, and UI responsiveness.
* **`VISUAL_DENSITY: 6` (High-Throughput Desktop Density)**
  - Information-dense tabular layouts allowing clerks to view 15 to 25 records without vertical scrolling on a 1080p display.
  - Compact toolbars (28px to 32px height) and optimized grid row padding (28px row template).

---

## 1. Institutional Palette and 3-Tier Optical Surfaces

The visual hierarchy uses an institutional civic palette inspired by the Bangsamoro parliamentary identity, engineered for WCAG 2.2 AAA contrast standards.

### Color Tokens

| Token Name | Hex Code | Purpose and Context | WCAG Contrast on White |
|---|---|---|---|
| `--color-canvas` | `#F4F6F8` | Form background, window canvas, outer layout margins | Neutral Base |
| `--color-surface` | `#FFFFFF` | Work area panels, GroupBox interiors, DataGridView canvas | Base White |
| `--color-well` | `#EEF2F5` | DataGridView header row, inactive search inputs, summary chips | Inset Surface |
| `--color-border` | `#DDE2E5` | Hairline 1px structural borders, gridlines, separator lines | Optical Divider |
| `--color-ink` | `#1B242C` | Primary typography, headers, active values | 14.2:1 (AAA) |
| `--color-ink-muted` | `#55606A` | Labels, helper microcopy, timestamps, watermark hints | 6.2:1 (AA) |
| `--color-primary` | `#146A3D` | Primary action buttons (`&Save`, `&Route`), active tab borders | 5.8:1 (AA) |
| `--color-primary-soft` | `#E6F2EB` | Selected DataGridView row wash, verified status badge | Accent Wash |
| `--color-accent-sg` | `#B08524` | Secretary-General special directives, gold urgent flags | 4.6:1 (AA) |
| `--color-accent-soft` | `#FEF9E7` | Secretary-General directive alert panel background | Warm Highlight |
| `--color-danger` | `#A93226` | Terminal cancellation, overdue warning flags, critical errors | 6.1:1 (AA) |
| `--color-danger-soft` | `#FBEAE8` | Overdue badge wash, validation error banner wash | Soft Alert |

### 3-Tier Optical Surface Hierarchy

WinForms desktop applications achieve clarity through distinct surface elevation tiers without heavy drop shadows:

1. **Tier 0 : Outer Canvas (`#F4F6F8`)**
   - The root Form `BackColor`. Provides a calm, low-glare backdrop that frames data panels.
2. **Tier 1 : Content Card (`#FFFFFF` with `#DDE2E5` 1px border)**
   - Used for primary editing panels, search group boxes, and DataGridView containers.
   - Distinct from the canvas, giving clear visual boundaries to active workspaces.
3. **Tier 2 : Recessed Inset Well (`#EEF2F5`)**
   - Used for table header rows, summary record counters, and read-only metadata fields.
   - Signals uneditable or structural framing elements.

### Status Badge Semantics

* **Pending / In Transit:** `#52606B` text on `#ECEDF0` wash.
* **Received / Verified:** `#146A3D` text on `#E6F2EB` wash.
* **Secretary-General Directive:** `#B08524` text on `#FEF9E7` wash with bold indicator.
* **Action Overdue / Urgent:** `#A93226` text on `#FBEAE8` wash.
* **Archived / Physical Storage:** `#3A4550` text on `#E2E6EA` wash.

---

## 2. Segoe UI Typography Ladder and High DPI Scaling

### System Typeface Ladder

The entire user interface strictly uses `Segoe UI`, the native Windows system font, with explicit point sizes and weights:

| Role | Font Specification | Alignment | Usage Context |
|---|---|---|---|
| **Form Title** | `Segoe UI`, 12pt Bold | Left | Main window and dialog headers |
| **Section Header** | `Segoe UI`, 10pt Bold | Left | GroupBox labels, panel headers |
| **Field Label** | `Segoe UI`, 9pt Bold | Left | Input descriptors above or beside fields |
| **Body Controls** | `Segoe UI`, 9pt Regular | Left | TextBoxes, ComboBoxes, standard buttons |
| **Tabular Data** | `Segoe UI`, 9pt Regular | Left | DataGridView text cells |
| **Identifiers & Timestamps** | `Segoe UI`, 8.5pt Regular | Right / Tabular | Document tracking numbers, dates, row counts |
| **Microcopy & Hints** | `Segoe UI`, 8pt Regular | Left | Field helper text, status bar info |

### High DPI Scaling Architecture

1. **System-Aware Initialization:**
   `Application.SetHighDpiMode(HighDpiMode.SystemAware)` must be called in `Program.Main` prior to initializing any visual forms.
2. **Layout Container Enforcement:**
   - Absolute pixel positioning (`Location = New Point(x, y)`) is banned for responsive layouts.
   - All forms must structure controls within `TableLayoutPanel`, `FlowLayoutPanel`, or explicit `Dock` / `Anchor` properties.
   - Form padding and control margins must use scalable integer units to prevent clipping on 125% and 150% display scaling factors.
3. **Auto-Ellipsis Protection:**
   Set `AutoEllipsis = True` on long labels and table column headers to prevent text truncation bugs.

---

## 3. Gestalt Desktop Spacing and Layout Invariants

The spatial layout obeys strict Gestalt proximity rules:

$$\text{Outer Form Margin (20–24px)} > \text{GroupBox / Panel Margin (12–16px)} > \text{Control-to-Control Gap (6–8px)} > \text{Label-to-Input Gap (4px)}$$

### Layout Rules

1. **Form Margins:**
   Root forms maintain `Padding = New Padding(20)` around the outer perimeter to prevent visual crowding against the window frame.
2. **Container Padding:**
   Panels and GroupBoxes maintain `Padding = New Padding(12)` on all internal borders.
3. **Control Spacing:**
   Adjacent input fields, buttons, and dropdowns maintain a standard `Margin = New Padding(4)` (yielding an 8px gutter between adjacent controls).
4. **Interactive Target Heights:**
   - Primary action buttons: 30px to 34px height.
   - Secondary / utility buttons: 26px to 28px height.
   - Text boxes and combo boxes: 24px to 26px height.
   - DataGridView row template height: 28px to 32px height.
5. **Zero Deadzone Rule:**
   Clickable rows, tab headers, and action buttons must fill their entire visual bounds. Padding is placed inside controls, never between broken click boundaries.

---

## 4. DataGridView Workhorse Engineering

The `DataGridView` is the central operational tool of the application. It must follow strict performance and readability standards:

### Performance and Visual Configuration
* `AutoGenerateColumns = False` (all columns explicitly defined with data bindings).
* `DoubleBuffered = True` enabled on all grids to eliminate repaint flicker during fast scrolling.
* `SelectionMode = DataGridViewSelectionMode.FullRowSelect`.
* `MultiSelect = False` (unless batch routing is explicitly enabled).
* `RowHeadersVisible = False` (eliminates wasted horizontal gutter space).
* `GridColor = ColorTranslator.FromHtml("#DDE2E5")`.
* Alternate row striping:
  - Default row `BackColor`: `#FFFFFF`.
  - Alternating row `BackColor`: `#F9FAFB`.
* Selected row style:
  - `SelectionBackColor`: `#E6F2EB` (soft green wash).
  - `SelectionForeColor`: `#146A3D` (bold institutional green).

### Column Alignment Standards
* Document Tracking Number: Left-aligned, monospace/tabular Segoe UI.
* Title / Subject / Source Agency: Left-aligned, fill weight column.
* Status Badge: Centered.
* Date / Time Stamp: Right-aligned.
* Sequence / Row Numbers: Right-aligned.

### 4-State Asynchronous DataGrid Lifecycle

Every DataGridView must support four distinct operational states:

1. **Loading State:**
   - Status bar displays: "Loading records from database...".
   - Application cursor toggles to `Cursors.WaitCursor`.
   - Grid content is preserved or displays an active loading overlay; UI remains responsive.
2. **Empty State:**
   - If a query returns 0 rows, a clean centered watermark label displays:
     *"No documents match the current filter criteria. Press Alt+C to clear filters."*
   - Never leave a blank grey canvas with empty gridlines.
3. **Error State:**
   - If a database timeout or network disconnection occurs, an inline banner appears above the grid:
     *"Unable to retrieve document records. [Retry Button] (Alt+R)"*.
   - Never throw raw unhandled SQL exceptions to the user.
4. **Populated State:**
   - Full-row selection active.
   - Bottom status bar reflects row count: *"Showing 42 documents | Last updated 10:14 AM"*.

---

## 5. Clerk Ergonomics, Keyboard Flow and Hardware Integration

High-speed administrative work requires minimizing mouse dependency. A trained clerk must be able to register, route, and print document slips entirely via keyboard and RFID scans.

### Keyboard-First Standards

1. **Sequential Tab Order:**
   - Every input form must have a strictly verified top-to-bottom, left-to-right `TabIndex` flow.
   - Uneditable display labels must have `TabStop = False`.
2. **Alt-Key Mnemonics:**
   Every primary button, menu item, and tab must include an accessible mnemonic:
   - `&New Document` (Alt+N)
   - `&Save` (Alt+S)
   - `&Route Document` (Alt+R)
   - `&Print Routing Slip` (Alt+P)
   - `&Filter / Search` (Alt+F)
   - `&Clear Filters` (Alt+C)
   - `&Close / Exit` (Alt+X)
3. **Dialog Confirmation Keys:**
   - Dialog `AcceptButton` bound to default confirmation (`btnSave` or `btnSearch`).
   - Dialog `CancelButton` bound to cancel/escape (`btnCancel`).
4. **Visual Focus Indicator:**
   Active input fields must have a visible highlight state (e.g. 1px active border or distinct focus color) so the clerk instantly knows where input focus resides.

### USB RFID Scanner Integration

1. **Non-Blocking Card Capture:**
   The RFID reader listener intercepts badge scans globally. The clerk does not need to manually click into an RFID text box before swiping.
2. **Input Sanitization:**
   Card scan inputs are trimmed of trailing carriage returns and control characters before credential verification.
3. **Lockout Protection:**
   Five consecutive invalid card reads lock the terminal and log a security alert to the append-only audit trail.

---

## 6. Non-Modal Feedback and Notification Architecture

1. **StatusStrip Over Popups:**
   - Routine successful actions (e.g. saving a record, routing a file, applying a filter) must update the bottom `StatusStrip` with a transient confirmation message.
   - Banned: Do not show blocking `MessageBox.Show("Document saved successfully!")` dialogs that force clerks to hit Enter on every operation.
2. **Modal Confirmation Reserved for Destruction:**
   - Modal message boxes are reserved exclusively for irreversible actions:
     - Document cancellation or voiding.
     - Terminal logout or session override.
     - Unsaved changes confirmation when closing an active edit form.
3. **Inline Validation:**
   - Missing required fields display an `ErrorProvider` icon directly beside the field with clear explanatory tooltip text, rather than throwing alert boxes.

---

## 7. Pre-Flight Mechanical Quality Checklist

Before committing any form or UI code, verify all seven points:

1. **Display Scaling Check:** Does the layout render without clipping or overflow at 100%, 125%, and 150% display scaling?
2. **Tab Order Verification:** Does pressing Tab cycle through all input fields in logical sequence?
3. **Mnemonic Audit:** Do all action buttons have unique, working Alt-key mnemonics?
4. **Thread Safety Audit:** Are all background threads and hardware listeners marshaled via `If Me.InvokeRequired Then Me.Invoke(...)`?
5. **Contrast Compliance:** Does all body text meet WCAG 2.2 AA standards (minimum 4.5:1 on white)?
6. **DataGrid Double-Buffering:** Is double-buffering active on every DataGridView to eliminate scroll flicker?
7. **Zero Em-Dash Compliance:** Are all documentation strings, comments, and microcopy free of em-dash (`—`) characters?
