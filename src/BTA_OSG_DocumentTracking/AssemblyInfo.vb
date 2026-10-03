Option Explicit On
Option Strict On

Imports System.Runtime.CompilerServices

' The MSTest suite verifies the internal offline-cache machinery (scratch DbPath, replay
' flags) that production code keeps out of its public surface.
<Assembly: InternalsVisibleTo("BTA_OSG_DocumentTracking.Tests")>
