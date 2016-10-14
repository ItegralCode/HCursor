Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Public Structure [HCURSOR]
#Region "API needs"
    ' Necessary API additions to carry out the needed tasks.
    <DllImport("user32.dll")>
    Private Shared Function GetCursorInfo(ByRef pci As CURSORINFO) As Boolean
    End Function
    <StructLayout(LayoutKind.Sequential)>
    Structure CURSORINFO
        ' Specifies the size, in bytes, of the structure.
        ' The caller must set this to Marshal.SizeOf(typeof(CURSORINFO)).
        Public cbSize As Int32
        ' Specifies the cursor state. This parameter can be one of the following values:
        '    0                 The cursor is hidden.
        '    CURSOR_SHOWING    The cursor is showing.
        Public flags As Int32
        ' Handle of the cursor.
        Public hCursor As IntPtr
        ' Position on the screen of the Cursor Handle.
        Public ptScreenPos As POINT
    End Structure
    <StructLayout(LayoutKind.Sequential)>
    Structure POINT
        Public x As Int32
        Public y As Int32
    End Structure
    ' Cursor's constant state, compare this value with the Flag of the Cursor's Handle to determine if the Cursor is showing. 
    Private Const CURSOR_SHOWING As Int32 = 1
#End Region
    ' All standard System Cursors to reference when parsing their IntPtr to their reference Name.
    Private Shared ReadOnly HCURSORHANDLES As New [Dictionary](Of [IntPtr], [String]) From
           {{[Cursors].AppStarting.Handle, "AppStarting"},
            {New [IntPtr](787597), "Arrow Flipped"},
            {[Cursors].Arrow.Handle, "Arrow"},
            {[Cursors].Cross.Handle, "Cross"},
            {New [IntPtr](65567), "Hand Alt"},
            {[Cursors].Hand.Handle, "Hand"},
            {[Cursors].Help.Handle, "Help"},
            {[Cursors].HSplit.Handle, "HSplit"},
            {[Cursors].IBeam.Handle, "IBeam"},
            {[Cursors].No.Handle, "No"},
            {[Cursors].NoMove2D.Handle, "NoMove2D"},
            {[Cursors].NoMoveHoriz.Handle, "NoMoveHoriz"},
            {[Cursors].NoMoveVert.Handle, "NoMoveVert"},
            {[Cursors].PanEast.Handle, "PanEast"},
            {[Cursors].PanNE.Handle, "PanNE"},
            {[Cursors].PanNorth.Handle, "PanNorth"},
            {[Cursors].PanNW.Handle, "PanNW"},
            {[Cursors].PanSE.Handle, "PanSE"},
            {[Cursors].PanSouth.Handle, "PanSouth"},
            {[Cursors].PanSW.Handle, "PanSW"},
            {[Cursors].PanWest.Handle, "PanWest"},
            {[Cursors].SizeAll.Handle, "SizeAll"},
            {[Cursors].SizeNESW.Handle, "SizeNESW"},
            {[Cursors].SizeNS.Handle, "SizeNS"},
            {[Cursors].SizeNWSE.Handle, "SizeNWSE"},
            {[Cursors].SizeWE.Handle, "SizeWE"},
            {[Cursors].UpArrow.Handle, "UpArrow"},
            {[Cursors].VSplit.Handle, "VSplit"},
            {[Cursors].WaitCursor.Handle, "WaitCursor"}}
    ' All standard System Cursors to reference to when comparing to the Current Cursor to determine if they are the same Value.
    Public Structure CursorHandles
        Public Shared ReadOnly AppStarting As [IntPtr] = [Cursors].AppStarting.Handle
        Public Shared ReadOnly Arrow As [IntPtr] = [Cursors].Arrow.Handle
        Public Shared ReadOnly ArrowFlipped As IntPtr = New IntPtr(787597)
        Public Shared ReadOnly Cross As [IntPtr] = [Cursors].Cross.Handle
        Public Shared ReadOnly _Default As [IntPtr] = [Cursors].Default.Handle
        Public Shared ReadOnly Hand As [IntPtr] = [Cursors].Hand.Handle
        Public Shared ReadOnly HandAlt As IntPtr = New IntPtr(65567)
        Public Shared ReadOnly Help As [IntPtr] = [Cursors].Help.Handle
        Public Shared ReadOnly HSplit As [IntPtr] = [Cursors].HSplit.Handle
        Public Shared ReadOnly IBeam As [IntPtr] = [Cursors].IBeam.Handle
        Public Shared ReadOnly No As [IntPtr] = [Cursors].No.Handle
        Public Shared ReadOnly NoMove2D As [IntPtr] = [Cursors].NoMove2D.Handle
        Public Shared ReadOnly NoMoveHoriz As [IntPtr] = [Cursors].NoMoveHoriz.Handle
        Public Shared ReadOnly NoMoveVert As [IntPtr] = [Cursors].NoMoveVert.Handle
        Public Shared ReadOnly PanEast As [IntPtr] = [Cursors].PanEast.Handle
        Public Shared ReadOnly PanNE As [IntPtr] = [Cursors].PanNE.Handle
        Public Shared ReadOnly PanNorth As [IntPtr] = [Cursors].PanNorth.Handle
        Public Shared ReadOnly PanNW As [IntPtr] = [Cursors].PanNW.Handle
        Public Shared ReadOnly PanSE As [IntPtr] = [Cursors].PanSE.Handle
        Public Shared ReadOnly PanSouth As [IntPtr] = [Cursors].PanSouth.Handle
        Public Shared ReadOnly PanSW As [IntPtr] = [Cursors].PanSW.Handle
        Public Shared ReadOnly PanWest As [IntPtr] = [Cursors].PanWest.Handle
        Public Shared ReadOnly SizeAll As [IntPtr] = [Cursors].SizeAll.Handle
        Public Shared ReadOnly SizeNESW As [IntPtr] = [Cursors].SizeNESW.Handle
        Public Shared ReadOnly SizeNS As [IntPtr] = [Cursors].SizeNS.Handle
        Public Shared ReadOnly SizeNWSE As [IntPtr] = [Cursors].SizeNWSE.Handle
        Public Shared ReadOnly SizeWE As [IntPtr] = [Cursors].SizeWE.Handle
        Public Shared ReadOnly UpArrow As [IntPtr] = [Cursors].UpArrow.Handle
        Public Shared ReadOnly VSplit As [IntPtr] = [Cursors].VSplit.Handle
        Public Shared ReadOnly WaitCursor As [IntPtr] = [Cursors].WaitCursor.Handle
    End Structure
    ' A make-shift Exception Message that is returned when attempting to retreive the Name from a Cursor's IntPtr and it is not in the HCURSORHANDLES Dictionary List of standard System Cursors.
    ' {May be altered to fit the user's needs.}
    Private Shared ReadOnly [CursorNotFound] As [String] = "Unknown Cursor Handle"
    ' The Previous Cursor is referenced to determine if the Cursor's Handle changed.
    Private Shared [HPREVIOUSCURSOR] As [IntPtr] = Nothing
    Public Shared Sub Initialize()
        ' Add Handler to the Internal timer.
        AddHandler CheckTimer.Tick, AddressOf CheckTimer_Tick
        ' Start the Internal timer.
        ' {Optional: May be carried out nearly just as easily with an external timer.}
        [HCURSOR].CheckTimer.Start()
    End Sub
    Private Shared Function CursorNfo() As [CURSORINFO]
        ' Getting the Cursor Info from the Cursor.
        Dim pci As [CURSORINFO] = New [CURSORINFO] With {.cbSize = [Marshal].SizeOf(GetType([CURSORINFO]))}
        GetCursorInfo(pci)
        ' Returning the Cursor's CURSORINFO.
        Return pci
    End Function
    Public Shared Property PreviousCursor() As [IntPtr]
        Get
            ' Return the Previous Cursor's IntPtr information.
            Return [HPREVIOUSCURSOR]
        End Get
        Set(ByVal value As [IntPtr])
            ' Sets the Previous Cursor to a given value. Should never be assigned by the user manually.
            [HPREVIOUSCURSOR] = [value]
        End Set
    End Property
    Public Shared ReadOnly Property PreviousCursorName() As [String]
        Get
            ' Return the Name of the Previous Cursor.
            Return If([HCURSORHANDLES].Keys.Contains([PreviousCursor]), [HCURSORHANDLES].Item([PreviousCursor]), [CursorNotFound] & " : " & PreviousCursor.ToInt32)
        End Get
    End Property
    Public Shared ReadOnly Property CurrentCursor() As [IntPtr]
        Get
            ' Return the Current Cursor's IntPtr information.
            Return CursorNfo().hCursor
        End Get
    End Property
    Public Shared ReadOnly Property CurrentCursorName() As [String]
        Get
            ' Return the Name of the Current Cursor.
            Return If([HCURSORHANDLES].Keys.Contains([CurrentCursor]), [HCURSORHANDLES].Item([CurrentCursor]), [CursorNotFound] & " : " & CurrentCursor.ToInt32)
        End Get
    End Property
    Public Shared ReadOnly Property CursorNameFromHandle(ByVal cursorHandle As [IntPtr]) As [String]
        Get
            ' Return the Name that corresponds to given Cursor Handle, or return error string if the dictionary doesn't contain the key.
            Return If([HCURSORHANDLES].Keys.Contains([CurrentCursor]), [HCURSORHANDLES].Item([CurrentCursor]), [CursorNotFound].ToString() & " : " & CurrentCursor.ToInt32)
        End Get
    End Property
    ' Add a handler to a [Sub] for this event to more conveniently handle your [Code] when the Cursor's Handle changes.
    Public Shared Event CursorChanged(ByRef Current_hCursor As [IntPtr], ByRef Previous_hCursor As [IntPtr])
    Public Shared Sub DETECTHCURSORCHANGE()
        ' Initialize the Previous Cursor if not done already.
        If [HCURSOR].PreviousCursor = Nothing Then [HCURSOR].PreviousCursor = [HCURSOR].CurrentCursor
        ' Raise the event if the Cursor changed.
        If Not [HCURSOR].PreviousCursor = [HCURSOR].CurrentCursor Then RaiseEvent CursorChanged([CurrentCursor], [PreviousCursor])
        ' Set Previous Cursor to Current Cursor to reinitialize the Cursor handle event detection.
        If Not [HCURSOR].PreviousCursor = [HCURSOR].CurrentCursor Then [HCURSOR].PreviousCursor = [HCURSOR].CurrentCursor
    End Sub
    ' Internal timer that runs every 75ms to update Cursor information in background.
    Public Shared CheckTimer As [Timer] = New [Timer] With {.Interval = 75, .Enabled = True}
    Private Shared Sub CheckTimer_Tick(sender As Object, e As EventArgs)
        ' Call the [Sub] to refresh the Cursor's Handle data and check for Cursor Handle changes.
        [HCURSOR].DETECTHCURSORCHANGE()
    End Sub
    Public Shared Property CurrentLocation() As Drawing.[Point]
        Get
            ' Return Cursor's global Location
            Return System.Windows.Forms.[Cursor].Position
        End Get
        Set(value As Drawing.[Point])
            ' Set the Cursor's global Location
            System.Windows.Forms.[Cursor].Position = value
        End Set
    End Property
End Structure
