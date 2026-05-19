Imports System.Collections
Imports System.Windows.Forms
Namespace SI.Controls
    Public Class LvSort
        Inherits System.Windows.Forms.ListView
#Region "Description"
        '
        '########################################################################################
        '
        ' Class     LvSort
        ' Author    Jan Schröder
        '           Schröder Informatik GmbH
        '           www.SchroederInformatik.de
        ' Version   1.0.3
        ' Date      05/03/08
        '
        ' This class is adding sort functionality to the inherited class "ListView" by 
        ' implementing a non case sensitive IComparer and visualizes the sort order by drawing 
        ' an appropriate triangle in the sorting column header.
        '
        ' It also adds the missing events "Scroll" and "Paint".
        '
        ' Use this class as follows:

        ' 1. Copy "LvSort.vb" to your project.
        ' 2. Add a "System.Windows.Forms.ListView" to your form.
        ' 3. Replace "System.Windows.Forms.ListView" with "SI.Controls.LvSort" in 
        '    the generated Code.
        '
        ' The advantage of this kind of usage is, you will not have to deploy an additional DLL,
        ' all code will be part of your executable.
        '
        '
        ' Versions:
        ' 1.0.0 05/03/02    First Version
        ' 1.0.1 05/03/05    LblMeasureHeader.Font = MyListViewValue.Font in
        '                   ListViewColumnSorter.VisualizeOrder
        ' 1.0.2 05/03/08    Invoking sort method when setting SortColum or Order, so setting one
        '                   of these properties (i.e. in a forms load event) will cause sorting.
        '                   Tip: To avoid an unnecessary sort, set Order to "none" before setting
        '                   the SortColumn and then Order to "ascending" or "descending".
        '                   If  you want to sort the list view by the fist two columns, notice 
        '                   the following sequence:
        '                   1. ListView1.Order = SortOrder.None
        '                   2. ListView1.SortColumn = 1
        '                   3. ListView1.Order = SortOrder.Ascending
        '                   4. ListView1.SortColumn = 0
        '                   After step 2 there will be no sorting, because Order = None. After
        '                   step 3 the list view will be sorted by column 1 in ascending order.
        '                   After step 4 the list view will be sorted by column 0 in ascending
        '                   order and if there are items with identical text in column 0, they
        '                   will be sorted by column 1.
        ' 1.0.3 05/03/08    MyBase.ListViewItemSorter = Nothing in properties Order and
        '                   SortColumn to avoid unnecessary sorting when loading the view.
        '                   To do so, Order has to be set to none or SortColumn to -1.
        ' 1.0.4 05/05/24    Translation from VB into C#
        '
        '########################################################################################
        '
#End Region
#Region "Constructors"
        Public Sub New()
            MyBase.New()
            LvwColumnSorter = New ListViewColumnSorter(Me)
        End Sub
#End Region
#Region "Data definition"
        Public Property FixedStringForLastItem() As String
            '
            ' Fixed string for the last item, so it will stay the last item
            '
            Get
                Return LvwColumnSorter.FixedStringForLastItem
            End Get
            Set(ByVal Value As String)
                '
                ' The IComparer has to know that
                '
                LvwColumnSorter.FixedStringForLastItem = Value
            End Set
        End Property
        Public Property SortColumn() As Integer
            '
            ' Index of sort column
            '
            Get
                Return LvwColumnSorter.SortColumn
            End Get
            Set(ByVal Value As Integer)
                '
                ' Put it through to IComparer
                '
                LvwColumnSorter.SortColumn = Value
                '
                ' Let's do it
                '
                If LvwColumnSorter.Order = SortOrder.None Then
                    MyBase.ListViewItemSorter = Nothing
                Else
                    If MyBase.ListViewItemSorter Is Nothing Then
                        MyBase.ListViewItemSorter = LvwColumnSorter
                    Else
                        MyBase.Sort()
                    End If
                End If
                '
            End Set
        End Property
        Public Property Order() As SortOrder
            '
            ' Sort order (none, ascending or descending)
            '
            Get
                Return LvwColumnSorter.Order
            End Get
            Set(ByVal Value As SortOrder)
                '
                ' Last information needed for IComparer
                '
                LvwColumnSorter.Order = Value
                If Value = SortOrder.None Then
                    If LvwColumnSorter.SortColumn > -1 Then
                        '
                        ' Setting the text causes a repaint of the header
                        '
                        MyBase.Columns(LvwColumnSorter.SortColumn).Text =
                            MyBase.Columns(LvwColumnSorter.SortColumn).Text
                        '
                        ' If there has been a triangle from a previous sorting,
                        ' it is now deleted
                        '
                    End If
                    LvwColumnSorter.SortColumn = -1
                End If
                '
                ' Let's do it
                '
                If LvwColumnSorter.SortColumn = -1 Then
                    MyBase.ListViewItemSorter = Nothing
                Else
                    If MyBase.ListViewItemSorter Is Nothing Then
                        MyBase.ListViewItemSorter = LvwColumnSorter
                    Else
                        MyBase.Sort()
                    End If
                End If
                '
            End Set
        End Property
        Private WithEvents LvlListener As LvlListenerClass
        Private WithEvents HdrListener As HdrListenerClass
        Private LvwColumnSorter As ListViewColumnSorter
#End Region
#Region "Event handling"
        Public Event Scroll(ByVal sender As Object, ByVal e As EventArgs)
        Public Shadows Event Paint(ByVal sender As Object, ByVal e As EventArgs)
        Private Class LvlListenerClass
            '
            ' Listen for operating system messages to raise the events
            ' "Scroll" or "Paint" for the ListView
            '
            Inherits NativeWindow
            Public Event Scroll(ByVal sender As Object, ByVal e As EventArgs)
            Public Event Paint(ByVal sender As Object, ByVal e As EventArgs)
            Const WM_HSCROLL = &H114
            Const WM_VSCROLL = &H115
            Const WM_PAINT = &HF
            Private ReadOnly Ctrl As Control
            Public Sub New()
                If Ctrl Is Nothing Then
                    Throw New ArgumentNullException(NameOf(Ctrl))
                End If

                AssignHandle(Ctrl.Handle)
                Me.Ctrl = Ctrl
            End Sub
            Protected Overrides Sub WndProc(ByRef m As Message)
                MyBase.WndProc(m)
                If m.Msg = WM_HSCROLL Or m.Msg = WM_VSCROLL Then
                    RaiseEvent Scroll(Ctrl, New EventArgs())
                End If
                If m.Msg = WM_PAINT Then
                    RaiseEvent Paint(Ctrl, New EventArgs())
                End If
            End Sub
            Protected Overrides Sub Finalize()
                ReleaseHandle()
                MyBase.Finalize()
            End Sub
        End Class
        Private Class HdrListenerClass
            '
            ' Listen for operating system messages to raise the event
            ' "HaederPaint", when the column headers are to be painted.
            ' On this event, a triangle, symbolizing sort order and
            ' column is to be drawn.
            '
            Inherits NativeWindow
            Public Event HaederPaint(ByVal sender As Object, ByVal e As EventArgs)
            Const WM_PAINT = &HF
            Private CtrlValue As Control
            Public Sub New(ByVal Ctrl As Control, ByVal HeaderHandle As System.IntPtr)
                AssignHandle(HeaderHandle)
                CtrlValue = Ctrl
            End Sub
            Protected Overrides Sub WndProc(ByRef m As Message)
                MyBase.WndProc(m)
                If m.Msg = WM_PAINT Then
                    RaiseEvent HaederPaint(CtrlValue, New EventArgs())
                End If
            End Sub
            Protected Overrides Sub Finalize()
                ReleaseHandle()
                MyBase.Finalize()
            End Sub
        End Class
        Protected Overrides Sub OnHandleCreated(ByVal e As EventArgs)
            '
            ' Now it's the right time to do some initializations
            '
            MyBase.OnHandleCreated(e)
            If Not Me.DesignMode Then
                LvlListener = New LvlListenerClass()
                LvwColumnSorter.GetHeaderHandle()
                HdrListener = New HdrListenerClass(Me, LvwColumnSorter.HeaderHandle)
            End If
        End Sub
        Private Sub LvlListener_Paint(ByVal sender As Object, ByVal e As System.EventArgs) _
                Handles LvlListener.Paint
            '
            ' Make this event public
            '
            RaiseEvent Paint(sender, e)
            '
        End Sub
        Private Sub LvlListener_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) _
                Handles LvlListener.Scroll
            '
            ' Make this event public
            '
            RaiseEvent Scroll(sender, e)
            '
        End Sub
        Private Sub HdrListener_HaederPaint(ByVal sender As Object,
                ByVal e As System.EventArgs) Handles HdrListener.HaederPaint
            '
            ' The column headers has been painted, so draw the sort triangle
            '
            LvwColumnSorter.VisualizeOrder()
            '
        End Sub
        Private Sub MyBase_ColumnClick(ByVal sender As Object, ByVal e As _
                System.Windows.Forms.ColumnClickEventArgs) Handles MyBase.ColumnClick
            '
            ' The user wants to sort the list view items
            '
            If (e.Column = SortColumn) Then
                '
                ' The column has been clicked twice, so switch the sort order
                '
                If (Order = SortOrder.Ascending) Then
                    Order = SortOrder.Descending
                Else
                    Order = SortOrder.Ascending
                End If
            Else
                '
                ' It has to be sort ascending by the column
                '
                SortColumn = e.Column
                Order = SortOrder.Ascending
            End If
            '
        End Sub
#End Region
#Region "API stuff"
        '
        ' The API function ChildWindowFromPoint is used to find out the window handle of
        ' the columns header.
        '
        Private Structure GdiPoint
            Dim x As Integer
            Dim y As Integer
        End Structure
        Private Declare Function ChildWindowFromPoint Lib "user32" _
            (ByVal hWndParent As System.IntPtr, ByVal Point As GdiPoint) As System.IntPtr
        '
#End Region
#Region "Subroutines and functions"
        Private Class ListViewColumnSorter
            Implements System.Collections.IComparer
            '
            ' This class implements an non case sensitve IComparer for
            ' sorting items in a ListView
            '
            Private ColumnToSort As Integer
            Private OrderOfSort As SortOrder
            Private ObjectCompare As CaseInsensitiveComparer
            Private FixedStringForLastItemValue As String
            Private MyListViewValue As LvSort
            Public Sub New(ByVal MyListView As LvSort)
                ColumnToSort = 0
                OrderOfSort = SortOrder.None
                ObjectCompare = New CaseInsensitiveComparer()
                MyListViewValue = MyListView
            End Sub
            Public Property FixedStringForLastItem() As String
                '
                ' Fixed string for the last item, so it will stay the last item
                '
                Get
                    Return FixedStringForLastItemValue
                End Get
                Set(ByVal Value As String)
                    FixedStringForLastItemValue = Value
                End Set
            End Property
            Public Property SortColumn() As Integer
                '
                ' Index of sort column
                '
                Set(ByVal Value As Integer)
                    ColumnToSort = Value
                End Set
                Get
                    Return ColumnToSort
                End Get
            End Property
            Public Property Order() As SortOrder
                '
                ' Sort order (none, ascending or descending)
                '
                Set(ByVal Value As SortOrder)
                    OrderOfSort = Value
                    VisualizeOrder()
                End Set
                Get
                    Return OrderOfSort
                End Get
            End Property
            Public ReadOnly Property HeaderHandle() As System.IntPtr
                '
                ' The handle of the columns header
                '
                Get
                    Return HeaderHandleValue
                End Get
            End Property
            Public Sub GetHeaderHandle()
                '
                ' Find out the handle of the columns header
                '
                Dim TestPoint As GdiPoint
                TestPoint.x = 5
                TestPoint.y = 5
                HeaderHandleValue =
                    ChildWindowFromPoint(MyListViewValue.Handle, TestPoint)
            End Sub
            Public Function Compare(ByVal a As Object, ByVal b As Object) As _
                    Integer Implements IComparer.Compare
                '
                ' Using a case insensitive comparer to compare the Text of
                ' two list view items
                '
                If OrderOfSort = SortOrder.None Then
                    '
                    ' Return 0 means that both items are equal, so no
                    ' change of the sequence will occure
                    '
                    Return 0
                    '
                End If
                '
                Dim Result As Integer
                Dim LvIa As ListViewItem = a
                Dim LvIb As ListViewItem = b
                Dim Ca As String = LvIa.SubItems(ColumnToSort).Text
                Dim Cb As String = LvIb.SubItems(ColumnToSort).Text
                '
                ' Compare the two strings
                '
                Result = ObjectCompare.Compare(Ca, Cb)
                If FixedStringForLastItemValue Is Nothing Then
                    ' nothing
                Else
                    If Cb = FixedStringForLastItemValue Then
                        Return -1
                    End If
                End If
                '
                ' If the strings contains numbers, correct the sequence. For example:
                ' 3, 20, 1 has to be sorted as 1, 3, 20 and not 1, 20, 3
                '
                If Result <> 0 And IsNumeric(Ca) And IsNumeric(Cb) Then
                    If (CDbl(Ca) > CDbl(Cb)) Then
                        Result = 1
                    Else
                        Result = -1
                    End If
                End If
                '
                ' The return value depends on the sort order
                '
                Select Case OrderOfSort
                    Case SortOrder.Ascending
                        '
                        ' As ascending sort is desired, return the compare result
                        '
                        Return Result
                        '
                    Case SortOrder.Descending
                        '
                        ' As descending sort is desired, compare result is to be
                        ' turned in the negative
                        '
                        Return -Result
                        '
                End Select
            End Function
            Public Sub VisualizeOrder()
                '
                ' The sort column and order are visualized by a little triangle,
                ' which is to be drawn right to the columns name
                '
                If ColumnToSort = -1 Then
                    Exit Sub
                End If
                '
                ' Draw the triangle in the header. Before that, 
                ' make theold triangle invisible. After that, store the 
                ' position of the triangle for usage at the next time.
                '
                If SortColumnOld > -1 And SortColumnOld <> ColumnToSort Then
                    '
                    ' Setting the text causes a repaint of the header
                    '
                    MyListViewValue.Columns(SortColumnOld).Text =
                        MyListViewValue.Columns(SortColumnOld).Text.TrimEnd
                    '
                End If
                If OrderOfSort = SortOrder.None Then
                    SortColumnOld = -1
                Else
                    Dim Grx As Graphics
                    Grx = Graphics.FromHwnd(HeaderHandle)
                    Dim LblMeasureHeader As New Label()
                    LblMeasureHeader.Font = MyListViewValue.Font
                    LblMeasureHeader.AutoSize = True
                    If MyListViewValue.Columns(ColumnToSort).TextAlign <>
                        HorizontalAlignment.Left Then
                        If MyListViewValue.Columns(ColumnToSort).Text.TrimEnd =
                                MyListViewValue.Columns(ColumnToSort).Text Then
                            MyListViewValue.Columns(ColumnToSort).Text += Space(5)
                        End If
                    End If
                    LblMeasureHeader.Text = MyListViewValue.Columns(ColumnToSort).Text
                    Dim x, y, i As Integer
                    Dim Triangle(2) As Point
                    '
                    ' Evaluate the length of the Name
                    '
                    x = LblMeasureHeader.Width + 13
                    ' 
                    ' Keep a minimum distance of the right bounce
                    '
                    If x + 15 > MyListViewValue.Columns(ColumnToSort).Width Then
                        x = MyListViewValue.Columns(ColumnToSort).Width - 13
                    End If
                    '
                    ' Sum the width of all columns left to the sort column
                    '
                    For i = 0 To ColumnToSort - 1
                        x += MyListViewValue.Columns(i).Width
                    Next
                    '
                    ' Top is determined by the text height
                    '
                    y = LblMeasureHeader.Height / 2 + 2
                    '
                    ' The three points of the triangle are depending on
                    ' the sort order
                    '
                    If OrderOfSort = SortOrder.Ascending Then
                        Triangle(0).X = x - 1
                        Triangle(0).Y = y + 4
                        Triangle(1).X = x + 9
                        Triangle(1).Y = y + 4
                        Triangle(2).X = x + 4
                        Triangle(2).Y = y - 2
                    Else
                        Triangle(0).X = x
                        Triangle(0).Y = y - 1
                        Triangle(1).X = x + 9
                        Triangle(1).Y = y - 1
                        Triangle(2).X = x + 4
                        Triangle(2).Y = y + 4
                    End If
                    Grx.FillPolygon(SystemBrushes.ControlDark, Triangle)
                    SortColumnOld = ColumnToSort
                    Grx.Dispose()
                End If
            End Sub
            Private SortColumnOld As Integer = -1
            Private HeaderHandleValue As System.IntPtr
        End Class
#End Region
    End Class
End Namespace