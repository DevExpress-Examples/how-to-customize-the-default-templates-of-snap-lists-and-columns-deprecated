Imports System
Imports System.Collections.Generic
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
Imports DevExpress.Snap.Core.API
Imports DevExpress.XtraRichEdit.API.Native
Imports Snap_Events.DataSet1TableAdapters

' ...
Namespace Snap_Events

    Public Partial Class Form1
        Inherits System.Windows.Forms.Form

        Const employeeStyleName As String = "Employees"

        Const customerStyleName As String = "Customers"

        Const employeeDataSourceName As String = "Employees"

        Const customerDataSourceName As String = "Customers"

        Private ReadOnly dataFieldInfoComparer As Snap_Events.Form1.DataFieldInfoComparerType = New Snap_Events.Form1.DataFieldInfoComparerType()

        Private targetColumnIndex As Integer

        Private targetColumnsCount As Integer

        Public Sub New()
            Me.InitializeComponent()
            Me.InitializeDataSources()
            Me.InitializeStyles()
            Me.RegisterEventHandlers()
        End Sub

        Private Sub snapControl1_InitializeDocument(ByVal sender As Object, ByVal e As System.EventArgs)
            Me.InitializeStyles()
        End Sub

        Private Sub InitializeDataSources()
            Dim dataSource = New Snap_Events.DataSet1()
            Dim connection = New System.Data.OleDb.OleDbConnection()
            connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\nwind.mdb"
            Dim employees As Snap_Events.DataSet1TableAdapters.EmployeesTableAdapter = New Snap_Events.DataSet1TableAdapters.EmployeesTableAdapter()
            employees.Connection = connection
            employees.Fill(dataSource.Employees)
            Dim customers As Snap_Events.DataSet1TableAdapters.CustomersTableAdapter = New Snap_Events.DataSet1TableAdapters.CustomersTableAdapter()
            customers.Connection = connection
            customers.Fill(dataSource.Customers)
            Dim employeeCustomers As Snap_Events.DataSet1TableAdapters.EmployeeCustomersTableAdapter = New Snap_Events.DataSet1TableAdapters.EmployeeCustomersTableAdapter()
            employeeCustomers.Connection = connection
            employeeCustomers.Fill(dataSource.EmployeeCustomers)
            Dim orders As Snap_Events.DataSet1TableAdapters.OrdersTableAdapter = New Snap_Events.DataSet1TableAdapters.OrdersTableAdapter()
            orders.Connection = connection
            orders.Fill(dataSource.Orders)
            Dim orderDetails As Snap_Events.DataSet1TableAdapters.Order_DetailsTableAdapter = New Snap_Events.DataSet1TableAdapters.Order_DetailsTableAdapter()
            orderDetails.Connection = connection
            orderDetails.Fill(dataSource.Order_Details)
            Dim document As DevExpress.Snap.Core.API.SnapDocument = Me.snapControl1.Document
            document.BeginUpdateDataSource()
            Dim employeesBinding = New System.Windows.Forms.BindingSource() With {.DataSource = dataSource, .DataMember = Snap_Events.Form1.employeeDataSourceName}
            document.DataSources.Add(New DevExpress.Snap.Core.API.DataSourceInfo(Snap_Events.Form1.employeeDataSourceName, employeesBinding))
            Dim customersBinding = New System.Windows.Forms.BindingSource() With {.DataSource = dataSource, .DataMember = Snap_Events.Form1.customerDataSourceName}
            document.DataSources.Add(New DevExpress.Snap.Core.API.DataSourceInfo(Snap_Events.Form1.customerDataSourceName, customersBinding))
            document.EndUpdateDataSource()
        End Sub

        Private Sub InitializeStyles()
            Dim document As DevExpress.Snap.Core.API.SnapDocument = Me.snapControl1.Document
            document.BeginUpdate()
            Dim employee As DevExpress.XtraRichEdit.API.Native.TableStyle = document.TableStyles.CreateNew()
            employee.Name = Snap_Events.Form1.employeeStyleName
            employee.CellBackgroundColor = System.Drawing.Color.PaleGreen
            document.TableStyles.Add(employee)
            Dim customer As DevExpress.XtraRichEdit.API.Native.TableStyle = document.TableStyles.CreateNew()
            customer.Name = Snap_Events.Form1.customerStyleName
            customer.CellBackgroundColor = System.Drawing.Color.Plum
            document.TableStyles.Add(customer)
            document.TableStyles(CStr(("List1"))).CellBackgroundColor = System.Drawing.Color.White
            document.TableStyles(CStr(("List2"))).CellBackgroundColor = System.Drawing.Color.LightYellow
            document.EndUpdate()
        End Sub

        Private Sub RegisterEventHandlers()
            Dim document As DevExpress.Snap.Core.API.SnapDocument = Me.snapControl1.Document
            AddHandler document.BeforeInsertSnList, AddressOf Me.document_BeforeInsertSnList
            AddHandler document.PrepareSnList, AddressOf Me.document_PrepareSnList
            AddHandler document.BeforeInsertSnListColumns, AddressOf Me.document_BeforeInsertSnListColumns
            AddHandler document.AfterInsertSnListColumns, AddressOf Me.document_AfterInsertSnListColumns
            AddHandler document.BeforeInsertSnListDetail, AddressOf Me.document_BeforeInsertSnListDetail
            AddHandler document.PrepareSnListDetail, AddressOf Me.document_PrepareSnListDetail
        End Sub

        Private Sub document_BeforeInsertSnList(ByVal sender As Object, ByVal e As DevExpress.Snap.Core.API.BeforeInsertSnListEventArgs)
            If e.DataFields.Count = 0 Then Return
            Dim dataSource As System.Windows.Forms.BindingSource = TryCast(e.DataFields(CInt((0))).DataSource, System.Windows.Forms.BindingSource)
            If dataSource Is Nothing Then Return
            ' If data member is Employee data table, make the inserted list always contain 
            ' FirstName and LastName data fields.
            If dataSource.DataMember.Equals(Snap_Events.Form1.employeeDataSourceName) Then
                Dim firstName As DevExpress.Snap.Core.API.DataFieldInfo = New DevExpress.Snap.Core.API.DataFieldInfo(dataSource, "FirstName")
                Dim lastName As DevExpress.Snap.Core.API.DataFieldInfo = New DevExpress.Snap.Core.API.DataFieldInfo(dataSource, "LastName")
                If Not e.DataFields.Contains(lastName, Me.dataFieldInfoComparer) Then e.DataFields.Insert(0, lastName)
                If Not e.DataFields.Contains(firstName, Me.dataFieldInfoComparer) Then e.DataFields.Insert(0, firstName)
            ' If data member is Customerts data table, make the inserted list always contain 
            ' ContactName field.
            ElseIf dataSource.DataMember.Equals(Snap_Events.Form1.customerDataSourceName) Then
                Dim contactName As DevExpress.Snap.Core.API.DataFieldInfo = New DevExpress.Snap.Core.API.DataFieldInfo(dataSource, "ContactName")
                If Not e.DataFields.Contains(contactName, Me.dataFieldInfoComparer) Then e.DataFields.Insert(0, contactName)
            End If
        End Sub

        Private Sub document_PrepareSnList(ByVal sender As Object, ByVal e As DevExpress.Snap.Core.API.PrepareSnListEventArgs)
            ' Change the style applied to the SnapList depending on its data source.
            For i As Integer = 0 To e.Template.Fields.Count - 1
                Dim field As DevExpress.XtraRichEdit.API.Native.Field = e.Template.Fields(i)
                Dim eTemplateParseField As DevExpress.Snap.Core.API.SnapEntity = e.Template.ParseField(field)
                Dim snList As DevExpress.Snap.Core.API.SnapList = TryCast(eTemplateParseField, DevExpress.Snap.Core.API.SnapList)
                If snList Is Nothing Then Continue For
                If snList.DataSourceName.Equals(Snap_Events.Form1.employeeDataSourceName) Then
                    snList.BeginUpdate()
                    Me.SetTablesStyle(snList, Snap_Events.Form1.employeeStyleName)
                    snList.EndUpdate()
                ElseIf snList.DataSourceName.Equals(Snap_Events.Form1.customerDataSourceName) Then
                    snList.BeginUpdate()
                    Me.SetTablesStyle(snList, Snap_Events.Form1.customerStyleName)
                    snList.EndUpdate()
                End If
            Next
        End Sub

        Private Sub SetTablesStyle(ByVal snList As DevExpress.Snap.Core.API.SnapList, ByVal styleName As String)
            Me.SetTablesStyleCore(snList.ListHeader, styleName)
            Me.SetTablesStyleCore(snList.RowTemplate, styleName)
            Me.SetTablesStyleCore(snList.ListFooter, styleName)
        End Sub

        Private Sub SetTablesStyleCore(ByVal document As DevExpress.Snap.Core.API.SnapDocument, ByVal styleName As String)
            Dim style As DevExpress.XtraRichEdit.API.Native.TableStyle = document.TableStyles(styleName)
            If style Is Nothing Then Return
            For Each table As DevExpress.XtraRichEdit.API.Native.Table In document.Tables
                table.Style = style
            Next
        End Sub

        Private Sub document_BeforeInsertSnListColumns(ByVal sender As Object, ByVal e As DevExpress.Snap.Core.API.BeforeInsertSnListColumnsEventArgs)
            Me.targetColumnIndex = e.TargetColumnIndex
            Me.targetColumnsCount = e.DataFields.Count
        End Sub

        Private Sub document_AfterInsertSnListColumns(ByVal sender As Object, ByVal e As DevExpress.Snap.Core.API.AfterInsertSnListColumnsEventArgs)
            ' Mark the inserted columns with red double borders.
            Dim snList As DevExpress.Snap.Core.API.SnapList = e.SnList
            snList.BeginUpdate()
            snList.RowTemplate.Tables(CInt((0))).ForEachRow(Sub(row, rowIdx)
                Dim leftBorder As DevExpress.XtraRichEdit.API.Native.TableCellBorder = row.Cells(CInt((Me.targetColumnIndex))).Borders.Left
                Dim rightBorder As DevExpress.XtraRichEdit.API.Native.TableCellBorder = row.Cells(CInt((Me.targetColumnIndex + Me.targetColumnsCount - 1))).Borders.Right
                Me.setBorder(leftBorder)
                Me.setBorder(rightBorder)
            End Sub)
            snList.EndUpdate()
            snList.Field.Update()
        End Sub

        Private Sub setBorder(ByVal border As DevExpress.XtraRichEdit.API.Native.TableCellBorder)
            border.LineStyle = DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle.[Double]
            border.LineColor = System.Drawing.Color.Red
            border.LineThickness = 1.0F
        End Sub

        Private Sub document_BeforeInsertSnListDetail(ByVal sender As Object, ByVal e As DevExpress.Snap.Core.API.BeforeInsertSnListDetailEventArgs)
            ' Force the EditorRowLimit property of the master list to be lesser than or equal to 5 after
            ' a detail list have been added.
            Dim masterList As DevExpress.Snap.Core.API.SnapList = e.Master
            If masterList.EditorRowLimit <= 5 Then Return
            masterList.BeginUpdate()
            masterList.EditorRowLimit = 5
            masterList.EndUpdate()
        End Sub

        Private Sub document_PrepareSnListDetail(ByVal sender As Object, ByVal e As DevExpress.Snap.Core.API.PrepareSnListDetailEventArgs)
            ' Set the row limit for every inserted detail list.
            For Each field As DevExpress.XtraRichEdit.API.Native.Field In e.Template.Fields
                Dim detailList As DevExpress.Snap.Core.API.SnapList = TryCast(e.Template.ParseField(field), DevExpress.Snap.Core.API.SnapList)
                If detailList Is Nothing Then Continue For
                detailList.BeginUpdate()
                detailList.EditorRowLimit = 5
                detailList.EndUpdate()
            Next
        End Sub

        Private Class DataFieldInfoComparerType
            Implements System.Collections.Generic.IEqualityComparer(Of DevExpress.Snap.Core.API.DataFieldInfo)

'#Region "IEqualityComparer<DataFieldInfo> Members"
            Public Overloads Function Equals(ByVal x As DevExpress.Snap.Core.API.DataFieldInfo, ByVal y As DevExpress.Snap.Core.API.DataFieldInfo) As Boolean Implements Global.System.Collections.Generic.IEqualityComparer(Of Global.DevExpress.Snap.Core.API.DataFieldInfo).Equals
                If x Is Nothing Then Return y Is Nothing
                If y Is Nothing Then Return False
                If Not System.[Object].ReferenceEquals(x.DataSource, y.DataSource) Then Return False
                Dim n As Integer = x.DataPaths.Length
                If y.DataPaths.Length <> n Then Return False
                For i As Integer = 0 To n - 1
                    If Not String.Equals(x.DataPaths(i), y.DataPaths(i)) Then Return False
                Next

                Return True
            End Function

            Public Overloads Function GetHashCode(ByVal obj As DevExpress.Snap.Core.API.DataFieldInfo) As Integer Implements Global.System.Collections.Generic.IEqualityComparer(Of Global.DevExpress.Snap.Core.API.DataFieldInfo).GetHashCode
                If obj Is Nothing Then Return 0
                Dim hash As Integer = obj.DataSource.GetHashCode()
                For Each path As String In obj.DataPaths
                    hash = hash Xor path.GetHashCode()
                Next

                Return hash
            End Function
'#End Region
        End Class
    End Class
End Namespace
