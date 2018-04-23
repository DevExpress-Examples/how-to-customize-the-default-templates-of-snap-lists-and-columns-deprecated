Imports Microsoft.VisualBasic
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
	Partial Public Class Form1
		Inherits Form
		Private Const employeeStyleName As String = "Employees"
		Private Const customerStyleName As String = "Customers"
		Private Const employeeDataSourceName As String = "Employees"
		Private Const customerDataSourceName As String = "Customers"
        Private ReadOnly dataFieldInfoComparer1 As New DataFieldInfoComparer()

		Private targetColumnIndex As Integer
		Private targetColumnsCount As Integer

		Public Sub New()
			InitializeComponent()
			InitializeDataSources()
			InitializeStyles()
			RegisterEventHandlers()
		End Sub

		Private Sub snapControl1_InitializeDocument(ByVal sender As Object, ByVal e As EventArgs)
			InitializeStyles()
		End Sub

		Private Sub InitializeDataSources()
			Dim dataSource = New DataSet1()
			Dim connection = New OleDbConnection()
			connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\nwind.mdb"

			Dim employees As New EmployeesTableAdapter()
			employees.Connection = connection
			employees.Fill(dataSource.Employees)

			Dim customers As New CustomersTableAdapter()
			customers.Connection = connection
			customers.Fill(dataSource.Customers)

			Dim employeeCustomers As New EmployeeCustomersTableAdapter()
			employeeCustomers.Connection = connection
			employeeCustomers.Fill(dataSource.EmployeeCustomers)

			Dim orders As New OrdersTableAdapter()
			orders.Connection = connection
			orders.Fill(dataSource.Orders)

			Dim orderDetails As New Order_DetailsTableAdapter()
			orderDetails.Connection = connection
			orderDetails.Fill(dataSource.Order_Details)

			Dim document As SnapDocument = snapControl1.Document
			document.BeginUpdateDataSource()

			Dim employeesBinding = New BindingSource() With {.DataSource = dataSource, .DataMember = employeeDataSourceName}
			document.DataSources.Add(New DataSourceInfo(employeeDataSourceName, employeesBinding))

			Dim customersBinding = New BindingSource() With {.DataSource = dataSource, .DataMember = customerDataSourceName}
			document.DataSources.Add(New DataSourceInfo(customerDataSourceName, customersBinding))

			document.EndUpdateDataSource()
		End Sub

		Private Sub InitializeStyles()
			Dim document As SnapDocument = snapControl1.Document
			document.BeginUpdate()

			Dim employee As TableStyle = document.TableStyles.CreateNew()
			employee.Name = employeeStyleName
			employee.CellBackgroundColor = Color.PaleGreen
			document.TableStyles.Add(employee)

			Dim customer As TableStyle = document.TableStyles.CreateNew()
			customer.Name = customerStyleName
			customer.CellBackgroundColor = Color.Plum
			document.TableStyles.Add(customer)

			document.TableStyles("List1").CellBackgroundColor = Color.White
			document.TableStyles("List2").CellBackgroundColor = Color.LightYellow
			document.EndUpdate()
		End Sub

		Private Sub RegisterEventHandlers()
			Dim document As SnapDocument = snapControl1.Document

			AddHandler document.BeforeInsertSnList, AddressOf document_BeforeInsertSnList
			AddHandler document.PrepareSnList, AddressOf document_PrepareSnList
			AddHandler document.BeforeInsertSnListColumns, AddressOf document_BeforeInsertSnListColumns
			AddHandler document.AfterInsertSnListColumns, AddressOf document_AfterInsertSnListColumns
			AddHandler document.BeforeInsertSnListDetail, AddressOf document_BeforeInsertSnListDetail
			AddHandler document.PrepareSnListDetail, AddressOf document_PrepareSnListDetail


		End Sub

		Private Sub document_BeforeInsertSnList(ByVal sender As Object, ByVal e As BeforeInsertSnListEventArgs)
			If e.DataFields.Count = 0 Then
				Return
			End If
			Dim dataSource As BindingSource = TryCast(e.DataFields(0).DataSource, BindingSource)
			If dataSource Is Nothing Then
				Return
			End If
			' If data member is Employee data table, make the inserted list always contain 
			' FirstName and LastName data fields.
			If dataSource.DataMember.Equals(employeeDataSourceName) Then
				Dim firstName As New DataFieldInfo(dataSource, "FirstName")
				Dim lastName As New DataFieldInfo(dataSource, "LastName")
                If (Not e.DataFields.Contains(lastName, dataFieldInfoComparer1)) Then
                    e.DataFields.Insert(0, lastName)
                End If
                If (Not e.DataFields.Contains(firstName, dataFieldInfoComparer1)) Then
                    e.DataFields.Insert(0, firstName)
                End If
			' If data member is Customerts data table, make the inserted list always contain 
			' ContactName field.
			ElseIf dataSource.DataMember.Equals(customerDataSourceName) Then
				Dim contactName As New DataFieldInfo(dataSource, "ContactName")
                If (Not e.DataFields.Contains(contactName, dataFieldInfoComparer1)) Then
                    e.DataFields.Insert(0, contactName)
                End If
			End If
		End Sub

		Private Sub document_PrepareSnList(ByVal sender As Object, ByVal e As PrepareSnListEventArgs)
			' Change the style applied to the SnapList depending on its data source.
			For i As Integer = 0 To e.Template.Fields.Count - 1
				Dim field As Field = e.Template.Fields(i)
				Dim eTemplateParseField As SnapEntity = e.Template.ParseField(field)
				Dim snList As SnapList = TryCast(eTemplateParseField, SnapList)
				If snList Is Nothing Then
					Continue For
				End If
				If snList.DataSourceName.Equals(employeeDataSourceName) Then
					snList.BeginUpdate()
					SetTablesStyle(snList, employeeStyleName)
					snList.EndUpdate()
				ElseIf snList.DataSourceName.Equals(customerDataSourceName) Then
					snList.BeginUpdate()
					SetTablesStyle(snList, customerStyleName)
					snList.EndUpdate()
				End If
			Next i
		End Sub

		Private Sub SetTablesStyle(ByVal snList As SnapList, ByVal styleName As String)
			SetTablesStyleCore(snList.ListHeader, styleName)
			SetTablesStyleCore(snList.RowTemplate, styleName)
			SetTablesStyleCore(snList.ListFooter, styleName)
		End Sub

		Private Sub SetTablesStyleCore(ByVal document As SnapDocument, ByVal styleName As String)
			Dim style As TableStyle = document.TableStyles(styleName)
			If style Is Nothing Then
				Return
			End If
			For Each table As Table In document.Tables
				table.Style = style
			Next table
		End Sub

		Private Sub document_BeforeInsertSnListColumns(ByVal sender As Object, ByVal e As BeforeInsertSnListColumnsEventArgs)
			targetColumnIndex = e.TargetColumnIndex
			targetColumnsCount = e.DataFields.Count
		End Sub

		Private Sub document_AfterInsertSnListColumns(ByVal sender As Object, ByVal e As AfterInsertSnListColumnsEventArgs)
			' Mark the inserted columns with red double borders.
			Dim snList As SnapList = e.SnList
			snList.BeginUpdate()
			snList.RowTemplate.Tables(0).ForEachRow(Function(row, rowIdx) AnonymousMethod1(row, rowIdx))
			snList.EndUpdate()
			snList.Field.Update()
		End Sub
		
        Private Function AnonymousMethod1(ByVal row As TableRow, ByVal rowIdx As Integer) As Boolean
            Dim leftBorder As TableCellBorder = row.Cells(targetColumnIndex).Borders.Left
            Dim rightBorder As TableCellBorder = row.Cells(targetColumnIndex + targetColumnsCount - 1).Borders.Right
            setBorder(leftBorder)
            setBorder(rightBorder)
            Return True
        End Function

		Private Sub setBorder(ByVal border As TableCellBorder)
			border.LineStyle = TableBorderLineStyle.Double
			border.LineColor = Color.Red
			border.LineThickness = 1.0f
		End Sub

		Private Sub document_BeforeInsertSnListDetail(ByVal sender As Object, ByVal e As BeforeInsertSnListDetailEventArgs)
			' Force the EditorRowLimit property of the master list to be lesser than or equal to 5 after
			' a detail list have been added.
			Dim masterList As SnapList = e.Master
			If masterList.EditorRowLimit <= 5 Then
				Return
			End If
			masterList.BeginUpdate()
			masterList.EditorRowLimit = 5
			masterList.EndUpdate()
		End Sub

		Private Sub document_PrepareSnListDetail(ByVal sender As Object, ByVal e As PrepareSnListDetailEventArgs)
			' Set the row limit for every inserted detail list.
			For Each field As Field In e.Template.Fields
				Dim detailList As SnapList = TryCast(e.Template.ParseField(field), SnapList)
				If detailList Is Nothing Then
					Continue For
				End If
				detailList.BeginUpdate()
				detailList.EditorRowLimit = 5
				detailList.EndUpdate()
			Next field
		End Sub

		Private Class DataFieldInfoComparer
			Implements IEqualityComparer(Of DataFieldInfo)

			#Region "IEqualityComparer<DataFieldInfo> Members"

			Public Overloads Function Equals(ByVal x As DataFieldInfo, ByVal y As DataFieldInfo) As Boolean Implements IEqualityComparer(Of DataFieldInfo).Equals
				If x Is Nothing Then
					Return y Is Nothing
				End If
				If y Is Nothing Then
					Return False
				End If
				If (Not Object.ReferenceEquals(x.DataSource, y.DataSource)) Then
					Return False
				End If
				Dim n As Integer = x.DataPaths.Length
				If y.DataPaths.Length <> n Then
					Return False
				End If
				For i As Integer = 0 To n - 1
					If (Not String.Equals(x.DataPaths(i), y.DataPaths(i))) Then
						Return False
					End If
				Next i
				Return True
			End Function

			Public Overloads Function GetHashCode(ByVal obj As DataFieldInfo) As Integer Implements IEqualityComparer(Of DataFieldInfo).GetHashCode
				If obj Is Nothing Then
					Return 0
				End If
				Dim hash As Integer = obj.DataSource.GetHashCode()
				For Each path As String In obj.DataPaths
					hash = hash Xor path.GetHashCode()
				Next path
				Return hash
			End Function

			#End Region
		End Class
	End Class
End Namespace