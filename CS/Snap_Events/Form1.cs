using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DevExpress.Snap.Core.API;
using DevExpress.XtraRichEdit.API.Native;
using Snap_Events.DataSet1TableAdapters;
// ...

namespace Snap_Events {
    public partial class Form1 : Form {
        const string employeeStyleName = "Employees";
        const string customerStyleName = "Customers";
        const string employeeDataSourceName = "Employees";
        const string customerDataSourceName = "Customers";
        readonly DataFieldInfoComparer dataFieldInfoComparer = new DataFieldInfoComparer();

        int targetColumnIndex;
        int targetColumnsCount;

        public Form1() {
            InitializeComponent();
            InitializeDataSources();
            InitializeStyles();
            RegisterEventHandlers();
        }

        private void snapControl1_InitializeDocument(object sender, EventArgs e) {
            InitializeStyles();
        }

        void InitializeDataSources() {
            var dataSource = new DataSet1();
            var connection = new OleDbConnection();
            connection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\nwind.mdb";

            EmployeesTableAdapter employees = new EmployeesTableAdapter();
            employees.Connection = connection;
            employees.Fill(dataSource.Employees);

            CustomersTableAdapter customers = new CustomersTableAdapter();
            customers.Connection = connection;
            customers.Fill(dataSource.Customers);

            EmployeeCustomersTableAdapter employeeCustomers = new EmployeeCustomersTableAdapter();
            employeeCustomers.Connection = connection;
            employeeCustomers.Fill(dataSource.EmployeeCustomers);

            OrdersTableAdapter orders = new OrdersTableAdapter();
            orders.Connection = connection;
            orders.Fill(dataSource.Orders);

            Order_DetailsTableAdapter orderDetails = new Order_DetailsTableAdapter();
            orderDetails.Connection = connection;
            orderDetails.Fill(dataSource.Order_Details);

            SnapDocument document = snapControl1.Document;
            document.BeginUpdateDataSource();

            var employeesBinding = new BindingSource() { DataSource = dataSource, DataMember = employeeDataSourceName };
            document.DataSources.Add(new DataSourceInfo(employeeDataSourceName, employeesBinding));

            var customersBinding = new BindingSource() { DataSource = dataSource, DataMember = customerDataSourceName };
            document.DataSources.Add(new DataSourceInfo(customerDataSourceName, customersBinding));

            document.EndUpdateDataSource();
        }

        void InitializeStyles() {
            SnapDocument document = snapControl1.Document;
            document.BeginUpdate();

            TableStyle employee = document.TableStyles.CreateNew();
            employee.Name = employeeStyleName;
            employee.CellBackgroundColor = Color.PaleGreen;
            document.TableStyles.Add(employee);

            TableStyle customer = document.TableStyles.CreateNew();
            customer.Name = customerStyleName;
            customer.CellBackgroundColor = Color.Plum;
            document.TableStyles.Add(customer);

            document.TableStyles["List1"].CellBackgroundColor = Color.White;
            document.TableStyles["List2"].CellBackgroundColor = Color.LightYellow;
            document.EndUpdate();
        }

        void RegisterEventHandlers() {
            SnapDocument document = snapControl1.Document;

            document.BeforeInsertSnList += document_BeforeInsertSnList;
            document.PrepareSnList += document_PrepareSnList;
            document.BeforeInsertSnListColumns += document_BeforeInsertSnListColumns;
            document.AfterInsertSnListColumns += document_AfterInsertSnListColumns;
            document.BeforeInsertSnListDetail += document_BeforeInsertSnListDetail;
            document.PrepareSnListDetail += document_PrepareSnListDetail;


        }
        
        void document_BeforeInsertSnList(object sender, BeforeInsertSnListEventArgs e) {
            if(e.DataFields.Count == 0)
                return;
            BindingSource dataSource = e.DataFields[0].DataSource as BindingSource;
            if(dataSource == null)
                return;
            // If data member is Employee data table, make the inserted list always contain 
            // FirstName and LastName data fields.
            if(dataSource.DataMember.Equals(employeeDataSourceName)) {
                DataFieldInfo firstName = new DataFieldInfo(dataSource, "FirstName");
                DataFieldInfo lastName = new DataFieldInfo(dataSource, "LastName");
                if(!e.DataFields.Contains(lastName, dataFieldInfoComparer))
                    e.DataFields.Insert(0, lastName);
                if(!e.DataFields.Contains(firstName, dataFieldInfoComparer))
                    e.DataFields.Insert(0, firstName);
            }
            // If data member is Customerts data table, make the inserted list always contain 
            // ContactName field.
            else if(dataSource.DataMember.Equals(customerDataSourceName)) {
                DataFieldInfo contactName = new DataFieldInfo(dataSource, "ContactName");
                if(!e.DataFields.Contains(contactName, dataFieldInfoComparer))
                    e.DataFields.Insert(0, contactName);
            }
        }

        void document_PrepareSnList(object sender, PrepareSnListEventArgs e) {
            // Change the style applied to the SnapList depending on its data source.
            for(int i = 0; i < e.Template.Fields.Count; i++) {
                Field field = e.Template.Fields[i];
                SnapEntity eTemplateParseField = e.Template.ParseField(field);
                SnapList snList = eTemplateParseField as SnapList;
                if(snList == null)
                    continue;
                if(snList.DataSourceName.Equals(employeeDataSourceName)) {
                    snList.BeginUpdate();
                    SetTablesStyle(snList, employeeStyleName);
                    snList.EndUpdate();
                }
                else if(snList.DataSourceName.Equals(customerDataSourceName)) {
                    snList.BeginUpdate();
                    SetTablesStyle(snList, customerStyleName);
                    snList.EndUpdate();
                }
            }
        }

        void SetTablesStyle(SnapList snList, string styleName) {
            SetTablesStyleCore(snList.ListHeader, styleName);
            SetTablesStyleCore(snList.RowTemplate, styleName);
            SetTablesStyleCore(snList.ListFooter, styleName);
        }

        void SetTablesStyleCore(SnapDocument document, string styleName) {
            TableStyle style = document.TableStyles[styleName];
            if(style == null)
                return;
            foreach(Table table in document.Tables)
                table.Style = style;
        }

        void document_BeforeInsertSnListColumns(object sender, BeforeInsertSnListColumnsEventArgs e) {
            targetColumnIndex = e.TargetColumnIndex;
            targetColumnsCount = e.DataFields.Count;
        }

        void document_AfterInsertSnListColumns(object sender, AfterInsertSnListColumnsEventArgs e) {
            // Mark the inserted columns with red double borders.
            SnapList snList = e.SnList;
            snList.BeginUpdate();
            snList.RowTemplate.Tables[0].ForEachRow((row, rowIdx) => {
                TableCellBorder leftBorder = row.Cells[targetColumnIndex].Borders.Left;
                TableCellBorder rightBorder = row.Cells[targetColumnIndex + targetColumnsCount - 1].Borders.Right;
                setBorder(leftBorder);
                setBorder(rightBorder);
            });
            snList.EndUpdate();
            snList.Field.Update();
        }

        void setBorder(TableCellBorder border) {
            border.LineStyle = TableBorderLineStyle.Double;
            border.LineColor = Color.Red;
            border.LineThickness = 1.0f;
        }

        void document_BeforeInsertSnListDetail(object sender, BeforeInsertSnListDetailEventArgs e) {
            // Force the EditorRowLimit property of the master list to be lesser than or equal to 5 after
            // a detail list have been added.
            SnapList masterList = e.Master;
            if(masterList.EditorRowLimit <= 5)
                return;
            masterList.BeginUpdate();
            masterList.EditorRowLimit = 5;
            masterList.EndUpdate();
        }

        void document_PrepareSnListDetail(object sender, PrepareSnListDetailEventArgs e) {
            // Set the row limit for every inserted detail list.
            foreach(Field field in e.Template.Fields) {
                SnapList detailList = e.Template.ParseField(field) as SnapList;
                if(detailList == null)
                    continue;
                detailList.BeginUpdate();
                detailList.EditorRowLimit = 5;
                detailList.EndUpdate();
            }
        }
        
        class DataFieldInfoComparer : IEqualityComparer<DataFieldInfo> {

            #region IEqualityComparer<DataFieldInfo> Members

            public bool Equals(DataFieldInfo x, DataFieldInfo y) {
                if(x == null)
                    return y == null;
                if(y == null)
                    return false;
                if(!Object.ReferenceEquals(x.DataSource, y.DataSource))
                    return false;
                int n = x.DataPaths.Length;
                if(y.DataPaths.Length != n)
                    return false;
                for(int i = 0; i < n; i++)
                    if(!string.Equals(x.DataPaths[i], y.DataPaths[i]))
                        return false;
                return true;
            }
          
            public int GetHashCode(DataFieldInfo obj) {
                if(obj == null)
                    return 0;
                int hash = obj.DataSource.GetHashCode();
                foreach(string path in obj.DataPaths)
                    hash ^= path.GetHashCode();
                return hash;
            }

            #endregion
        }
    }
}