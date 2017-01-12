﻿'------------------------------------------------------------------------------
' <autogenerated>
'     This code was generated by a tool.
'     Runtime Version: 1.1.4322.573
'
'     Changes to this file may cause incorrect behavior and will be lost if 
'     the code is regenerated.
' </autogenerated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.Data
Imports System.Runtime.Serialization
Imports System.Xml


<Serializable(),  _
 System.ComponentModel.DesignerCategoryAttribute("code"),  _
 System.Diagnostics.DebuggerStepThrough(),  _
 System.ComponentModel.ToolboxItem(true)>  _
Public Class DsRemarks
    Inherits DataSet
    
    Private tableRemrksVW As RemrksVWDataTable
    
    Public Sub New()
        MyBase.New
        Me.InitClass
        Dim schemaChangedHandler As System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
        AddHandler Me.Tables.CollectionChanged, schemaChangedHandler
        AddHandler Me.Relations.CollectionChanged, schemaChangedHandler
    End Sub
    
    Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.New
        Dim strSchema As String = CType(info.GetValue("XmlSchema", GetType(System.String)),String)
        If (Not (strSchema) Is Nothing) Then
            Dim ds As DataSet = New DataSet
            ds.ReadXmlSchema(New XmlTextReader(New System.IO.StringReader(strSchema)))
            If (Not (ds.Tables("RemrksVW")) Is Nothing) Then
                Me.Tables.Add(New RemrksVWDataTable(ds.Tables("RemrksVW")))
            End If
            Me.DataSetName = ds.DataSetName
            Me.Prefix = ds.Prefix
            Me.Namespace = ds.Namespace
            Me.Locale = ds.Locale
            Me.CaseSensitive = ds.CaseSensitive
            Me.EnforceConstraints = ds.EnforceConstraints
            Me.Merge(ds, false, System.Data.MissingSchemaAction.Add)
            Me.InitVars
        Else
            Me.InitClass
        End If
        Me.GetSerializationData(info, context)
        Dim schemaChangedHandler As System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
        AddHandler Me.Tables.CollectionChanged, schemaChangedHandler
        AddHandler Me.Relations.CollectionChanged, schemaChangedHandler
    End Sub
    
    <System.ComponentModel.Browsable(false),  _
     System.ComponentModel.DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Content)>  _
    Public ReadOnly Property RemrksVW As RemrksVWDataTable
        Get
            Return Me.tableRemrksVW
        End Get
    End Property
    
    Public Overrides Function Clone() As DataSet
        Dim cln As DsRemarks = CType(MyBase.Clone,DsRemarks)
        cln.InitVars
        Return cln
    End Function
    
    Protected Overrides Function ShouldSerializeTables() As Boolean
        Return false
    End Function
    
    Protected Overrides Function ShouldSerializeRelations() As Boolean
        Return false
    End Function
    
    Protected Overrides Sub ReadXmlSerializable(ByVal reader As XmlReader)
        Me.Reset
        Dim ds As DataSet = New DataSet
        ds.ReadXml(reader)
        If (Not (ds.Tables("RemrksVW")) Is Nothing) Then
            Me.Tables.Add(New RemrksVWDataTable(ds.Tables("RemrksVW")))
        End If
        Me.DataSetName = ds.DataSetName
        Me.Prefix = ds.Prefix
        Me.Namespace = ds.Namespace
        Me.Locale = ds.Locale
        Me.CaseSensitive = ds.CaseSensitive
        Me.EnforceConstraints = ds.EnforceConstraints
        Me.Merge(ds, false, System.Data.MissingSchemaAction.Add)
        Me.InitVars
    End Sub
    
    Protected Overrides Function GetSchemaSerializable() As System.Xml.Schema.XmlSchema
        Dim stream As System.IO.MemoryStream = New System.IO.MemoryStream
        Me.WriteXmlSchema(New XmlTextWriter(stream, Nothing))
        stream.Position = 0
        Return System.Xml.Schema.XmlSchema.Read(New XmlTextReader(stream), Nothing)
    End Function
    
    Friend Sub InitVars()
        Me.tableRemrksVW = CType(Me.Tables("RemrksVW"),RemrksVWDataTable)
        If (Not (Me.tableRemrksVW) Is Nothing) Then
            Me.tableRemrksVW.InitVars
        End If
    End Sub
    
    Private Sub InitClass()
        Me.DataSetName = "DsRemarks"
        Me.Prefix = ""
        Me.Namespace = "http://www.tempuri.org/DsRemarks.xsd"
        Me.Locale = New System.Globalization.CultureInfo("en-US")
        Me.CaseSensitive = false
        Me.EnforceConstraints = true
        Me.tableRemrksVW = New RemrksVWDataTable
        Me.Tables.Add(Me.tableRemrksVW)
    End Sub
    
    Private Function ShouldSerializeRemrksVW() As Boolean
        Return false
    End Function
    
    Private Sub SchemaChanged(ByVal sender As Object, ByVal e As System.ComponentModel.CollectionChangeEventArgs)
        If (e.Action = System.ComponentModel.CollectionChangeAction.Remove) Then
            Me.InitVars
        End If
    End Sub
    
    Public Delegate Sub RemrksVWRowChangeEventHandler(ByVal sender As Object, ByVal e As RemrksVWRowChangeEvent)
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class RemrksVWDataTable
        Inherits DataTable
        Implements System.Collections.IEnumerable
        
        Private columnADMNO As DataColumn
        
        Private columnREMARKS As DataColumn
        
        Private columnTEACHERNAME As DataColumn
        
        Private columnRDate As DataColumn
        
        Private column_STUDENTS_NAME As DataColumn
        
        Private columnFRNAME As DataColumn
        
        Private columnRollNo As DataColumn
        
        Private column_CLASSMAIN_NAME As DataColumn
        
        Private columnSECTIONNAME As DataColumn
        
        Friend Sub New()
            MyBase.New("RemrksVW")
            Me.InitClass
        End Sub
        
        Friend Sub New(ByVal table As DataTable)
            MyBase.New(table.TableName)
            If (table.CaseSensitive <> table.DataSet.CaseSensitive) Then
                Me.CaseSensitive = table.CaseSensitive
            End If
            If (table.Locale.ToString <> table.DataSet.Locale.ToString) Then
                Me.Locale = table.Locale
            End If
            If (table.Namespace <> table.DataSet.Namespace) Then
                Me.Namespace = table.Namespace
            End If
            Me.Prefix = table.Prefix
            Me.MinimumCapacity = table.MinimumCapacity
            Me.DisplayExpression = table.DisplayExpression
        End Sub
        
        <System.ComponentModel.Browsable(false)>  _
        Public ReadOnly Property Count As Integer
            Get
                Return Me.Rows.Count
            End Get
        End Property
        
        Friend ReadOnly Property ADMNOColumn As DataColumn
            Get
                Return Me.columnADMNO
            End Get
        End Property
        
        Friend ReadOnly Property REMARKSColumn As DataColumn
            Get
                Return Me.columnREMARKS
            End Get
        End Property
        
        Friend ReadOnly Property TEACHERNAMEColumn As DataColumn
            Get
                Return Me.columnTEACHERNAME
            End Get
        End Property
        
        Friend ReadOnly Property RDateColumn As DataColumn
            Get
                Return Me.columnRDate
            End Get
        End Property
        
        Friend ReadOnly Property _STUDENTS_NAMEColumn As DataColumn
            Get
                Return Me.column_STUDENTS_NAME
            End Get
        End Property
        
        Friend ReadOnly Property FRNAMEColumn As DataColumn
            Get
                Return Me.columnFRNAME
            End Get
        End Property
        
        Friend ReadOnly Property RollNoColumn As DataColumn
            Get
                Return Me.columnRollNo
            End Get
        End Property
        
        Friend ReadOnly Property _CLASSMAIN_NAMEColumn As DataColumn
            Get
                Return Me.column_CLASSMAIN_NAME
            End Get
        End Property
        
        Friend ReadOnly Property SECTIONNAMEColumn As DataColumn
            Get
                Return Me.columnSECTIONNAME
            End Get
        End Property
        
        Public Default ReadOnly Property Item(ByVal index As Integer) As RemrksVWRow
            Get
                Return CType(Me.Rows(index),RemrksVWRow)
            End Get
        End Property
        
        Public Event RemrksVWRowChanged As RemrksVWRowChangeEventHandler
        
        Public Event RemrksVWRowChanging As RemrksVWRowChangeEventHandler
        
        Public Event RemrksVWRowDeleted As RemrksVWRowChangeEventHandler
        
        Public Event RemrksVWRowDeleting As RemrksVWRowChangeEventHandler
        
        Public Overloads Sub AddRemrksVWRow(ByVal row As RemrksVWRow)
            Me.Rows.Add(row)
        End Sub
        
        Public Overloads Function AddRemrksVWRow(ByVal ADMNO As Integer, ByVal REMARKS As String, ByVal TEACHERNAME As String, ByVal RDate As Date, ByVal _STUDENTS_NAME As String, ByVal FRNAME As String, ByVal RollNo As Short, ByVal _CLASSMAIN_NAME As String, ByVal SECTIONNAME As String) As RemrksVWRow
            Dim rowRemrksVWRow As RemrksVWRow = CType(Me.NewRow,RemrksVWRow)
            rowRemrksVWRow.ItemArray = New Object() {ADMNO, REMARKS, TEACHERNAME, RDate, _STUDENTS_NAME, FRNAME, RollNo, _CLASSMAIN_NAME, SECTIONNAME}
            Me.Rows.Add(rowRemrksVWRow)
            Return rowRemrksVWRow
        End Function
        
        Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return Me.Rows.GetEnumerator
        End Function
        
        Public Overrides Function Clone() As DataTable
            Dim cln As RemrksVWDataTable = CType(MyBase.Clone,RemrksVWDataTable)
            cln.InitVars
            Return cln
        End Function
        
        Protected Overrides Function CreateInstance() As DataTable
            Return New RemrksVWDataTable
        End Function
        
        Friend Sub InitVars()
            Me.columnADMNO = Me.Columns("ADMNO")
            Me.columnREMARKS = Me.Columns("REMARKS")
            Me.columnTEACHERNAME = Me.Columns("TEACHERNAME")
            Me.columnRDate = Me.Columns("RDate")
            Me.column_STUDENTS_NAME = Me.Columns("STUDENTS.NAME")
            Me.columnFRNAME = Me.Columns("FRNAME")
            Me.columnRollNo = Me.Columns("RollNo")
            Me.column_CLASSMAIN_NAME = Me.Columns("CLASSMAIN.NAME")
            Me.columnSECTIONNAME = Me.Columns("SECTIONNAME")
        End Sub
        
        Private Sub InitClass()
            Me.columnADMNO = New DataColumn("ADMNO", GetType(System.Int32), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnADMNO)
            Me.columnREMARKS = New DataColumn("REMARKS", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnREMARKS)
            Me.columnTEACHERNAME = New DataColumn("TEACHERNAME", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnTEACHERNAME)
            Me.columnRDate = New DataColumn("RDate", GetType(System.DateTime), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnRDate)
            Me.column_STUDENTS_NAME = New DataColumn("STUDENTS.NAME", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.column_STUDENTS_NAME)
            Me.columnFRNAME = New DataColumn("FRNAME", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnFRNAME)
            Me.columnRollNo = New DataColumn("RollNo", GetType(System.Int16), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnRollNo)
            Me.column_CLASSMAIN_NAME = New DataColumn("CLASSMAIN.NAME", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.column_CLASSMAIN_NAME)
            Me.columnSECTIONNAME = New DataColumn("SECTIONNAME", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnSECTIONNAME)
        End Sub
        
        Public Function NewRemrksVWRow() As RemrksVWRow
            Return CType(Me.NewRow,RemrksVWRow)
        End Function
        
        Protected Overrides Function NewRowFromBuilder(ByVal builder As DataRowBuilder) As DataRow
            Return New RemrksVWRow(builder)
        End Function
        
        Protected Overrides Function GetRowType() As System.Type
            Return GetType(RemrksVWRow)
        End Function
        
        Protected Overrides Sub OnRowChanged(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowChanged(e)
            If (Not (Me.RemrksVWRowChangedEvent) Is Nothing) Then
                RaiseEvent RemrksVWRowChanged(Me, New RemrksVWRowChangeEvent(CType(e.Row,RemrksVWRow), e.Action))
            End If
        End Sub
        
        Protected Overrides Sub OnRowChanging(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowChanging(e)
            If (Not (Me.RemrksVWRowChangingEvent) Is Nothing) Then
                RaiseEvent RemrksVWRowChanging(Me, New RemrksVWRowChangeEvent(CType(e.Row,RemrksVWRow), e.Action))
            End If
        End Sub
        
        Protected Overrides Sub OnRowDeleted(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowDeleted(e)
            If (Not (Me.RemrksVWRowDeletedEvent) Is Nothing) Then
                RaiseEvent RemrksVWRowDeleted(Me, New RemrksVWRowChangeEvent(CType(e.Row,RemrksVWRow), e.Action))
            End If
        End Sub
        
        Protected Overrides Sub OnRowDeleting(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowDeleting(e)
            If (Not (Me.RemrksVWRowDeletingEvent) Is Nothing) Then
                RaiseEvent RemrksVWRowDeleting(Me, New RemrksVWRowChangeEvent(CType(e.Row,RemrksVWRow), e.Action))
            End If
        End Sub
        
        Public Sub RemoveRemrksVWRow(ByVal row As RemrksVWRow)
            Me.Rows.Remove(row)
        End Sub
    End Class
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class RemrksVWRow
        Inherits DataRow
        
        Private tableRemrksVW As RemrksVWDataTable
        
        Friend Sub New(ByVal rb As DataRowBuilder)
            MyBase.New(rb)
            Me.tableRemrksVW = CType(Me.Table,RemrksVWDataTable)
        End Sub
        
        Public Property ADMNO As Integer
            Get
                Try 
                    Return CType(Me(Me.tableRemrksVW.ADMNOColumn),Integer)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableRemrksVW.ADMNOColumn) = value
            End Set
        End Property
        
        Public Property REMARKS As String
            Get
                Try 
                    Return CType(Me(Me.tableRemrksVW.REMARKSColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableRemrksVW.REMARKSColumn) = value
            End Set
        End Property
        
        Public Property TEACHERNAME As String
            Get
                Try 
                    Return CType(Me(Me.tableRemrksVW.TEACHERNAMEColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableRemrksVW.TEACHERNAMEColumn) = value
            End Set
        End Property
        
        Public Property RDate As Date
            Get
                Try 
                    Return CType(Me(Me.tableRemrksVW.RDateColumn),Date)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableRemrksVW.RDateColumn) = value
            End Set
        End Property
        
        Public Property _STUDENTS_NAME As String
            Get
                Try 
                    Return CType(Me(Me.tableRemrksVW._STUDENTS_NAMEColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableRemrksVW._STUDENTS_NAMEColumn) = value
            End Set
        End Property
        
        Public Property FRNAME As String
            Get
                Try 
                    Return CType(Me(Me.tableRemrksVW.FRNAMEColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableRemrksVW.FRNAMEColumn) = value
            End Set
        End Property
        
        Public Property RollNo As Short
            Get
                Try 
                    Return CType(Me(Me.tableRemrksVW.RollNoColumn),Short)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableRemrksVW.RollNoColumn) = value
            End Set
        End Property
        
        Public Property _CLASSMAIN_NAME As String
            Get
                Try 
                    Return CType(Me(Me.tableRemrksVW._CLASSMAIN_NAMEColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableRemrksVW._CLASSMAIN_NAMEColumn) = value
            End Set
        End Property
        
        Public Property SECTIONNAME As String
            Get
                Try 
                    Return CType(Me(Me.tableRemrksVW.SECTIONNAMEColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableRemrksVW.SECTIONNAMEColumn) = value
            End Set
        End Property
        
        Public Function IsADMNONull() As Boolean
            Return Me.IsNull(Me.tableRemrksVW.ADMNOColumn)
        End Function
        
        Public Sub SetADMNONull()
            Me(Me.tableRemrksVW.ADMNOColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsREMARKSNull() As Boolean
            Return Me.IsNull(Me.tableRemrksVW.REMARKSColumn)
        End Function
        
        Public Sub SetREMARKSNull()
            Me(Me.tableRemrksVW.REMARKSColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsTEACHERNAMENull() As Boolean
            Return Me.IsNull(Me.tableRemrksVW.TEACHERNAMEColumn)
        End Function
        
        Public Sub SetTEACHERNAMENull()
            Me(Me.tableRemrksVW.TEACHERNAMEColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsRDateNull() As Boolean
            Return Me.IsNull(Me.tableRemrksVW.RDateColumn)
        End Function
        
        Public Sub SetRDateNull()
            Me(Me.tableRemrksVW.RDateColumn) = System.Convert.DBNull
        End Sub
        
        Public Function Is_STUDENTS_NAMENull() As Boolean
            Return Me.IsNull(Me.tableRemrksVW._STUDENTS_NAMEColumn)
        End Function
        
        Public Sub Set_STUDENTS_NAMENull()
            Me(Me.tableRemrksVW._STUDENTS_NAMEColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsFRNAMENull() As Boolean
            Return Me.IsNull(Me.tableRemrksVW.FRNAMEColumn)
        End Function
        
        Public Sub SetFRNAMENull()
            Me(Me.tableRemrksVW.FRNAMEColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsRollNoNull() As Boolean
            Return Me.IsNull(Me.tableRemrksVW.RollNoColumn)
        End Function
        
        Public Sub SetRollNoNull()
            Me(Me.tableRemrksVW.RollNoColumn) = System.Convert.DBNull
        End Sub
        
        Public Function Is_CLASSMAIN_NAMENull() As Boolean
            Return Me.IsNull(Me.tableRemrksVW._CLASSMAIN_NAMEColumn)
        End Function
        
        Public Sub Set_CLASSMAIN_NAMENull()
            Me(Me.tableRemrksVW._CLASSMAIN_NAMEColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsSECTIONNAMENull() As Boolean
            Return Me.IsNull(Me.tableRemrksVW.SECTIONNAMEColumn)
        End Function
        
        Public Sub SetSECTIONNAMENull()
            Me(Me.tableRemrksVW.SECTIONNAMEColumn) = System.Convert.DBNull
        End Sub
    End Class
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class RemrksVWRowChangeEvent
        Inherits EventArgs
        
        Private eventRow As RemrksVWRow
        
        Private eventAction As DataRowAction
        
        Public Sub New(ByVal row As RemrksVWRow, ByVal action As DataRowAction)
            MyBase.New
            Me.eventRow = row
            Me.eventAction = action
        End Sub
        
        Public ReadOnly Property Row As RemrksVWRow
            Get
                Return Me.eventRow
            End Get
        End Property
        
        Public ReadOnly Property Action As DataRowAction
            Get
                Return Me.eventAction
            End Get
        End Property
    End Class
End Class