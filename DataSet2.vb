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
Public Class DataSet2
    Inherits DataSet
    
    Private tablemanual As manualDataTable
    
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
            If (Not (ds.Tables("manual")) Is Nothing) Then
                Me.Tables.Add(New manualDataTable(ds.Tables("manual")))
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
    Public ReadOnly Property manual As manualDataTable
        Get
            Return Me.tablemanual
        End Get
    End Property
    
    Public Overrides Function Clone() As DataSet
        Dim cln As DataSet2 = CType(MyBase.Clone,DataSet2)
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
        If (Not (ds.Tables("manual")) Is Nothing) Then
            Me.Tables.Add(New manualDataTable(ds.Tables("manual")))
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
        Me.tablemanual = CType(Me.Tables("manual"),manualDataTable)
        If (Not (Me.tablemanual) Is Nothing) Then
            Me.tablemanual.InitVars
        End If
    End Sub
    
    Private Sub InitClass()
        Me.DataSetName = "DataSet2"
        Me.Prefix = ""
        Me.Namespace = "http://www.tempuri.org/DataSet2.xsd"
        Me.Locale = New System.Globalization.CultureInfo("ru-RU")
        Me.CaseSensitive = false
        Me.EnforceConstraints = true
        Me.tablemanual = New manualDataTable
        Me.Tables.Add(Me.tablemanual)
    End Sub
    
    Private Function ShouldSerializemanual() As Boolean
        Return false
    End Function
    
    Private Sub SchemaChanged(ByVal sender As Object, ByVal e As System.ComponentModel.CollectionChangeEventArgs)
        If (e.Action = System.ComponentModel.CollectionChangeAction.Remove) Then
            Me.InitVars
        End If
    End Sub
    
    Public Delegate Sub manualRowChangeEventHandler(ByVal sender As Object, ByVal e As manualRowChangeEvent)
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class manualDataTable
        Inherits DataTable
        Implements System.Collections.IEnumerable
        
        Private columnДатаПоследнегоСохранения As DataColumn
        
        Private columnПапкаИдентификатор As DataColumn
        
        Private columnСодержаниеФайла As DataColumn
        
        Private columnСтрока As DataColumn
        
        Private columnТипОбъекта As DataColumn
        
        Private columnФайлВПапке As DataColumn
        
        Private columnФайлИдентификатор As DataColumn
        
        Friend Sub New()
            MyBase.New("manual")
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
        
        Friend ReadOnly Property ДатаПоследнегоСохраненияColumn As DataColumn
            Get
                Return Me.columnДатаПоследнегоСохранения
            End Get
        End Property
        
        Friend ReadOnly Property ПапкаИдентификаторColumn As DataColumn
            Get
                Return Me.columnПапкаИдентификатор
            End Get
        End Property
        
        Friend ReadOnly Property СодержаниеФайлаColumn As DataColumn
            Get
                Return Me.columnСодержаниеФайла
            End Get
        End Property
        
        Friend ReadOnly Property СтрокаColumn As DataColumn
            Get
                Return Me.columnСтрока
            End Get
        End Property
        
        Friend ReadOnly Property ТипОбъектаColumn As DataColumn
            Get
                Return Me.columnТипОбъекта
            End Get
        End Property
        
        Friend ReadOnly Property ФайлВПапкеColumn As DataColumn
            Get
                Return Me.columnФайлВПапке
            End Get
        End Property
        
        Friend ReadOnly Property ФайлИдентификаторColumn As DataColumn
            Get
                Return Me.columnФайлИдентификатор
            End Get
        End Property
        
        Public Default ReadOnly Property Item(ByVal index As Integer) As manualRow
            Get
                Return CType(Me.Rows(index),manualRow)
            End Get
        End Property
        
        Public Event manualRowChanged As manualRowChangeEventHandler
        
        Public Event manualRowChanging As manualRowChangeEventHandler
        
        Public Event manualRowDeleted As manualRowChangeEventHandler
        
        Public Event manualRowDeleting As manualRowChangeEventHandler
        
        Public Overloads Sub AddmanualRow(ByVal row As manualRow)
            Me.Rows.Add(row)
        End Sub
        
        Public Overloads Function AddmanualRow(ByVal ДатаПоследнегоСохранения As Date, ByVal ПапкаИдентификатор As String, ByVal СодержаниеФайла As String, ByVal ТипОбъекта As String, ByVal ФайлВПапке As String, ByVal ФайлИдентификатор As String) As manualRow
            Dim rowmanualRow As manualRow = CType(Me.NewRow,manualRow)
            rowmanualRow.ItemArray = New Object() {ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Nothing, ТипОбъекта, ФайлВПапке, ФайлИдентификатор}
            Me.Rows.Add(rowmanualRow)
            Return rowmanualRow
        End Function
        
        Public Function FindByСтрока(ByVal Строка As Integer) As manualRow
            Return CType(Me.Rows.Find(New Object() {Строка}),manualRow)
        End Function
        
        Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return Me.Rows.GetEnumerator
        End Function
        
        Public Overrides Function Clone() As DataTable
            Dim cln As manualDataTable = CType(MyBase.Clone,manualDataTable)
            cln.InitVars
            Return cln
        End Function
        
        Protected Overrides Function CreateInstance() As DataTable
            Return New manualDataTable
        End Function
        
        Friend Sub InitVars()
            Me.columnДатаПоследнегоСохранения = Me.Columns("ДатаПоследнегоСохранения")
            Me.columnПапкаИдентификатор = Me.Columns("ПапкаИдентификатор")
            Me.columnСодержаниеФайла = Me.Columns("СодержаниеФайла")
            Me.columnСтрока = Me.Columns("Строка")
            Me.columnТипОбъекта = Me.Columns("ТипОбъекта")
            Me.columnФайлВПапке = Me.Columns("ФайлВПапке")
            Me.columnФайлИдентификатор = Me.Columns("ФайлИдентификатор")
        End Sub
        
        Private Sub InitClass()
            Me.columnДатаПоследнегоСохранения = New DataColumn("ДатаПоследнегоСохранения", GetType(System.DateTime), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnДатаПоследнегоСохранения)
            Me.columnПапкаИдентификатор = New DataColumn("ПапкаИдентификатор", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnПапкаИдентификатор)
            Me.columnСодержаниеФайла = New DataColumn("СодержаниеФайла", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnСодержаниеФайла)
            Me.columnСтрока = New DataColumn("Строка", GetType(System.Int32), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnСтрока)
            Me.columnТипОбъекта = New DataColumn("ТипОбъекта", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnТипОбъекта)
            Me.columnФайлВПапке = New DataColumn("ФайлВПапке", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnФайлВПапке)
            Me.columnФайлИдентификатор = New DataColumn("ФайлИдентификатор", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnФайлИдентификатор)
            Me.Constraints.Add(New UniqueConstraint("Constraint1", New DataColumn() {Me.columnСтрока}, true))
            Me.columnСтрока.AutoIncrement = true
            Me.columnСтрока.AllowDBNull = false
            Me.columnСтрока.Unique = true
        End Sub
        
        Public Function NewmanualRow() As manualRow
            Return CType(Me.NewRow,manualRow)
        End Function
        
        Protected Overrides Function NewRowFromBuilder(ByVal builder As DataRowBuilder) As DataRow
            Return New manualRow(builder)
        End Function
        
        Protected Overrides Function GetRowType() As System.Type
            Return GetType(manualRow)
        End Function
        
        Protected Overrides Sub OnRowChanged(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowChanged(e)
            If (Not (Me.manualRowChangedEvent) Is Nothing) Then
                RaiseEvent manualRowChanged(Me, New manualRowChangeEvent(CType(e.Row,manualRow), e.Action))
            End If
        End Sub
        
        Protected Overrides Sub OnRowChanging(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowChanging(e)
            If (Not (Me.manualRowChangingEvent) Is Nothing) Then
                RaiseEvent manualRowChanging(Me, New manualRowChangeEvent(CType(e.Row,manualRow), e.Action))
            End If
        End Sub
        
        Protected Overrides Sub OnRowDeleted(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowDeleted(e)
            If (Not (Me.manualRowDeletedEvent) Is Nothing) Then
                RaiseEvent manualRowDeleted(Me, New manualRowChangeEvent(CType(e.Row,manualRow), e.Action))
            End If
        End Sub
        
        Protected Overrides Sub OnRowDeleting(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowDeleting(e)
            If (Not (Me.manualRowDeletingEvent) Is Nothing) Then
                RaiseEvent manualRowDeleting(Me, New manualRowChangeEvent(CType(e.Row,manualRow), e.Action))
            End If
        End Sub
        
        Public Sub RemovemanualRow(ByVal row As manualRow)
            Me.Rows.Remove(row)
        End Sub
    End Class
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class manualRow
        Inherits DataRow
        
        Private tablemanual As manualDataTable
        
        Friend Sub New(ByVal rb As DataRowBuilder)
            MyBase.New(rb)
            Me.tablemanual = CType(Me.Table,manualDataTable)
        End Sub
        
        Public Property ДатаПоследнегоСохранения As Date
            Get
                Try 
                    Return CType(Me(Me.tablemanual.ДатаПоследнегоСохраненияColumn),Date)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Невозможно получить значение, т.к. оно является DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tablemanual.ДатаПоследнегоСохраненияColumn) = value
            End Set
        End Property
        
        Public Property ПапкаИдентификатор As String
            Get
                Try 
                    Return CType(Me(Me.tablemanual.ПапкаИдентификаторColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Невозможно получить значение, т.к. оно является DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tablemanual.ПапкаИдентификаторColumn) = value
            End Set
        End Property
        
        Public Property СодержаниеФайла As String
            Get
                Try 
                    Return CType(Me(Me.tablemanual.СодержаниеФайлаColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Невозможно получить значение, т.к. оно является DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tablemanual.СодержаниеФайлаColumn) = value
            End Set
        End Property
        
        Public Property Строка As Integer
            Get
                Return CType(Me(Me.tablemanual.СтрокаColumn),Integer)
            End Get
            Set
                Me(Me.tablemanual.СтрокаColumn) = value
            End Set
        End Property
        
        Public Property ТипОбъекта As String
            Get
                Try 
                    Return CType(Me(Me.tablemanual.ТипОбъектаColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Невозможно получить значение, т.к. оно является DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tablemanual.ТипОбъектаColumn) = value
            End Set
        End Property
        
        Public Property ФайлВПапке As String
            Get
                Try 
                    Return CType(Me(Me.tablemanual.ФайлВПапкеColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Невозможно получить значение, т.к. оно является DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tablemanual.ФайлВПапкеColumn) = value
            End Set
        End Property
        
        Public Property ФайлИдентификатор As String
            Get
                Try 
                    Return CType(Me(Me.tablemanual.ФайлИдентификаторColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Невозможно получить значение, т.к. оно является DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tablemanual.ФайлИдентификаторColumn) = value
            End Set
        End Property
        
        Public Function IsДатаПоследнегоСохраненияNull() As Boolean
            Return Me.IsNull(Me.tablemanual.ДатаПоследнегоСохраненияColumn)
        End Function
        
        Public Sub SetДатаПоследнегоСохраненияNull()
            Me(Me.tablemanual.ДатаПоследнегоСохраненияColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsПапкаИдентификаторNull() As Boolean
            Return Me.IsNull(Me.tablemanual.ПапкаИдентификаторColumn)
        End Function
        
        Public Sub SetПапкаИдентификаторNull()
            Me(Me.tablemanual.ПапкаИдентификаторColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsСодержаниеФайлаNull() As Boolean
            Return Me.IsNull(Me.tablemanual.СодержаниеФайлаColumn)
        End Function
        
        Public Sub SetСодержаниеФайлаNull()
            Me(Me.tablemanual.СодержаниеФайлаColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsТипОбъектаNull() As Boolean
            Return Me.IsNull(Me.tablemanual.ТипОбъектаColumn)
        End Function
        
        Public Sub SetТипОбъектаNull()
            Me(Me.tablemanual.ТипОбъектаColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsФайлВПапкеNull() As Boolean
            Return Me.IsNull(Me.tablemanual.ФайлВПапкеColumn)
        End Function
        
        Public Sub SetФайлВПапкеNull()
            Me(Me.tablemanual.ФайлВПапкеColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsФайлИдентификаторNull() As Boolean
            Return Me.IsNull(Me.tablemanual.ФайлИдентификаторColumn)
        End Function
        
        Public Sub SetФайлИдентификаторNull()
            Me(Me.tablemanual.ФайлИдентификаторColumn) = System.Convert.DBNull
        End Sub
    End Class
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class manualRowChangeEvent
        Inherits EventArgs
        
        Private eventRow As manualRow
        
        Private eventAction As DataRowAction
        
        Public Sub New(ByVal row As manualRow, ByVal action As DataRowAction)
            MyBase.New
            Me.eventRow = row
            Me.eventAction = action
        End Sub
        
        Public ReadOnly Property Row As manualRow
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