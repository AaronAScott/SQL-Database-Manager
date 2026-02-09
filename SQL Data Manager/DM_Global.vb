Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Module DM_Global
	'**********************************************************
	' SQL Data Manager Public Module
	' DM_GLOBAL.BAS
	' Written: October 2017
	' Programmer: Aaron Scott
	' Copyright (C) 2017 Sirius Software All Rights Reserved
	'**********************************************************

	' Define a structure for passing Foreign Key information to the NewRow routine.

	Public Structure ForeignKeyValues
		Dim RelationshipName As String
		Dim PrimaryTable As String
		Dim PrimaryField As String
		Dim RelatedTable As String
		Dim RelatedField As String
		Dim CascadeUpdate As Boolean
		Dim CascadeDelete As Boolean
	End Structure

	' Define a flag that will allow us to prevent MDI child forms from
	' being automatically moved/resized when the main form is moved
	' or resized.

	Public DontAllowResize As Boolean
End Module
