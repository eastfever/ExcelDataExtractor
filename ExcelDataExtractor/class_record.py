# -*- coding: utf-8 -*-

from class_schema import eDataType
from class_schema import Schema

# 엑셀 테이블에서 추출한 데이터 단위.
class RecordField :
	dataType = eDataType.Int
	fieldName = "Unknown"
	value = ""

class Record :
	fieldTable = {}
	
	def __init__( self ) :
		self.fieldTable = {}
	
	def AddField( self, schema, value ) :		
		
		# 테이블 필드는 추가하지 않는다.
		if( schema.dataType == eDataType.Table ) :
			return
		
		newRecordField = RecordField()
		newRecordField.dataType = schema.dataType
		newRecordField.fieldName = schema.fieldName
		
		newRecordField.value = value		
		if( schema.dataType == eDataType.Float ) :
			newRecordField.value *= 1.0 # 강제로 소수점이 되도록 하기.
		if( schema.dataType == eDataType.String ) : # 강제로 문자열 만들기.
			newRecordField.value = str( newRecordField.value )
		self.fieldTable[schema.fieldName] = newRecordField
		
	def PrintMe( self ) :		
		logString = ""		
		for field in self.fieldTable.values() :
			logString += field.fieldName
			logString += "("
			logString += str( field.dataType )
			logString += ")"
			logString += field.value
			logString += ", "		
		print logString