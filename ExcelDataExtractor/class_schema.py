# -*- coding: utf-8 -*-

import json
import types

# 파이썬에서 enum을 쓸 수 있도록 함.
def enum(*sequential, **named):
    enums = dict(zip(sequential, range(len(sequential))), **named)
    reverse = dict((value, key) for key, value in enums.iteritems())
    enums['reverse_mapping'] = reverse
    return type('Enum', (), enums)

# 데이터 타입 enumeration
eDataType = enum( "Int", "Float", "String", "Table", "IntList", "StringList" )

# 엑셀 테이블의 각 컬럼에 해당하는 데이터형을 나타낸다.
class Schema :
	columnName = "A"
	fieldName = ""
	dataType = eDataType.String
	isUnique = False
	
	def SetDataType( self, dataTypeText, firstRowValue ) :
		if( dataTypeText == 1 ) :
			self.dataType = eDataType.Int
		elif( dataTypeText == 2 ) :
			self.dataType = eDataType.Float
		elif( dataTypeText == 0 ) :			
			self.dataType = eDataType.String
			if( self.IsArray( firstRowValue ) == True ) :
				jsonArray = json.loads( firstRowValue )				
				if( type( jsonArray[0] ) == int ) :
					self.dataType = eDataType.IntList
				else :
					self.dataType = eDataType.StringList
		elif( dataTypeText == 3 ) :
			self.dataType = eDataType.Table		
	
	def SetAsUniqueField( self ) :
		self.isUnique = True
	
	# type의 문자열 반환.
	def TypeEnumString( self ) :
		if( self.dataType == eDataType.Int ) :
			return "int"
		elif( self.dataType == eDataType.Float ) :
			return "float"
		elif( self.dataType == eDataType.String ) :
			return "string"
		elif( self.dataType == eDataType.Table ) :
			return "table"
		elif( self.dataType == eDataType.IntList ) :
			return "List<int>"
		elif( self.dataType == eDataType.StringList ) :
			return "List<string>"
	
	def InfoString( self ) :
		return "<Schema:%s>Field = %s, Type = %s" % ( self.columnName, self.fieldName, self.dataType )
		
	def IsArray( self, stringValue ) :				
		if( stringValue == None ) :			
			return False		
		stringValueType = type( stringValue )
		if( stringValueType is not str and stringValueType is not unicode ) :			
			return False		 
		
		arrayStringLength = len( stringValue )
		if( arrayStringLength <= 2 ) : # 빈 배열일 경우 처리 무시
			return False
		lastIndex = arrayStringLength - 1
		if( stringValue[0] == '[' and stringValue[lastIndex] == ']' ) :
			return True		
		return False
		
		
		
	

	

	

