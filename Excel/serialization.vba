'Take a dictionary of key/value pairs to process as a serialized data string
Function SerializeDictionary(dictionary)

Dim key As Variant
Dim value
Dim returnString As String

'Init returnString with {
returnString = "{"

'Iterate over the dictionary of data
For Each key In dictionary
    'Get the current value
    value = dictionary.Item(key)
    'Add the serialized pair to the return string
    returnString = returnString + SerializeKeyValue(key, value)
Next

'Close returnString with }
returnString = returnString + "}"

'Return serialized string
SerializeDictionary = returnString

End Function

'===============================

'Convert a key/value pair into a serialized format
Function SerializeKeyValue(key, value)

Dim returnString As String
Dim keyLength As Integer
Dim valLength As Integer

'The key is always a string
If (VarType(key) = vbString) Then
    'Do not output if key is an empty string (used for arrays)
    If (key <> "") Then
        keyLength = Len(key)
        returnString = "s:" + CStr(keyLength) + ":""" + key + """;"
    End If
End If

'Convert the value based on type
If (VarType(value) = vbInteger) Then
    'Integers are the value prefixed with i:
    returnString = returnString + "i:" + CStr(value) + ";"
ElseIf (VarType(value) = vbString) Then
    'If the value is an empty array (passed in as a string, special case)
    If (value = "[]") Then
        returnString = returnString + "a:0:{}"
    Else
        valLength = Len(value)
        returnString = returnString + "s:" + CStr(valLength) + ":" + """" + value + """;"
    End If
ElseIf (VarType(value) = 8204) Then
    'Value is an array - process each item
    returnString = returnString + "a:" + CStr(GetLength(value)) + ":{"
    'Iterate over elements in the value array
    For i = LBound(value) To UBound(value)
        'Get the current value
        arrValue = value(i)
        'Debug.Print "PROCESSING ARRAY VALUE FOR " + key + ": " + CStr(arrValue) + " OF TYPE " + CStr(VarType(arrValue)) + " => " + SerializeKeyValue("", arrValue)
        returnString = returnString + "i:" + CStr(i) + ";" + SerializeKeyValue("", arrValue)
    Next
    returnString = returnString + "}"
End If

'Set as the function name to return
SerializeKeyValue = returnString

End Function

'===============================

Sub SerializeExample()

Dim d                   'Create a variable
Set d = CreateObject("Scripting.Dictionary")
Dim productData(1) As Variant

d.Add "product", 140
'productsData = Array(124, 764)
productsData = Array(124)
d.Add "products", productsData
d.Add "orderby", "date"
d.Add "order", "DESC"
d.Add "quantity", 1
d.Add "min_quantity", 1
d.Add "max_quantity", ""
d.Add "discount_type", "percentage"
d.Add "discount", ""
d.Add "title", ""
d.Add "description", ""
d.Add "select_product_title", "Please select a product!"
d.Add "product_list_title", "Please select your product!"
d.Add "modal_header_title", "Please select your product!"
d.Add "optional", "false"
d.Add "selected", "false"
d.Add "excluded_products", "[]"
d.Add "categories", "[]"
d.Add "excluded_categories", "[]"
d.Add "tags", "[]"
d.Add "excluded_tags", "[]"
d.Add "query_relation", "OR"
d.Add "edit_quantity", "false"
d.Add "image_url", ""

Dim key As Variant
Dim value
Dim returnString As String

'Init returnString
returnString = ""

'Iterate over the dictionary of data
For Each key In d
   value = d.Item(key)
   returnString = returnString + SerializeKeyValue(key, value)
Next

MsgBox returnString, Title = "FullData"

End Sub

'======================================

'Source: https://stackoverflow.com/a/53581681
Public Function GetLength(a As Variant) As Integer
   If IsEmpty(a) Then
      GetLength = 0
   Else
      GetLength = UBound(a) - LBound(a) + 1
   End If
End Function
