Attribute VB_Name = "Módulo1"
' iniciar (obtener handle) y terminar(liberar handle)
Declare Function ParseInit Lib "parse32.dll" () As Long
Declare Sub ParseEnd Lib "parse32.dll" (ByVal handle As Long)
' resetear ( borra variables),  y determinar case-sensitive
Declare Sub ParseReset Lib "parse32.dll" (ByVal handle As Long, ByVal flags As Long)
' obtener el numero de error y un texto explicativo
Declare Function GetError Lib "parse32.dll" (ByVal handle As Long, ByVal texto_del_error As String) As Integer
' para saber si una variable se encuentra definida.
Declare Function DefValue Lib "parse32.dll" (ByVal handle As Long, ByVal nombre_variable As String) As Integer
' eliminar una variable (la convierte en indefinida)
Declare Sub DelValue Lib "parse32.dll" (ByVal handle As Long, ByVal nombre_variable As String)
' analizar una expresion, devuelve 0 si ha habido algun error
Declare Function GetValue Lib "parse32.dll" (ByVal handle As Long, ByVal expresion As String, valor_devuelto As Double) As Integer
' asignar un valor a una variable
Declare Sub SetValue Lib "parse32.dll" (ByVal handle As Long, ByVal nombre_variable As String, ByVal valor_asignar As Double)
' asignar un valor constante
Declare Sub SetValueCte Lib "parse32.dll" (ByVal handle As Long, ByVal nombre_variable As String, ByVal valor_asignar As Double)
' asignar una expresion o valor en formato texto
Declare Sub SetValueTxt Lib "parse32.dll" (ByVal handle As Long, ByVal nombre_variable As String, ByVal valor_asignar As String)

