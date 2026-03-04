import zipfile
import os
import io
import win32com.client as win32
from pathlib import Path

def crear_excel_con_macros():
    """
    Crea un archivo Excel con macros para editar productos JSON
    """
    # Crear una instancia de Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    
    try:
        # Crear nuevo libro de trabajo
        workbook = excel.Workbooks.Add()
        sheet = workbook.Sheets(1)
        
        # Configurar la hoja
        sheet.Name = "Productos"
        
        # Agregar encabezados
        headers = ["id", "nombreKey", "nombre", "descripcion", "precioUnidad", 
                  "precioCaja", "categoria", "etiqueta", "imagen"]
        
        for i, header in enumerate(headers, 1):
            sheet.Cells(1, i).Value = header
            # Formato de encabezados
            sheet.Cells(1, i).Font.Bold = True
            sheet.Cells(1, i).Interior.Color = 12632256  # Gris claro
        
        # Ajustar columnas
        sheet.Columns("A:I").AutoFit()
        
        # Agregar notas instructivas
        sheet.Range("A3").Value = "INSTRUCCIONES:"
        sheet.Range("A4").Value = "1. Usa los botones de arriba para importar/exportar JSON"
        sheet.Range("A5").Value = "2. Las categorías y etiquetas tienen lista desplegable"
        sheet.Range("A6").Value = "3. Las imágenes deben estar en ./assets/img/"
        sheet.Range("A7").Value = "4. Guarda como .xlsm para mantener las macros"
        
        sheet.Range("A3:A7").Font.Bold = True
        sheet.Range("A3:A7").Font.Color = 255  # Rojo
        
        # Agregar ejemplos de productos (primeros 5)
        productos_ejemplo = [
            [1, "prod_mermelada_mora", "Mermelada de Mora", "Mermelada artesanal, 12 unidades por caja", 
             4.50, 54.00, "Mermeladas", "ARTESANAL", "./assets/img/mermelada_mora.png"],
            [2, "prod_mermelada_alcayota", "Mermelada de Alcayota", "Tradicional alcayota, 24 unidades por caja",
             4.50, 108.00, "Mermeladas", "TRADICIÓN", "./assets/img/mermelada_alcayota.png"],
            [3, "prod_mermelada_damasco", "Mermelada de Damasco", "Damasco seleccionado, 12 unidades por caja",
             4.50, 54.00, "Mermeladas", "ARTESANAL", "./assets/img/mermelada_damasco.png"],
            [4, "prod_mermelada_durazno", "Mermelada de Durazno", "Durazno fresco, 12 unidades por caja",
             4.50, 54.00, "Mermeladas", "ARTESANAL", "./assets/img/mermelada_durazno.png"],
            [5, "prod_mermelada_mango", "Mermelada de Mango", "Mango tropical, 12 unidades por caja",
             4.50, 54.00, "Mermeladas", "EXÓTICO", "./assets/img/mermelada_mango.png"]
        ]
        
        for row_idx, producto in enumerate(productos_ejemplo, start=10):
            for col_idx, valor in enumerate(producto, start=1):
                sheet.Cells(row_idx, col_idx).Value = valor
        
        # Crear rangos con nombre para validaciones
        categorias = ["Mermeladas", "Galletas", "Chocolates", "Salsas", "Despensa", "Bebidas"]
        etiquetas = ["ARTESANAL", "TRADICIÓN", "EXÓTICO", "PREMIUM", "POPULAR", "CLÁSICO", 
                    "ECONÓMICO", "ESPECIAL", "INFANTIL", "PICANTE", "TÍPICO", "REFRESCANTE", "SOPAS"]
        
        # Agregar validaciones en una hoja oculta
        config_sheet = workbook.Sheets.Add()
        config_sheet.Name = "Config"
        config_sheet.Visible = -1  # xlSheetHidden
        
        for i, cat in enumerate(categorias, 1):
            config_sheet.Cells(i, 1).Value = cat
        
        for i, etiq in enumerate(etiquetas, 1):
            config_sheet.Cells(i, 2).Value = etiq
        
        # Crear los rangos con nombre
        workbook.Names.Add("CategoriasList", config_sheet.Range("A1:A" & len(categorias)))
        workbook.Names.Add("EtiquetasList", config_sheet.Range("B1:B" & len(etiquetas)))
        
        # Volver a la hoja de productos
        sheet.Activate()
        
        # Agregar el código VBA
        vba_code = '''
Option Explicit

Sub ImportarJSON()
    Dim filePath As Variant
    Dim jsonText As String
    Dim jsonObj As Object
    Dim productos As Object
    Dim i As Long
    Dim ws As Worksheet
    
    filePath = Application.GetOpenFilename("Archivos JSON (*.json),*.json", , "Selecciona el archivo JSON de productos")
    If filePath = False Then Exit Sub
    
    Open filePath For Input As #1
    jsonText = Input$(LOF(1), 1)
    Close #1
    
    Set jsonObj = JsonConverter.ParseJson(jsonText)
    
    Set ws = ThisWorkbook.Sheets("Productos")
    
    ' Limpiar datos existentes (excepto encabezados y ejemplos)
    ws.Rows("10:1000").ClearContents
    
    ' Cargar datos
    For i = 0 To jsonObj.Count - 1
        Set productos = jsonObj(i)
        ws.Range("A" & i + 10) = productos("id")
        ws.Range("B" & i + 10) = productos("nombreKey")
        ws.Range("C" & i + 10) = productos("nombre")
        ws.Range("D" & i + 10) = productos("descripcion")
        ws.Range("E" & i + 10) = productos("precioUnidad")
        ws.Range("F" & i + 10) = productos("precioCaja")
        ws.Range("G" & i + 10) = productos("categoria")
        ws.Range("H" & i + 10) = productos("etiqueta")
        ws.Range("I" & i + 10) = productos("imagen")
    Next i
    
    AplicarValidaciones
    
    MsgBox "¡Productos importados correctamente!" & vbNewLine & _
           "Total: " & jsonObj.Count & " productos", vbInformation
End Sub

Sub ExportarJSON()
    Dim filePath As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim jsonText As String
    Dim ws As Worksheet
    
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="productos_actualizado.json", _
        fileFilter:="Archivos JSON (*.json),*.json", _
        Title:="Guardar archivo JSON")
    
    If filePath = False Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("Productos")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 10 Then
        MsgBox "No hay datos para exportar", vbExclamation
        Exit Sub
    End If
    
    jsonText = "[" & vbNewLine
    
    For i = 10 To lastRow
        If ws.Range("A" & i).Value <> "" Then
            jsonText = jsonText & "    {" & vbNewLine
            jsonText = jsonText & "        ""id"": " & ws.Range("A" & i).Value & "," & vbNewLine
            jsonText = jsonText & "        ""nombreKey"": """ & EscapeJSON(ws.Range("B" & i).Value) & """," & vbNewLine
            jsonText = jsonText & "        ""nombre"": """ & EscapeJSON(ws.Range("C" & i).Value) & """," & vbNewLine
            jsonText = jsonText & "        ""descripcion"": """ & EscapeJSON(ws.Range("D" & i).Value) & """," & vbNewLine
            jsonText = jsonText & "        ""precioUnidad"": " & ws.Range("E" & i).Value & "," & vbNewLine
            jsonText = jsonText & "        ""precioCaja"": " & ws.Range("F" & i).Value & "," & vbNewLine
            jsonText = jsonText & "        ""categoria"": """ & EscapeJSON(ws.Range("G" & i).Value) & """," & vbNewLine
            jsonText = jsonText & "        ""etiqueta"": """ & EscapeJSON(ws.Range("H" & i).Value) & """," & vbNewLine
            jsonText = jsonText & "        ""imagen"": """ & EscapeJSON(ws.Range("I" & i).Value) & """"
            jsonText = jsonText & "    }"
            
            If i < lastRow Then
                If ws.Range("A" & i + 1).Value <> "" Then
                    jsonText = jsonText & ","
                End If
            End If
            jsonText = jsonText & vbNewLine
        End If
    Next i
    
    jsonText = jsonText & "]"
    
    Open filePath For Output As #1
    Print #1, jsonText
    Close #1
    
    MsgBox "¡JSON exportado correctamente!" & vbNewLine & _
           "Total: " & (lastRow - 9) & " productos", vbInformation
End Sub

Sub AplicarValidaciones()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Productos")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 10 Then lastRow = 1000
    
    ' Validación para categorías
    With ws.Range("G10:G" & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=CategoriasList"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' Validación para etiquetas
    With ws.Range("H10:H" & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=EtiquetasList"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Function EscapeJSON(texto As String) As String
    Dim temp As String
    If IsNull(texto) Then
        EscapeJSON = ""
        Exit Function
    End If
    
    temp = Replace(texto, "\", "\\")
    temp = Replace(temp, """", "\""")
    temp = Replace(temp, vbCrLf, "\n")
    temp = Replace(temp, vbCr, "\n")
    temp = Replace(temp, vbLf, "\n")
    EscapeJSON = temp
End Function

Sub NuevoProducto()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim newRow As Long
    
    Set ws = ThisWorkbook.Sheets("Productos")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    newRow = lastRow + 1
    
    ' Encontrar el máximo ID
    Dim maxID As Long
    maxID = Application.WorksheetFunction.Max(ws.Range("A10:A" & lastRow))
    
    ws.Range("A" & newRow) = maxID + 1
    ws.Range("B" & newRow) = "prod_nuevo"
    ws.Range("C" & newRow) = "Nuevo Producto"
    ws.Range("I" & newRow) = "./assets/img/producto_default.png"
    
    ws.Range("A" & newRow & ":I" & newRow).Select
    MsgBox "Nuevo producto agregado en fila " & newRow, vbInformation
End Sub

Sub ActualizarRutasImagen()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim contador As Long
    
    Set ws = ThisWorkbook.Sheets("Productos")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    contador = 0
    
    For i = 10 To lastRow
        If ws.Range("I" & i).Value = "" Or InStr(ws.Range("I" & i).Value, "./assets/img/") = 0 Then
            If ws.Range("B" & i).Value <> "" Then
                Dim nombreKey As String
                nombreKey = Replace(ws.Range("B" & i).Value, "prod_", "")
                ws.Range("I" & i).Value = "./assets/img/" & nombreKey & ".png"
                contador = contador + 1
            End If
        End If
    Next i
    
    MsgBox "Rutas actualizadas para " & contador & " productos", vbInformation
End Sub

Sub CrearBotonesInterfaz()
    Dim ws As Worksheet
    Dim btn As Object
    
    Set ws = ThisWorkbook.Sheets("Productos")
    
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    ' Crear botones en la fila 2
    With ws
        Set btn = .Buttons.Add(.Range("A2").Left, .Range("A2").Top, 120, 25)
        btn.OnAction = "ImportarJSON"
        btn.Caption = "📂 Importar JSON"
        
        Set btn = .Buttons.Add(.Range("B2").Left, .Range("B2").Top, 120, 25)
        btn.OnAction = "ExportarJSON"
        btn.Caption = "💾 Exportar JSON"
        
        Set btn = .Buttons.Add(.Range("C2").Left, .Range("C2").Top, 120, 25)
        btn.OnAction = "NuevoProducto"
        btn.Caption = "➕ Nuevo Producto"
        
        Set btn = .Buttons.Add(.Range("D2").Left, .Range("D2").Top, 150, 25)
        btn.OnAction = "ActualizarRutasImagen"
        btn.Caption = "🖼️ Actualizar Imágenes"
        
        Set btn = .Buttons.Add(.Range("E2").Left, .Range("E2").Top, 120, 25)
        btn.OnAction = "AplicarValidaciones"
        btn.Caption = "✓ Validar Datos"
    End With
End Sub

Sub Workbook_Open()
    CrearBotonesInterfaz
    AplicarValidaciones
End Sub
'''
        
        # Agregar el módulo VBA
        vba_module = workbook.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
        vba_module.Name = "ProductosModule"
        vba_module.CodeModule.AddFromString(vba_code)
        
        # Crear botones automáticamente
        excel.Run("CrearBotonesInterfaz")
        
        # Guardar el libro con macros
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        file_path = os.path.join(desktop, "EditorProductos.xlsm")
        
        workbook.SaveAs(file_path, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
        print(f"✅ Archivo Excel creado exitosamente en: {file_path}")
        
        # Crear archivo JSON de ejemplo
        json_path = os.path.join(desktop, "productos_ejemplo.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            f.write('''[
  {
    "id": 1,
    "nombreKey": "prod_mermelada_mora",
    "nombre": "Mermelada de Mora",
    "descripcion": "Mermelada artesanal, 12 unidades por caja",
    "precioUnidad": 4.5,
    "precioCaja": 54.0,
    "categoria": "Mermeladas",
    "etiqueta": "ARTESANAL",
    "imagen": "./assets/img/mermelada_mora.png"
  },
  {
    "id": 2,
    "nombreKey": "prod_mermelada_alcayota",
    "nombre": "Mermelada de Alcayota",
    "descripcion": "Tradicional alcayota, 24 unidades por caja",
    "precioUnidad": 4.5,
    "precioCaja": 108.0,
    "categoria": "Mermeladas",
    "etiqueta": "TRADICIÓN",
    "imagen": "./assets/img/mermelada_alcayota.png"
  }
]''')
        print(f"✅ Archivo JSON de ejemplo creado en: {json_path}")
        
        # Crear archivo README
        readme_path = os.path.join(desktop, "INSTRUCCIONES.txt")
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write("""INSTRUCCIONES PARA USAR EL EDITOR DE PRODUCTOS
=============================================

1. ARCHIVOS INCLUIDOS:
   - EditorProductos.xlsm : Archivo Excel con macros para editar productos
   - productos_ejemplo.json : Archivo JSON de ejemplo para comenzar

2. PRIMEROS PASOS:
   - Abre el archivo "EditorProductos.xlsm"
   - Si aparece advertencia de seguridad, habilita las macros
   - Verás botones en la fila 2 y ejemplos desde la fila 10

3. CÓMO USAR:
   a) IMPORTAR: Haz clic en "Importar JSON" y selecciona tu archivo JSON
   b) EDITAR: Modifica los datos directamente en Excel
   c) AGREGAR: Usa "Nuevo Producto" para agregar filas
   d) VALIDAR: Usa "Validar Datos" para actualizar listas desplegables
   e) EXPORTAR: Haz clic en "Exportar JSON" para guardar los cambios

4. CAMPOS IMPORTANTES:
   - imagen: Debe comenzar con "./assets/img/" (ej: ./assets/img/mermelada_mora.png)
   - categoría y etiqueta: Tienen lista desplegable para evitar errores

5. NOTAS:
   - Guarda siempre como .xlsm para mantener las macros
   - Los IDs se generan automáticamente
   - Las imágenes deben existir en la carpeta assets/img/

¡A EDITAR PRODUCTOS! 🚀
""")
        print(f"✅ Archivo de instrucciones creado en: {readme_path}")
        
        # Crear archivo ZIP
        zip_path = os.path.join(desktop, "EditorProductos.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(file_path, arcname="EditorProductos.xlsm")
            zipf.write(json_path, arcname="productos_ejemplo.json")
            zipf.write(readme_path, arcname="INSTRUCCIONES.txt")
        
        print(f"✅ Archivo ZIP creado en: {zip_path}")
        print("\n📦 TODO LISTO! Los archivos están en tu escritorio:")
        print(f"   - {file_path}")
        print(f"   - {json_path}")
        print(f"   - {readme_path}")
        print(f"   - {zip_path}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        workbook.Close(SaveChanges=False)
        excel.Quit()

if __name__ == "__main__":
    print("🚀 Creando editor de productos Excel...")
    crear_excel_con_macros()