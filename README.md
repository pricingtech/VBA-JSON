# VBA-JSON

Conversão e análise de JSON para VBA (Excel no Windows e Mac, Access e outras aplicações do Office).
Este projeto evoluiu a partir do excelente projeto vba-json, com adições e melhorias feitas para resolver bugs e melhorar o desempenho (como parte do [VBA-Web](https://github.com/VBA-tools/VBA-Web)).

Testado no Excel 2013 para Windows e Excel 2011 para Mac, mas deve funcionar a partir do Excel 2007.

- Para suporte apenas no Windows, inclua uma referência para "Microsoft Scripting Runtime"
- Para suporte no Mac e Windows, inclua [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)

<a href="https://www.patreon.com/timhall">
  <img src="https://timhall.github.io/assets/donate-patreon@2x.png" width="217" alt="Donate">
</a>

# Exemplos

```vb
Dim Json As Object
Set Json = JsonConverter.ParseJson("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")

' Json("a") -> 123
' Json("b")(2) -> 2
' Json("c")("d") -> 456
Json("c")("e") = 789

Debug.Print JsonConverter.ConvertToJson(Json)
' -> "{"a":123,"b":[1,2,3,4],"c":{"d":456,"e":789}}"

Debug.Print JsonConverter.ConvertToJson(Json, Whitespace:=2)
' -> "{
'       "a": 123,
'       "b": [
'         1,
'         2,
'         3,
'         4
'       ],
'       "c": {
'         "d": 456,
'         "e": 789  
'       }
'     }"
```

```vb
' Advanced example: Read .json file and load into sheet (Windows-only)
' (add reference to Microsoft Scripting Runtime)
' {"values":[{"a":1,"b":2,"c": 3},...]}

Dim FSO As New FileSystemObject
Dim JsonTS As TextStream
Dim JsonText As String
Dim Parsed As Dictionary

' Read .json file
Set JsonTS = FSO.OpenTextFile("example.json", ForReading)
JsonText = JsonTS.ReadAll
JsonTS.Close

' Parse json to Dictionary
' "values" is parsed as Collection
' each item in "values" is parsed as Dictionary
Set Parsed = JsonConverter.ParseJson(JsonText)

' Prepare and write values to sheet
Dim Values As Variant
ReDim Values(Parsed("values").Count, 3)

Dim Value As Dictionary
Dim i As Long

i = 0
For Each Value In Parsed("values")
  Values(i, 0) = Value("a")
  Values(i, 1) = Value("b")
  Values(i, 2) = Value("c")
  i = i + 1
Next Value

Sheets("example").Range(Cells(1, 1), Cells(Parsed("values").Count, 3)) = Values
```

## Opções

VBA-JSON inclui algumas opções para personalizar a análise/conversão, se necessário:

- __UseDoubleForLargeNumbers__ (Default = `False`) VBA only stores 15 significant digits, so any numbers larger than that are truncated.
  This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits.
  By default, VBA-JSON will use `String` for numbers longer than 15 characters that contain only digits, use this option to use `Double` instead.
- __AllowUnquotedKeys__ (Default = `False`) The JSON standard requires object keys to be quoted (`"` or `'`), use this option to allow unquoted keys.
- __EscapeSolidus__ (Default = `False`) The solidus (`/`) is not required to be escaped, use this option to escape them as `\/` in `ConvertToJson`.

```VB.net
JsonConverter.JsonOptions.EscapeSolidus = True
```

## Instalação

1. Baixe a [latest release](https://github.com/VBA-tools/VBA-JSON/releases)
2. Importe `JsonConverter.bas` para o seu projeto (Abra o Editor VBA, `Alt + F11`; Arquivo > Importar Arquivo)
3. Adicione a referência/classe
   - Para Windows apenas, inclua uma referência para "Microsoft Scripting Runtime"
   - Para Windows e Mac, inclua [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)

## Recursos

- [Vídeo Tutorial (Red Stapler)](https://youtu.be/CFFLRmHsEAs)