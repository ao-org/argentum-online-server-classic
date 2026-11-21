# AGENTS.md

Guidance for AI coding agents working on the Argentum Online (AO20) VB6 server, classic branch.

## Project overview

- Legacy MMORPG, client-server architecture. Both server and client are written in VB6 and talk over a custom packet protocol.
- This repository is the **server** (classic branch): game logic, database, NPCs, combat systems.
- Write legacy-compatible VB6 only. Do NOT use VB.NET syntax.
- Prioritize clean, readable code with minimal risk of regression.
- VB6 reference: https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation

### Repositories

| Repository | Type | Access | URL |
|------------|------|--------|-----|
| **Assets** | Resources/Config | Private | https://github.com/ao-org/argentum-online-assets-classic |
| **Server** | VB6 Server Code | Public | https://github.com/ao-org/argentum-online-server-classic |
| **Client** | VB6 Client Code | Public | https://github.com/ao-org/argentum-online-client-classic |

- **Assets (Recursos)**: graphics, sounds, maps, configuration files, shared resources.
- **Server**: game server logic, database, NPCs, combat systems.
- **Client**: game client, UI, rendering, input handling.

### Repository layout

- `Codigo/`: server source (`.bas`, `.cls`, `.frm`).
- `Server.VBP`: VB6 project file.
- `Documentation/`: per-module docs.
- `docs/`: planning docs.
- `ScriptsDB/`: database scripts.
- `Recursos/config/`: server configuration.

## Build commands

Compiler: `C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe`

### Server

```cmd
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe" /make ^
/d UsarQueSocket=1:ConUpTime=1:AntiExternos=0:Lac=1:DEBUGGING=0:PYMMO=1:UNLOCK_CPU=1:DIRECT_PLAY=0 ^
/out vb6build.log ^
Server.VBP
```

### Client (in the client repository)

```cmd
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe" /make ^
/d Compresion=1:DEBUGGING=1:PYMMO=1:ENABLE_ANTICHEAT=1:REMOTE_CLOSE=1:BATTLESERVER=1:DIRECT_PLAY=0:ENABLE_BASS=1 ^
/out vb6build.log ^
Client.VBP
```

## Workflow for agents

1. Make code changes to `.bas`, `.cls` or `.frm` files.
2. Run the VB6 build command above.
3. Check `vb6build.log` for errors.
4. If errors exist, parse the log and attempt fixes.
5. Repeat until the build succeeds.

## Code style

### 1. Mandatory use of `Call` and parentheses

Always use `Call` when invoking any `Sub`. Always include parentheses, even if there are no arguments.

```vb
Call GuardarDatos()
Call EnviarMensaje("Hola", 2)
```

### 2. Functions always with parentheses

Even when ignoring the return value, functions must be called with parentheses.

```vb
Call ObtenerTiempoActual()
Dim puntos As Long
puntos = CalcularPuntos(usuarioId)
```

### 3. Naming conventions

| Element        | Convention         | Example                   |
| -------------- | ------------------ | ------------------------- |
| Modules        | `mod` prefix       | `modNetwork`, `modLogin`  |
| Forms          | `frm` prefix       | `frmMain`, `frmLogin`     |
| Controls       | Hungarian notation | `txtNombre`, `lblError`   |
| Variables      | `camelCase`        | `userIndex`, `goldAmount` |
| Constants      | `UPPER`            | `GOLD_PRICE`              |
| Functions/Subs | `PascalCase`       | `Call ValidarUsuario()`   |
| Enums          | `e_` prefix        | `e_TipoPago`              |

### 4. Spacing and formatting

- Use 4-space indentation.
- One blank line between procedures.
- Align related variable declarations:

```vb
Dim userGold   As Long
Dim userSilver As Long
Dim userName   As String
```

### 5. Avoid magic numbers

Declare all fixed values as constants in shared modules.

```vb
Public Const GOLD_PRICE As Long = 50000

If .Stats.GLD < GOLD_PRICE Then
    Call EscribirError("No tenés suficiente oro.")
End If
```

### 6. Standardized error handling

Always use `On Error GoTo Name_Err` with a label at the end of the `Sub`.

```vb
Private Sub ValidarSesion()
    On Error GoTo ValidarSesion_Err

    Call HacerAlgo()
    Exit Sub

ValidarSesion_Err:
    Call TraceError(Err.Number, Err.Description, "modLogin.ValidarSesion", Erl)
End Sub
```

### 7. Explicit identifiers

Use clear, specific names. Avoid generic names like `dato`, `res`, `temp`.

```vb
Dim creditAmount As Long
Dim connectionId As Integer
```

### 8. SQL queries

Always use parameterized queries (`?`). Always close all `Recordset` objects.

```vb
Dim RS As ADODB.Recordset
Set RS = Query("SELECT nivel FROM user WHERE id = ?;", userId)

If Not RS.EOF Then
    nivel = RS!nivel
End If

Call RS.Close
```

### 9. Localized messages

#### Server-side (`WriteLocaleMsg`)

The server sends a message ID; the client resolves it from a message index file.

```vb
Call WriteLocaleMsg(UserIndex, "1291", FONTTYPE_INFOBOLD, GOLD_PRICE)
```

`"1291"` maps to: *"You need at least ¬1 gold to sell your character"*. `GOLD_PRICE` replaces `¬1`.

Message files:
- [Spanish: SP_LocalMsg.dat](https://github.com/ao-org/Recursos/blob/master/init/SP_LocalMsg.dat)
- [English: EN_LocalMsg.dat](https://github.com/ao-org/Recursos/blob/master/init/EN_LocalMsg.dat)

To regenerate: run `python generar_localindex.py` or `generar_localindex.exe` from `/tools/`.

#### Client-side (`JsonLanguage`)

Client-side UI translation uses JSON files in the [`Languages`](https://github.com/ao-org/argentum-online-client/tree/master/Languages) folder:
- [Languages/1.json (Spanish)](https://github.com/ao-org/argentum-online-client/blob/master/Languages/1.json)
- [Languages/2.json (English)](https://github.com/ao-org/argentum-online-client/blob/master/Languages/2.json)

```vb
Call MsgBox(JsonLanguage.Item("MENSAJE_ERROR_CARGAR_OPCIONES"), vbCritical, JsonLanguage.Item("TITULO_ERROR_CARGAR"))
```

### 10. Clear control flow

Use `Exit Sub` for early exits. Avoid unnecessary nesting. Avoid `GoTo` flow control except for error handling, or unless absolutely necessary for performance.

```vb
If Not EstaAutenticado(UserIndex) Then
    Call Desconectar(UserIndex)
    Exit Sub
End If
```

### 11. Pure functions and no side effects

Prefer `Function` for validations or transformations. Avoid modifying global variables unnecessarily.

### 12. File encoding

All source code files must be saved with Windows-1252 encoding for VB6 compatibility. Do not re-save them as UTF-8.

### 13. Function parameters

Avoid optional or default parameters and named-argument syntax:

```vb
Public Function GetUserName(ByVal userId As Long = 0, Optional Something As Boolean = False) As String
End Function
Call GetUserName(userId, Something:= False)
```

Use simple `ByVal` parameters unless necessary (for example, when refactoring legacy code that uses `ByRef`).

### 14. Helper functions

`Matematicas.bas` holds mathematical helper functions. Reuse them before writing new ones.


# Running tests

"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make /d DEBUGGING=1:PYMM0=0:UNIT_TEST=1 /out vb6build.log Server.VBP
start "" Server.exe

Wait ~5 seconds, then:

type test_results.txt
taskkill /f /im Server.exe >nul 2>&1

test_results.txt will contain RESULT: PASS or RESULT: FAIL with details.


Writing tests

Create Codigo/Tests/Unit_YourFeature.bas. Template:

Attribute VB_Name = "Unit_YourFeature"
```
Public Sub RunTests()
    Call UnitTesting.StartSuite("YourFeature")

    Call UnitTesting.AssertEqual(2 + 2, 4, "basic math")
    Call UnitTesting.AssertTrue(SomeFunction() = "expected", "SomeFunction returns expected")

    Call UnitTesting.EndSuite
End Sub
```

Then register it in UnitTesting.bas – find where other suites are called (search for Unit_Math.RunTests) and add your suite in the same place.

Also add the module to Server.VBP:
```
Module=Unit_YourFeature; Codigo\Tests\Unit_YourFeature.bas
```
