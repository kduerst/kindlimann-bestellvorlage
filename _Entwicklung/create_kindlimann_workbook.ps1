param(
    [string]$OutputFileName = "Metallraum_Kindlimann_Sammelbestellung.xlsm"
)

$ErrorActionPreference = "Stop"

function Get-ConfigValue {
    param(
        [hashtable]$Source,
        [string]$Name,
        $Default
    )

    if ($Source -and $Source.ContainsKey($Name) -and $null -ne $Source[$Name]) {
        return $Source[$Name]
    }
    return $Default
}

function ConvertTo-VbaStringValue {
    param([string]$Value)
    return ($Value -replace '"', '""')
}

function ConvertTo-InvariantNumber {
    param([double]$Value)
    return $Value.ToString([System.Globalization.CultureInfo]::InvariantCulture)
}

$configPath = Join-Path $PSScriptRoot "Config\kindlimann.settings.psd1"
$settings = @{}
if (Test-Path -LiteralPath $configPath) {
    $settings = Import-PowerShellDataFile -Path $configPath
}

if (-not $PSBoundParameters.ContainsKey("OutputFileName")) {
    $OutputFileName = Get-ConfigValue -Source $settings -Name "OutputFileName" -Default $OutputFileName
}

$targetDir = Get-ConfigValue -Source $settings -Name "TargetDir" -Default "C:\Meine Programme\Kindlimann"
$targetFile = Join-Path $targetDir $OutputFileName
$logoPath = Get-ConfigValue -Source $settings -Name "LogoPath" -Default "C:\Meine Programme\Metallraum Logo.jpg"
$password = Get-ConfigValue -Source $settings -Name "Password" -Default "9500Wil"
$minimumStartDateRaw = Get-ConfigValue -Source $settings -Name "MinimumStartDate" -Default "2026-06-29"
$minimumStartDate = [datetime]::Parse([string]$minimumStartDateRaw, [System.Globalization.CultureInfo]::InvariantCulture)
$profiles = @(Get-ConfigValue -Source $settings -Name "ProfileList" -Default @(
    "40 x 15 x 2 mm",
    "FL 30 x 6 mm",
    "FL 40 x 10 mm",
    "FL 40 x 15 mm",
    "RD 10 mm"
))

$mailSettings = Get-ConfigValue -Source $settings -Name "Mail" -Default @{}
$mailTo = Get-ConfigValue -Source $mailSettings -Name "To" -Default "claudia.desantis@kindlimann.ch"
$mailSalutation = Get-ConfigValue -Source $mailSettings -Name "Salutation" -Default "Guten Tag Frau De Santis"
$mailSubjectPrefix = Get-ConfigValue -Source $mailSettings -Name "SubjectPrefix" -Default "Sammelbestellung Metallraum"
$mailClosing = Get-ConfigValue -Source $mailSettings -Name "Closing" -Default "Freundliche Grüsse"
$mailSignature = Get-ConfigValue -Source $mailSettings -Name "Signature" -Default "Metallraum"

$layoutSettings = Get-ConfigValue -Source $settings -Name "Layout" -Default @{}
$columnSettings = Get-ConfigValue -Source $layoutSettings -Name "Columns" -Default @{}
$rowSettings = Get-ConfigValue -Source $layoutSettings -Name "Rows" -Default @{}
$columnWidthA = [double](Get-ConfigValue -Source $columnSettings -Name "A" -Default 25)
$columnWidthB = [double](Get-ConfigValue -Source $columnSettings -Name "B" -Default 60)
$columnWidthC = [double](Get-ConfigValue -Source $columnSettings -Name "C" -Default 20)
$columnWidthD = [double](Get-ConfigValue -Source $columnSettings -Name "D" -Default 48)
$rowHeader = [double](Get-ConfigValue -Source $rowSettings -Name "Header" -Default 64)
$rowMeta = [double](Get-ConfigValue -Source $rowSettings -Name "Meta" -Default 28)
$rowStatus = [double](Get-ConfigValue -Source $rowSettings -Name "Status" -Default 22)
$rowSpacer = [double](Get-ConfigValue -Source $rowSettings -Name "Spacer" -Default 16)
$rowOrderTitle = [double](Get-ConfigValue -Source $rowSettings -Name "OrderTitle" -Default 28)
$rowOrderInfo = [double](Get-ConfigValue -Source $rowSettings -Name "OrderInfo" -Default 32)
$rowProfileHeader = [double](Get-ConfigValue -Source $rowSettings -Name "ProfileHeader" -Default 24)
$rowProfile = [double](Get-ConfigValue -Source $rowSettings -Name "Profile" -Default 27)
$rowBlank = [double](Get-ConfigValue -Source $rowSettings -Name "Blank" -Default 18)
$buttonGapPoints = [double](Get-ConfigValue -Source $layoutSettings -Name "ButtonGapPoints" -Default 24)
$pdfMarginCm = [double](Get-ConfigValue -Source $layoutSettings -Name "PdfMarginCm" -Default 2)
$pdfMarginPoints = $pdfMarginCm * 28.3464567

function Set-Cell {
    param($Sheet, [string]$Address, $Value)
    $Sheet.Range($Address).Value2 = $Value
}

function XlColor {
    param([int]$R, [int]$G, [int]$B)
    return $R + ($G * 256) + ($B * 65536)
}

function Get-ImageRatio {
    param([string]$Path)

    Add-Type -AssemblyName System.Drawing
    $image = [System.Drawing.Image]::FromFile($Path)
    try {
        return [double]$image.Width / [double]$image.Height
    }
    finally {
        $image.Dispose()
    }
}

function Format-ProfileRow {
    param($Sheet, [int]$Row)

    $Sheet.Range("A${Row}:D${Row}").Interior.Color = (XlColor 255 255 255)
    $Sheet.Range("A${Row}:D${Row}").Borders.LineStyle = 1
    $Sheet.Range("A${Row}:D${Row}").Borders.Weight = 2
    $Sheet.Range("A${Row}:D${Row}").Borders.Color = (XlColor 217 220 224)
    $Sheet.Range("A${Row}:D${Row}").Font.Name = "Aptos"
    $Sheet.Range("A${Row}:D${Row}").Font.Size = 10
    $Sheet.Range("A${Row}").HorizontalAlignment = -4108
    $Sheet.Range("B${Row}").HorizontalAlignment = -4131
    $Sheet.Range("C${Row}").HorizontalAlignment = -4108
    $Sheet.Range("B${Row}").WrapText = $true
    $Sheet.Range("C${Row}").ShrinkToFit = $true
    $Sheet.Range("A${Row}:C${Row}").Locked = $false
    $Sheet.Range("D${Row}").Locked = $true
    $Sheet.Range("C${Row}").Value2 = "6000 mm"
    $Sheet.Range("E${Row}").Value2 = "PROFILE"
    $Sheet.Range("A${Row}").NumberFormat = "0"
    $Sheet.Range("B${Row}").Validation.Delete()
    $Sheet.Range("B${Row}").Validation.Add(3, 1, 1, "=ProfilListe")
    $Sheet.Range("B${Row}").Validation.IgnoreBlank = $true
    $Sheet.Range("B${Row}").Validation.InCellDropdown = $true
    $Sheet.Range("B${Row}").Validation.ShowError = $false
    $Sheet.Rows("${Row}:${Row}").RowHeight = $rowProfile
}

function Add-OrderBlock {
    param($Sheet, [int]$StartRow, [int]$OrderNumber)

    $titleRow = $StartRow
    $infoRow = $StartRow + 1
    $headerRow = $StartRow + 2
    $profileRow = $StartRow + 3

    $Sheet.Range("A${titleRow}:D${titleRow}").Merge()
    $Sheet.Range("A${titleRow}").Value2 = "Auftrag $OrderNumber"
    $Sheet.Range("A${titleRow}:D${titleRow}").Interior.Color = (XlColor 28 28 28)
    $Sheet.Range("A${titleRow}:D${titleRow}").Font.Color = (XlColor 255 255 255)
    $Sheet.Range("A${titleRow}:D${titleRow}").Font.Bold = $true
    $Sheet.Range("A${titleRow}:D${titleRow}").Font.Size = 12
    $Sheet.Range("A${titleRow}:D${titleRow}").Font.Name = "Aptos Display"
    $Sheet.Range("A${titleRow}:D${titleRow}").HorizontalAlignment = -4131
    $Sheet.Range("A${titleRow}:D${titleRow}").Locked = $true
    $Sheet.Range("E${titleRow}").Value2 = "ORDER_TITLE"

    $Sheet.Range("A${infoRow}").Value2 = "Kommissionsnummer"
    $Sheet.Range("C${infoRow}").Value2 = "Strasse / Objekt"
    $Sheet.Range("A${infoRow}:D${infoRow}").Borders.LineStyle = 1
    $Sheet.Range("A${infoRow}:D${infoRow}").Borders.Weight = 2
    $Sheet.Range("A${infoRow}:D${infoRow}").Borders.Color = (XlColor 215 217 220)
    $Sheet.Range("A${infoRow}:D${infoRow}").Font.Name = "Aptos"
    $Sheet.Range("A${infoRow}:D${infoRow}").Font.Size = 10
    $Sheet.Range("A${infoRow}").Interior.Color = (XlColor 246 246 244)
    $Sheet.Range("C${infoRow}").Interior.Color = (XlColor 246 246 244)
    $Sheet.Range("B${infoRow}").Interior.Color = (XlColor 255 255 255)
    $Sheet.Range("D${infoRow}").Interior.Color = (XlColor 255 255 255)
    $Sheet.Range("B${infoRow}").Font.Size = 11
    $Sheet.Range("D${infoRow}").Font.Size = 11
    $Sheet.Range("B${infoRow}").ShrinkToFit = $true
    $Sheet.Range("D${infoRow}").ShrinkToFit = $true
    $Sheet.Range("A${infoRow}").Font.Bold = $true
    $Sheet.Range("C${infoRow}").Font.Bold = $true
    $Sheet.Range("B${infoRow}").Locked = $false
    $Sheet.Range("D${infoRow}").Locked = $false
    $Sheet.Range("A${infoRow}").Locked = $true
    $Sheet.Range("C${infoRow}").Locked = $true
    $Sheet.Range("E${infoRow}").Value2 = "ORDER_INFO"

    $Sheet.Range("A${headerRow}").Value2 = "Anzahl"
    $Sheet.Range("B${headerRow}").Value2 = "Bezeichnung der Profile"
    $Sheet.Range("C${headerRow}").Value2 = "Länge"
    $Sheet.Range("A${headerRow}:D${headerRow}").Interior.Color = (XlColor 232 234 237)
    $Sheet.Range("A${headerRow}:D${headerRow}").Font.Bold = $true
    $Sheet.Range("A${headerRow}:D${headerRow}").Font.Name = "Aptos"
    $Sheet.Range("A${headerRow}:D${headerRow}").Font.Size = 9
    $Sheet.Range("A${headerRow}:D${headerRow}").Font.Color = (XlColor 52 55 60)
    $Sheet.Range("A${headerRow}:D${headerRow}").Borders.LineStyle = 1
    $Sheet.Range("A${headerRow}:D${headerRow}").Borders.Weight = 2
    $Sheet.Range("A${headerRow}:D${headerRow}").Borders.Color = (XlColor 215 217 220)
    $Sheet.Range("A${headerRow}:D${headerRow}").Locked = $true
    $Sheet.Range("E${headerRow}").Value2 = "PROFILE_HEADER"

    Format-ProfileRow -Sheet $Sheet -Row $profileRow
    $Sheet.Rows("${titleRow}:${titleRow}").RowHeight = $rowOrderTitle
    $Sheet.Rows("${infoRow}:${infoRow}").RowHeight = $rowOrderInfo
    $Sheet.Rows("${headerRow}:${headerRow}").RowHeight = $rowProfileHeader
}

function Add-Button {
    param(
        $Sheet,
        [string]$Caption,
        [string]$Macro,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height
    )

    $button = $Sheet.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $button.TextFrame.Characters().Text = $Caption
    $button.TextFrame.HorizontalAlignment = -4108
    $button.TextFrame.VerticalAlignment = -4108
    $button.TextFrame.Characters().Font.Name = "Aptos"
    $button.TextFrame.Characters().Font.Size = 9
    $button.TextFrame.Characters().Font.Bold = $true
    $button.OnAction = $Macro
    if ($Macro -eq "ExportAndMailKindlimannOrder") {
        $button.Fill.ForeColor.RGB = (XlColor 28 28 28)
        $button.TextFrame.Characters().Font.Color = (XlColor 255 255 255)
        $button.Line.ForeColor.RGB = (XlColor 28 28 28)
    }
    elseif ($Caption.StartsWith("-")) {
        $button.Fill.ForeColor.RGB = (XlColor 255 255 255)
        $button.TextFrame.Characters().Font.Color = (XlColor 86 88 92)
        $button.Line.ForeColor.RGB = (XlColor 196 199 204)
    }
    else {
        $button.Fill.ForeColor.RGB = (XlColor 244 245 246)
        $button.TextFrame.Characters().Font.Color = (XlColor 28 28 28)
        $button.Line.ForeColor.RGB = (XlColor 196 199 204)
    }
    $button.Line.Weight = 1
    $button.DrawingObject.PrintObject = $false
    $button.Placement = 3
    $button.Locked = $true
    return $button
}

function Add-OrderProfileButton {
    param($Sheet, [int]$TitleRow)

    $buttonWidth = 24
    $buttonHeight = 20
    $cell = $Sheet.Range("D${TitleRow}")
    $left = $cell.Left + $cell.Width - $buttonWidth - 4
    $top = $cell.Top + 2
    $button = $Sheet.Shapes.AddShape(5, $left, $top, $buttonWidth, $buttonHeight)
    $button.Name = "btnProfil_" + $Sheet.Name + "_" + $TitleRow
    $button.TextFrame.Characters().Text = "+"
    $button.TextFrame.HorizontalAlignment = -4108
    $button.TextFrame.VerticalAlignment = -4108
    $button.TextFrame.Characters().Font.Name = "Aptos"
    $button.TextFrame.Characters().Font.Size = 13
    $button.TextFrame.Characters().Font.Bold = $true
    $button.TextFrame.Characters().Font.Color = (XlColor 28 28 28)
    $button.Fill.ForeColor.RGB = (XlColor 255 255 255)
    $button.Line.ForeColor.RGB = (XlColor 196 199 204)
    $button.Line.Weight = 1
    $button.OnAction = "AddProfileRowForOrderButton"
    $button.DrawingObject.PrintObject = $false
    $button.Placement = 1
    $button.Locked = $true
    return $button
}

function Setup-KWSheet {
    param($Sheet, [string]$Logo)

    $Sheet.Cells.Clear()
    $Sheet.Cells.Locked = $true
    $Sheet.Cells.Font.Name = "Aptos"
    $Sheet.Cells.Font.Size = 10
    $Sheet.Cells.Interior.Color = (XlColor 255 255 255)
    $Sheet.Columns("A").ColumnWidth = $columnWidthA
    $Sheet.Columns("B").ColumnWidth = $columnWidthB
    $Sheet.Columns("C").ColumnWidth = $columnWidthC
    $Sheet.Columns("D").ColumnWidth = $columnWidthD
    $Sheet.Columns("E").ColumnWidth = 12
    $Sheet.Columns("F").ColumnWidth = 20
    $Sheet.Columns("G").ColumnWidth = 20
    $Sheet.Columns("E").Hidden = $true
    $Sheet.Rows("1").RowHeight = $rowHeader
    $Sheet.Rows("2").RowHeight = $rowMeta
    $Sheet.Rows("3").RowHeight = $rowStatus
    $Sheet.Rows("4").RowHeight = $rowSpacer

    $Sheet.Range("A1:C1").Merge()
    $Sheet.Range("A1").Value2 = "Sammelbestellung an Kindlimann"
    $Sheet.Range("A1").Font.Bold = $true
    $Sheet.Range("A1").Font.Size = 21
    $Sheet.Range("A1").Font.Name = "Aptos Display"
    $Sheet.Range("A1").Font.Color = (XlColor 28 28 28)
    $Sheet.Range("A1").HorizontalAlignment = -4131
    $Sheet.Range("A1").VerticalAlignment = -4108
    $Sheet.Range("A1:D1").Interior.Color = (XlColor 255 255 255)
    $Sheet.Range("A1:D1").Borders(9).LineStyle = 1
    $Sheet.Range("A1:D1").Borders(9).Weight = 3
    $Sheet.Range("A1:D1").Borders(9).Color = (XlColor 28 28 28)

    $Sheet.Range("A2").Value2 = "Kalenderwoche"
    $Sheet.Range("B2").Value2 = ""
    $Sheet.Range("C2").Value2 = "Jahr"
    $Sheet.Range("D2").Value2 = ""
    $Sheet.Range("A2:D2").Interior.Color = (XlColor 246 246 244)
    $Sheet.Range("A2:D2").Font.Name = "Aptos"
    $Sheet.Range("A2:D2").Font.Size = 10
    $Sheet.Range("A2:D2").Borders.LineStyle = 1
    $Sheet.Range("A2:D2").Borders.Weight = 2
    $Sheet.Range("A2:D2").Borders.Color = (XlColor 215 217 220)
    $Sheet.Range("A2").Font.Color = (XlColor 100 104 110)
    $Sheet.Range("C2").Font.Color = (XlColor 100 104 110)
    $Sheet.Range("B2:D2").Font.Bold = $true
    $Sheet.Range("B2").HorizontalAlignment = -4108
    $Sheet.Range("D2").HorizontalAlignment = -4108

    $Sheet.Range("A3").Value2 = "Status"
    $Sheet.Range("B3:D3").Merge()
    $Sheet.Range("B3").Value2 = ""
    $Sheet.Range("A3:D3").Font.Size = 9
    $Sheet.Range("A3:D3").Font.Color = (XlColor 100 104 110)
    $Sheet.Range("A3:D3").Interior.Color = (XlColor 255 255 255)
    $Sheet.Range("A3:D3").Borders(9).LineStyle = 1
    $Sheet.Range("A3:D3").Borders(9).Weight = 2
    $Sheet.Range("A3:D3").Borders(9).Color = (XlColor 232 234 237)

    if (Test-Path -Path $Logo) {
        $logoWidth = 210
        $logoHeight = $logoWidth / (Get-ImageRatio -Path $Logo)
        $logoLeft = $Sheet.Range("D1").Left + $Sheet.Range("D1").Width - $logoWidth
        $logoTop = $Sheet.Range("D1").Top + (($Sheet.Range("D1").Height - $logoHeight) / 2)
        $logoShape = $Sheet.Shapes.AddPicture($Logo, $false, $true, $logoLeft, $logoTop, $logoWidth, $logoHeight)
        $logoShape.Name = "MetallraumLogo"
        $logoShape.LockAspectRatio = -1
        $logoShape.DrawingObject.PrintObject = $true
        $logoShape.Locked = $true
    }

    Add-OrderBlock -Sheet $Sheet -StartRow 5 -OrderNumber 1
    Add-OrderProfileButton -Sheet $Sheet -TitleRow 5 | Out-Null

    $left1 = $Sheet.Range("D5").Left + $Sheet.Range("D5").Width + $buttonGapPoints
    $left2 = $left1 + 108
    Add-Button -Sheet $Sheet -Caption "+ Auftrag" -Macro "AddOrderBlockButton" -Left $left1 -Top $Sheet.Range("F5").Top -Width 96 -Height 25 | Out-Null
    Add-Button -Sheet $Sheet -Caption "- Auftrag" -Macro "DeleteOrderBlock" -Left $left2 -Top $Sheet.Range("G5").Top -Width 96 -Height 25 | Out-Null
    Add-Button -Sheet $Sheet -Caption "- Profil" -Macro "DeleteProfileRow" -Left $left1 -Top $Sheet.Range("F7").Top -Width 96 -Height 25 | Out-Null
    Add-Button -Sheet $Sheet -Caption "Ablegen und E-Mail vorbereiten" -Macro "ExportAndMailKindlimannOrder" -Left $left1 -Top $Sheet.Range("F10").Top -Width 206 -Height 44 | Out-Null

    $Sheet.PageSetup.Orientation = 2
    $Sheet.PageSetup.PaperSize = 9
    $Sheet.PageSetup.Zoom = $false
    $Sheet.PageSetup.FitToPagesWide = 1
    $Sheet.PageSetup.FitToPagesTall = $false
    $Sheet.PageSetup.PrintTitleRows = "`$1:`$3"
    $Sheet.PageSetup.LeftMargin = $pdfMarginPoints
    $Sheet.PageSetup.RightMargin = $pdfMarginPoints
    $Sheet.PageSetup.TopMargin = $pdfMarginPoints
    $Sheet.PageSetup.BottomMargin = $pdfMarginPoints
    $Sheet.PageSetup.HeaderMargin = 14.173228
    $Sheet.PageSetup.FooterMargin = 14.173228
    $Sheet.PageSetup.PrintArea = "`$A`$1:`$D`$8"
}

function Setup-AnleitungSheet {
    param($Sheet)

    $Sheet.Cells.Clear()
    $Sheet.Cells.Locked = $true
    $Sheet.Columns("A").ColumnWidth = 32
    $Sheet.Columns("B").ColumnWidth = 90
    $Sheet.Range("A1:B1").Merge()
    $Sheet.Range("A1").Value2 = "Anleitung - Sammelbestellung Kindlimann"
    $Sheet.Range("A1").Font.Bold = $true
    $Sheet.Range("A1").Font.Size = 18
    $rows = @(
        @("1", "Auf dem passenden KW-Blatt Auftrag ausfüllen: Kommissionsnummer, Strasse/Objekt und Profile."),
        @("2", "Mit + Profil eine Profilzeile unter der aktuellen Profilzeile einfügen."),
        @("3", "Mit + Auftrag einen neuen Auftrag am Ende der Bestellung einfügen. Es bleibt eine Leerzeile Abstand."),
        @("4", "Löschen funktioniert nur über die beiden Lösch-Buttons und fragt vorher nach."),
        @("5", "Der Export prüft die Vollständigkeit. Bei Fehlern wird kein PDF erstellt."),
        @("6", "Bei erfolgreichem Export wird die PDF im Ordner Bestellt abgelegt und Outlook mit Anhang geöffnet."),
        @("7", "Die Datei muss mit Excel Desktop geöffnet werden. Excel im Browser führt die Makros nicht aus.")
    )
    $r = 3
    foreach ($item in $rows) {
        $Sheet.Range("A$r").Value2 = $item[0]
        $Sheet.Range("B$r").Value2 = $item[1]
        $Sheet.Range("A${r}:B${r}").Borders.LineStyle = 1
        $Sheet.Range("A${r}").Font.Bold = $true
        $Sheet.Range("B${r}").WrapText = $true
        $r++
    }
    $Sheet.Range("A1:B$($r + 1)").VerticalAlignment = -4160
}

function Setup-StammdatenSheet {
    param($Sheet, [string[]]$Profiles)

    $Sheet.Cells.Clear()
    $Sheet.Range("A1").Value2 = "Profile"
    for ($i = 0; $i -lt $Profiles.Count; $i++) {
        $Sheet.Range("A$($i + 2)").Value2 = $Profiles[$i]
    }
    $Sheet.Range("C1").Value2 = "Empfänger"
    $Sheet.Range("D1").Value2 = $mailTo
    $Sheet.Range("C2").Value2 = "Mailtext"
    $Sheet.Range("D2").Value2 = $mailSalutation + [Environment]::NewLine + [Environment]::NewLine + "Anbei erhalten Sie unsere Sammelbestellung für die Kalenderwoche {KW}/{JAHR}." + [Environment]::NewLine + [Environment]::NewLine + "Bei Fragen stehen wir Ihnen gerne zur Verfügung." + [Environment]::NewLine + [Environment]::NewLine + $mailClosing + [Environment]::NewLine + $mailSignature
    $Sheet.Columns("A:D").AutoFit() | Out-Null
}

$vbaCode = @'
Option Explicit

Public Const PROTECT_PASSWORD As String = "__PROTECT_PASSWORD__"
Private Const MARKER_COLUMN As Long = 5
Private Const MARKER_ORDER_TITLE As String = "ORDER_TITLE"
Private Const MARKER_ORDER_INFO As String = "ORDER_INFO"
Private Const MARKER_PROFILE_HEADER As String = "PROFILE_HEADER"
Private Const MARKER_PROFILE As String = "PROFILE"
Private Const MARKER_BLANK As String = "BLANK"
Private Const LOCK_STALE_HOURS As Double = 12
Private gLockOwner As Boolean
Private gStartupDone As Boolean

Public Sub Auto_Open()
    WorkbookStartup
End Sub

Public Sub WorkbookStartup()
    If gStartupDone Then Exit Sub
    gStartupDone = True
    On Error GoTo StartupError
    RefreshVisibleWeeks
    RefreshProfileValidations
    AcquireWorkbookLock
    Exit Sub

StartupError:
    Application.ScreenUpdating = True
    MsgBox "Die Bestellliste konnte nicht vollständig vorbereitet werden." & vbCrLf & vbCrLf & _
           "Bitte Datei schliessen und erneut öffnen." & vbCrLf & vbCrLf & _
           "Details: " & Err.Description, vbExclamation
End Sub

Public Sub RefreshVisibleWeeks()
    Dim visibleWeeks As Object
    Dim i As Long
    Dim d As Date
    Dim weekNo As Long
    Dim weekYear As Long
    Dim sheetName As String
    Dim currentWeekName As String
    Dim startDate As Date
    Dim ws As Worksheet

    Set visibleWeeks = CreateObject("Scripting.Dictionary")
    startDate = Date
    If startDate < DateSerial(__MIN_START_YEAR__, __MIN_START_MONTH__, __MIN_START_DAY__) Then startDate = DateSerial(__MIN_START_YEAR__, __MIN_START_MONTH__, __MIN_START_DAY__)

    For i = 0 To 9
        d = DateAdd("d", i * 7, startDate)
        weekNo = ISOWeekNumber(d)
        weekYear = ISOWeekYear(d)
        sheetName = "KW" & Format$(weekNo, "00")
        visibleWeeks(sheetName) = weekYear
        If i = 0 Then currentWeekName = sheetName
    Next i

    Application.ScreenUpdating = False
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PROTECT_PASSWORD
    On Error GoTo 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "KW##" Then
            If visibleWeeks.Exists(ws.Name) Then
                ws.Visible = xlSheetVisible
                UpdateWeekHeader ws, CLng(visibleWeeks(ws.Name))
            Else
                ws.Visible = xlSheetVeryHidden
            End If
        ElseIf ws.Name = "Anleitung" Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ProtectAllSheets
    ThisWorkbook.Protect Password:=PROTECT_PASSWORD, Structure:=True
    If Len(currentWeekName) > 0 Then ThisWorkbook.Worksheets(currentWeekName).Activate
    UpdateVisibleStatusLines
    Application.ScreenUpdating = True
End Sub

Private Function ISOWeekNumber(ByVal d As Date) As Long
    ISOWeekNumber = DatePart("ww", d, vbMonday, vbFirstFourDays)
End Function

Private Function ISOWeekYear(ByVal d As Date) As Long
    ISOWeekYear = Year(DateAdd("d", 4 - Weekday(d, vbMonday), d))
End Function

Private Sub UpdateWeekHeader(ByVal ws As Worksheet, ByVal yearNo As Long)
    On Error Resume Next
    ws.Unprotect Password:=PROTECT_PASSWORD
    ws.Range("B2").Value = Replace(ws.Name, "KW", "")
    ws.Range("D2").Value = yearNo
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Sub ProtectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=PROTECT_PASSWORD
        ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        On Error GoTo 0
    Next ws
End Sub

Private Sub RefreshProfileValidations()
    Dim ws As Worksheet
    Dim r As Long
    Dim maxRow As Long

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "KW##" Then
            On Error Resume Next
            ws.Unprotect Password:=PROTECT_PASSWORD
            maxRow = ws.Cells(ws.Rows.Count, MARKER_COLUMN).End(xlUp).Row
            For r = 1 To maxRow
                If MarkerAt(ws, r) = MARKER_PROFILE Then
                    ws.Range("A" & r & ":C" & r).Locked = False
                    ApplyProfileValidation ws, r
                End If
            Next r
            ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
            On Error GoTo 0
        End If
    Next ws
End Sub

Private Sub ApplyProfileValidation(ByVal ws As Worksheet, ByVal rowNo As Long)
    On Error Resume Next
    ws.Range("B" & rowNo).Validation.Delete
    On Error GoTo 0

    With ws.Range("B" & rowNo).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:="=ProfilListe"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With
End Sub

Private Sub SaveWorkbookQuietly()
    On Error Resume Next
    If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
    On Error GoTo 0
End Sub

Private Sub CloseWorkbookQuietly()
    On Error Resume Next
    ThisWorkbook.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Public Sub AcquireWorkbookLock()
    Dim ws As Worksheet
    Dim lockUser As String
    Dim lockComputer As String
    Dim lockTime As Variant
    Dim currentUser As String
    Dim currentComputer As String
    Dim ageHours As Double
    Dim lockErrorText As String

    On Error GoTo LockError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Diese Bestellliste ist schreibgeschützt geöffnet. Vermutlich arbeitet bereits eine andere Person daran. Die Datei wird geschlossen.", vbExclamation
        CloseWorkbookQuietly
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("Stammdaten")
    currentUser = CurrentUserLabel()
    currentComputer = Environ$("ComputerName")

    ws.Unprotect Password:=PROTECT_PASSWORD
    lockUser = Trim$(CStr(ws.Range("G2").Value))
    lockComputer = Trim$(CStr(ws.Range("G3").Value))
    lockTime = ws.Range("G4").Value

    If lockUser <> vbNullString And Not SameLockOwner(lockUser, lockComputer, currentUser, currentComputer) Then
        If IsDate(lockTime) Then
            ageHours = (Now - CDate(lockTime)) * 24
        Else
            ageHours = 9999
        End If

        If ageHours < LOCK_STALE_HOURS Then
            ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
            MsgBox "Diese Bestellliste ist bereits geöffnet." & vbCrLf & vbCrLf & _
                   "Aktiver Benutzer: " & lockUser & vbCrLf & _
                   "Computer: " & lockComputer & vbCrLf & _
                   "Seit: " & Format$(CDate(lockTime), "dd.mm.yyyy hh:nn") & vbCrLf & vbCrLf & _
                   "Bitte später erneut öffnen.", vbExclamation
            CloseWorkbookQuietly
            Exit Sub
        End If
    End If

    ws.Range("F1").Value = "Sperre"
    ws.Range("F2").Value = "Benutzer"
    ws.Range("F3").Value = "Computer"
    ws.Range("F4").Value = "Seit"
    ws.Range("G2").Value = currentUser
    ws.Range("G3").Value = currentComputer
    ws.Range("G4").Value = Now
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True

    gLockOwner = True
    UpdateVisibleStatusLines
    SaveWorkbookQuietly
    Exit Sub

LockError:
    lockErrorText = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    MsgBox "Die Benutzersperre konnte nicht zuverlässig geprüft werden." & vbCrLf & vbCrLf & _
           "Die Datei wird zur Sicherheit geschlossen." & vbCrLf & vbCrLf & _
           "Details: " & lockErrorText, vbExclamation
    CloseWorkbookQuietly
End Sub

Public Sub ReleaseWorkbookLock()
    Dim ws As Worksheet
    Dim currentUser As String
    Dim currentComputer As String

    If ThisWorkbook.ReadOnly Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Stammdaten")
    currentUser = CurrentUserLabel()
    currentComputer = Environ$("ComputerName")

    On Error Resume Next
    ws.Unprotect Password:=PROTECT_PASSWORD
    If gLockOwner Or SameLockOwner(CStr(ws.Range("G2").Value), CStr(ws.Range("G3").Value), currentUser, currentComputer) Then
        ws.Range("G2:G4").ClearContents
        gLockOwner = False
        SaveWorkbookQuietly
    End If
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Function CurrentUserLabel() As String
    Dim windowsUser As String
    Dim officeUser As String

    windowsUser = Trim$(Environ$("Username"))
    officeUser = Trim$(Application.UserName)

    If windowsUser <> vbNullString Then
        CurrentUserLabel = windowsUser
    ElseIf officeUser <> vbNullString Then
        CurrentUserLabel = officeUser
    Else
        CurrentUserLabel = "Unbekannt"
    End If
End Function

Private Function SameLockOwner(ByVal lockUser As String, ByVal lockComputer As String, ByVal currentUser As String, ByVal currentComputer As String) As Boolean
    SameLockOwner = (LCase$(Trim$(lockUser)) = LCase$(Trim$(currentUser)) And LCase$(Trim$(lockComputer)) = LCase$(Trim$(currentComputer)))
End Function

Private Sub UpdateVisibleStatusLines()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.Name Like "KW##" Then
            On Error Resume Next
            ws.Unprotect Password:=PROTECT_PASSWORD
            ws.Range("B3").Value = "Aktiv: " & CurrentUserLabel() & " seit " & Format$(Now, "dd.mm.yyyy hh:nn")
            ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
            On Error GoTo 0
        End If
    Next ws
End Sub

Private Sub SetExportStatus(ByVal ws As Worksheet, ByVal pdfName As String)
    ws.Range("B3").Value = "Aktiv: " & CurrentUserLabel() & " | Bestellt am " & Format$(Now, "dd.mm.yyyy hh:nn") & " | " & pdfName
End Sub

Public Sub AddProfileRow()
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim orderTitleRow As Long
    Dim headerRow As Long
    Dim insertRow As Long

    Set ws = ActiveSheet
    If Not IsKindlimannSheet(ws) Then Exit Sub

    currentRow = ActiveCell.Row
    orderTitleRow = FindOrderTitleRow(ws, currentRow)
    If orderTitleRow = 0 Then
        MsgBox "Bitte zuerst eine Zelle innerhalb eines Auftrags markieren.", vbInformation
        Exit Sub
    End If

    ws.Unprotect Password:=PROTECT_PASSWORD

    If MarkerAt(ws, currentRow) = MARKER_PROFILE Then
        insertRow = currentRow + 1
    Else
        headerRow = FindProfileHeaderRow(ws, orderTitleRow)
        insertRow = LastProfileRowInOrder(ws, orderTitleRow) + 1
        If insertRow = 1 Then insertRow = headerRow + 1
    End If

    ws.Rows(insertRow).Insert Shift:=xlDown
    FormatProfileRowVba ws, insertRow
    ws.Range("A" & insertRow).Select
    RenumberOrders ws
    ResetPrintArea ws
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

Public Sub AddProfileRowForOrderButton()
    Dim ws As Worksheet
    Dim callerName As String
    Dim callerShape As Shape
    Dim titleRow As Long
    Dim insertRow As Long
    Dim activeOrderTitle As Long

    Set ws = ActiveSheet
    If Not IsKindlimannSheet(ws) Then Exit Sub

    callerName = CStr(Application.Caller)
    On Error Resume Next
    Set callerShape = ws.Shapes(callerName)
    On Error GoTo 0

    If callerShape Is Nothing Then
        MsgBox "Der Auftrag konnte nicht eindeutig erkannt werden.", vbExclamation
        Exit Sub
    End If

    titleRow = FindOrderTitleRow(ws, callerShape.TopLeftCell.Row)
    If titleRow = 0 Then
        MsgBox "Der Auftrag konnte nicht eindeutig erkannt werden.", vbExclamation
        Exit Sub
    End If

    ws.Unprotect Password:=PROTECT_PASSWORD

    activeOrderTitle = FindOrderTitleRow(ws, ActiveCell.Row)
    If activeOrderTitle = titleRow And MarkerAt(ws, ActiveCell.Row) = MARKER_PROFILE Then
        insertRow = ActiveCell.Row + 1
    Else
        insertRow = LastProfileRowInOrder(ws, titleRow) + 1
    End If

    ws.Rows(insertRow).Insert Shift:=xlDown
    FormatProfileRowVba ws, insertRow
    ws.Range("A" & insertRow).Select
    RenumberOrders ws
    RefreshOrderProfileButtons ws
    ResetPrintArea ws
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

Public Sub AddOrderBlockButton()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim orderNo As Long

    Set ws = ActiveSheet
    If Not IsKindlimannSheet(ws) Then Exit Sub

    ws.Unprotect Password:=PROTECT_PASSWORD
    startRow = LastOrderContentRow(ws) + 1
    ws.Rows(startRow & ":" & startRow + 4).Insert Shift:=xlDown
    ws.Range("A" & startRow & ":D" & startRow).Clear
    ws.Range("E" & startRow).Value = MARKER_BLANK
    orderNo = CountOrders(ws) + 1
    FormatBlankRow ws, startRow
    FormatOrderBlockVba ws, startRow + 1, orderNo
    ws.Range("B" & startRow + 2).Select
    RenumberOrders ws
    RefreshOrderProfileButtons ws
    ResetPrintArea ws
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

Public Sub DeleteProfileRow()
    Dim ws As Worksheet
    Dim r As Long

    Set ws = ActiveSheet
    If Not IsKindlimannSheet(ws) Then Exit Sub
    r = ActiveCell.Row

    If MarkerAt(ws, r) <> MARKER_PROFILE Then
        MsgBox "Bitte zuerst eine Profilzeile markieren.", vbInformation
        Exit Sub
    End If

    If MsgBox("Diese Profilzeile wirklich löschen?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    ws.Unprotect Password:=PROTECT_PASSWORD
    ws.Rows(r).Delete
    RenumberOrders ws
    RefreshOrderProfileButtons ws
    ResetPrintArea ws
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

Public Sub DeleteOrderBlock()
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim titleRow As Long
    Dim endRow As Long
    Dim deleteFrom As Long

    Set ws = ActiveSheet
    If Not IsKindlimannSheet(ws) Then Exit Sub
    currentRow = ActiveCell.Row
    titleRow = FindOrderTitleRow(ws, currentRow)
    If titleRow = 0 Then
        MsgBox "Bitte zuerst eine Zelle innerhalb des Auftrags markieren.", vbInformation
        Exit Sub
    End If

    If MsgBox("Diesen Auftrag inklusive aller Profilzeilen wirklich löschen?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    ws.Unprotect Password:=PROTECT_PASSWORD
    endRow = LastRowInOrder(ws, titleRow)
    deleteFrom = titleRow
    If titleRow > 5 And MarkerAt(ws, titleRow - 1) = MARKER_BLANK Then deleteFrom = titleRow - 1
    ws.Rows(deleteFrom & ":" & endRow).Delete

    If CountOrders(ws) = 0 Then
        FormatOrderBlockVba ws, 5, 1
    End If

    RenumberOrders ws
    RefreshOrderProfileButtons ws
    ResetPrintArea ws
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

Public Sub ExportAndMailKindlimannOrder()
    Dim ws As Worksheet
    Dim errors As Collection
    Dim msg As String
    Dim exportFolder As String
    Dim pdfPath As String
    Dim weekNo As String
    Dim yearNo As String
    Dim baseName As String
    Dim lastRow As Long
    Dim pageErrors As Collection
    Dim errorText As String

    On Error GoTo ExportError

    Set ws = ActiveSheet
    If Not IsKindlimannSheet(ws) Then Exit Sub

    If ThisWorkbook.ReadOnly Then
        MsgBox "Diese Sammelbestellung ist schreibgeschützt geöffnet. Bitte später erneut öffnen.", vbExclamation
        Exit Sub
    End If

    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Or Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Bitte die Datei aus dem synchronisierten SharePoint-Ordner mit Excel Desktop öffnen. Der PDF-Ordner kann sonst nicht sicher bestimmt werden.", vbExclamation
        Exit Sub
    End If

    Set errors = ValidateOrderSheet(ws)
    If errors.Count > 0 Then
        msg = "PDF wurde nicht erstellt." & vbCrLf & vbCrLf & "Bitte folgende Punkte korrigieren:" & vbCrLf & CollectionToText(errors, 14)
        MsgBox msg, vbExclamation
        Exit Sub
    End If

    exportFolder = ThisWorkbook.Path & Application.PathSeparator & "Bestellt"
    If Dir(exportFolder, vbDirectory) = vbNullString Then
        MsgBox "Der Ordner 'Bestellt' wurde nicht gefunden:" & vbCrLf & exportFolder, vbExclamation
        Exit Sub
    End If

    weekNo = Format$(CLng(ws.Range("B2").Value), "00")
    yearNo = CStr(ws.Range("D2").Value)
    baseName = "Metallraum_Kindlimann_Sammelbestellung_KW" & weekNo & "_" & yearNo & ".pdf"
    pdfPath = UniquePdfPath(exportFolder, baseName)

    ws.Unprotect Password:=PROTECT_PASSWORD
    lastRow = LastOrderContentRow(ws)
    ws.PageSetup.PrintArea = "$A$1:$D$" & lastRow
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$3"
        .LeftMargin = Application.CentimetersToPoints(__PDF_MARGIN_CM__)
        .RightMargin = Application.CentimetersToPoints(__PDF_MARGIN_CM__)
        .TopMargin = Application.CentimetersToPoints(__PDF_MARGIN_CM__)
        .BottomMargin = Application.CentimetersToPoints(__PDF_MARGIN_CM__)
    End With

    Set pageErrors = ApplyOrderPageBreaks(ws, lastRow)
    If pageErrors.Count > 0 Then
        ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        MsgBox "PDF wurde nicht erstellt." & vbCrLf & vbCrLf & "Bitte folgende Punkte korrigieren:" & vbCrLf & CollectionToText(pageErrors, 10), vbExclamation
        Exit Sub
    End If

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    SetExportStatus ws, Dir$(pdfPath)
    SafeSetExportTabColor ws
    ws.Activate
    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True

    OpenOutlookMail pdfPath, weekNo, yearNo
    MsgBox "PDF wurde erstellt und die E-Mail wurde vorbereitet:" & vbCrLf & pdfPath, vbInformation
    Exit Sub

ExportError:
    errorText = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    MsgBox "PDF oder E-Mail konnte nicht vollständig erstellt werden." & vbCrLf & vbCrLf & _
           "Details: " & errorText, vbExclamation
End Sub

Private Sub SafeSetExportTabColor(ByVal ws As Worksheet)
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PROTECT_PASSWORD
    ws.Tab.Color = RGB(140, 140, 140)
    ThisWorkbook.Protect Password:=PROTECT_PASSWORD, Structure:=True
    On Error GoTo 0
End Sub

Private Function ValidateOrderSheet(ByVal ws As Worksheet) As Collection
    Dim errors As New Collection
    Dim r As Long
    Dim lastRow As Long
    Dim orderNo As Long
    Dim infoRow As Long
    Dim titleRow As Long
    Dim endRow As Long
    Dim profileCount As Long
    Dim usedProfile As Boolean
    Dim amountValue As Variant

    ws.Unprotect Password:=PROTECT_PASSWORD
    ClearValidationMarks ws

    lastRow = LastOrderContentRow(ws)
    orderNo = 0
    r = 1
    Do While r <= lastRow
        If MarkerAt(ws, r) = MARKER_ORDER_TITLE Then
            orderNo = orderNo + 1
            titleRow = r
            infoRow = r + 1
            endRow = LastRowInOrder(ws, titleRow)
            profileCount = 0

            If Trim$(CStr(ws.Range("B" & infoRow).Value)) = vbNullString Then
                errors.Add "Auftrag " & orderNo & ": Kommissionsnummer fehlt."
                MarkCell ws.Range("B" & infoRow)
            End If

            If Trim$(CStr(ws.Range("D" & infoRow).Value)) = vbNullString Then
                errors.Add "Auftrag " & orderNo & ": Strasse / Objekt fehlt."
                MarkCell ws.Range("D" & infoRow)
            End If

            Dim pr As Long
            Dim profileIndex As Long
            profileIndex = 0
            For pr = titleRow To endRow
                If MarkerAt(ws, pr) = MARKER_PROFILE Then
                    usedProfile = ProfileRowHasUserInput(ws, pr)
                    If usedProfile Then
                        profileCount = profileCount + 1
                        profileIndex = profileIndex + 1
                        amountValue = ws.Range("A" & pr).Value

                        If Trim$(CStr(amountValue)) = vbNullString Then
                            errors.Add "Auftrag " & orderNo & ", Profilzeile " & profileIndex & ": Anzahl fehlt."
                            MarkCell ws.Range("A" & pr)
                        ElseIf Not IsNumeric(amountValue) Or CDbl(amountValue) <= 0 Or CDbl(amountValue) <> CLng(amountValue) Then
                            errors.Add "Auftrag " & orderNo & ", Profilzeile " & profileIndex & ": Anzahl muss eine ganze Zahl grösser 0 sein."
                            MarkCell ws.Range("A" & pr)
                        End If

                        If Trim$(CStr(ws.Range("B" & pr).Value)) = vbNullString Then
                            errors.Add "Auftrag " & orderNo & ", Profilzeile " & profileIndex & ": Profilbezeichnung fehlt."
                            MarkCell ws.Range("B" & pr)
                        End If

                        If Trim$(CStr(ws.Range("C" & pr).Value)) = vbNullString Then
                            errors.Add "Auftrag " & orderNo & ", Profilzeile " & profileIndex & ": Länge fehlt."
                            MarkCell ws.Range("C" & pr)
                        End If
                    End If
                End If
            Next pr

            If profileCount = 0 Then
                errors.Add "Auftrag " & orderNo & ": Bitte leere Aufträge entfernen oder ausfüllen."
                MarkCell ws.Range("A" & titleRow & ":D" & endRow)
            End If

            r = endRow + 1
        Else
            r = r + 1
        End If
    Loop

    ws.Protect Password:=PROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Set ValidateOrderSheet = errors
End Function

Private Sub OpenOutlookMail(ByVal pdfPath As String, ByVal weekNo As String, ByVal yearNo As String)
    Dim olApp As Object
    Dim mail As Object
    Dim bodyText As String

    bodyText = "__MAIL_SALUTATION__" & vbCrLf & vbCrLf & _
               "Anbei erhalten Sie unsere Sammelbestellung für die Kalenderwoche " & weekNo & "/" & yearNo & "." & vbCrLf & vbCrLf & _
               "Bei Fragen stehen wir Ihnen gerne zur Verfügung." & vbCrLf & vbCrLf & _
               "__MAIL_CLOSING__" & vbCrLf & _
               "__MAIL_SIGNATURE__"

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If olApp Is Nothing Then
        MsgBox "Outlook konnte nicht geöffnet werden. Das PDF wurde trotzdem erstellt:" & vbCrLf & pdfPath, vbExclamation
        Exit Sub
    End If

    Set mail = olApp.CreateItem(0)
    With mail
        .To = "__MAIL_TO__"
        .Subject = "__MAIL_SUBJECT_PREFIX__ KW" & weekNo & "/" & yearNo
        .Body = bodyText
        .Attachments.Add pdfPath
        .Display
    End With
End Sub

Private Function IsKindlimannSheet(ByVal ws As Worksheet) As Boolean
    IsKindlimannSheet = (ws.Name Like "KW##")
    If Not IsKindlimannSheet Then
        MsgBox "Diese Funktion ist nur auf einem KW-Blatt verfügbar.", vbInformation
    End If
End Function

Private Function MarkerAt(ByVal ws As Worksheet, ByVal rowNo As Long) As String
    MarkerAt = CStr(ws.Cells(rowNo, MARKER_COLUMN).Value)
End Function

Private Function FindOrderTitleRow(ByVal ws As Worksheet, ByVal rowNo As Long) As Long
    Dim r As Long
    For r = rowNo To 1 Step -1
        If MarkerAt(ws, r) = MARKER_ORDER_TITLE Then
            FindOrderTitleRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindProfileHeaderRow(ByVal ws As Worksheet, ByVal orderTitleRow As Long) As Long
    Dim r As Long
    For r = orderTitleRow To LastRowInOrder(ws, orderTitleRow)
        If MarkerAt(ws, r) = MARKER_PROFILE_HEADER Then
            FindProfileHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function LastProfileRowInOrder(ByVal ws As Worksheet, ByVal orderTitleRow As Long) As Long
    Dim r As Long
    Dim endRow As Long
    endRow = LastRowInOrder(ws, orderTitleRow)
    For r = endRow To orderTitleRow Step -1
        If MarkerAt(ws, r) = MARKER_PROFILE Then
            LastProfileRowInOrder = r
            Exit Function
        End If
    Next r
End Function

Private Function LastRowInOrder(ByVal ws As Worksheet, ByVal orderTitleRow As Long) As Long
    Dim r As Long
    Dim maxRow As Long

    maxRow = ws.Cells(ws.Rows.Count, MARKER_COLUMN).End(xlUp).Row
    For r = orderTitleRow + 1 To maxRow
        If MarkerAt(ws, r) = MARKER_ORDER_TITLE Then
            LastRowInOrder = r - 1
            If MarkerAt(ws, LastRowInOrder) = MARKER_BLANK Then LastRowInOrder = LastRowInOrder - 1
            Exit Function
        End If
    Next r
    LastRowInOrder = maxRow
End Function

Private Function LastOrderContentRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim maxRow As Long
    maxRow = ws.Cells(ws.Rows.Count, MARKER_COLUMN).End(xlUp).Row
    For r = maxRow To 1 Step -1
        If MarkerAt(ws, r) <> vbNullString And MarkerAt(ws, r) <> MARKER_BLANK Then
            LastOrderContentRow = r
            Exit Function
        End If
    Next r
    LastOrderContentRow = 8
End Function

Private Function CountOrders(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim maxRow As Long
    maxRow = ws.Cells(ws.Rows.Count, MARKER_COLUMN).End(xlUp).Row
    For r = 1 To maxRow
        If MarkerAt(ws, r) = MARKER_ORDER_TITLE Then CountOrders = CountOrders + 1
    Next r
End Function

Private Sub RenumberOrders(ByVal ws As Worksheet)
    Dim r As Long
    Dim orderNo As Long
    Dim maxRow As Long

    maxRow = ws.Cells(ws.Rows.Count, MARKER_COLUMN).End(xlUp).Row
    orderNo = 0
    For r = 1 To maxRow
        If MarkerAt(ws, r) = MARKER_ORDER_TITLE Then
            orderNo = orderNo + 1
            ws.Range("A" & r).Value = "Auftrag " & orderNo
        End If
    Next r
End Sub

Private Sub RefreshOrderProfileButtons(ByVal ws As Worksheet)
    Dim i As Long
    Dim r As Long
    Dim maxRow As Long

    For i = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        If ws.Shapes(i).OnAction = "AddProfileRowForOrderButton" Then ws.Shapes(i).Delete
        On Error GoTo 0
    Next i

    maxRow = ws.Cells(ws.Rows.Count, MARKER_COLUMN).End(xlUp).Row
    For r = 1 To maxRow
        If MarkerAt(ws, r) = MARKER_ORDER_TITLE Then
            AddProfileButtonForOrder ws, r
        End If
    Next r
End Sub

Private Sub AddProfileButtonForOrder(ByVal ws As Worksheet, ByVal titleRow As Long)
    Dim button As Shape
    Dim target As Range
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    Dim leftPos As Double
    Dim topPos As Double

    Set target = ws.Range("D" & titleRow)
    buttonWidth = 24
    buttonHeight = 20
    leftPos = target.Left + target.Width - buttonWidth - 4
    topPos = target.Top + 2

    Set button = ws.Shapes.AddShape(5, leftPos, topPos, buttonWidth, buttonHeight)
    button.Name = "btnProfil_" & titleRow & "_" & Format$(CLng(Timer * 1000), "00000000")
    button.TextFrame.Characters.Text = "+"
    button.TextFrame.HorizontalAlignment = xlHAlignCenter
    button.TextFrame.VerticalAlignment = xlVAlignCenter
    With button.TextFrame.Characters.Font
        .Name = "Aptos"
        .Size = 13
        .Bold = True
        .Color = RGB(28, 28, 28)
    End With
    button.Fill.ForeColor.RGB = RGB(255, 255, 255)
    button.Line.ForeColor.RGB = RGB(196, 199, 204)
    button.Line.Weight = 1
    button.OnAction = "AddProfileRowForOrderButton"
    button.DrawingObject.PrintObject = False
    button.Placement = xlMoveAndSize
    button.Locked = True
End Sub

Private Sub ResetPrintArea(ByVal ws As Worksheet)
    ws.PageSetup.PrintArea = "$A$1:$D$" & LastOrderContentRow(ws)
End Sub

Private Function ApplyOrderPageBreaks(ByVal ws As Worksheet, ByVal lastRow As Long) As Collection
    Dim errors As New Collection
    Dim capacity As Double
    Dim remaining As Double
    Dim titleRow As Long
    Dim endRow As Long
    Dim orderNo As Long
    Dim blockHeight As Double
    Dim blankHeight As Double
    Dim r As Long

    ws.ResetAllPageBreaks
    capacity = A4LandscapeBodyHeight(ws)
    remaining = capacity - ws.Rows(4).RowHeight
    If remaining <= 0 Then remaining = capacity

    orderNo = 0
    For titleRow = 5 To lastRow
        If MarkerAt(ws, titleRow) = MARKER_ORDER_TITLE Then
            orderNo = orderNo + 1
            endRow = LastRowInOrder(ws, titleRow)
            blockHeight = 0
            For r = titleRow To endRow
                blockHeight = blockHeight + ws.Rows(r).RowHeight
            Next r

            If blockHeight > capacity Then
                errors.Add "Auftrag " & orderNo & ": ist zu lang für eine A4-Seite. Bitte weniger Profilzeilen verwenden oder den Auftrag teilen."
            Else
                blankHeight = 0
                If titleRow > 5 And MarkerAt(ws, titleRow - 1) = MARKER_BLANK Then
                    blankHeight = ws.Rows(titleRow - 1).RowHeight
                End If

                If blockHeight + blankHeight > remaining Then
                    ws.HPageBreaks.Add Before:=ws.Rows(titleRow)
                    remaining = capacity - blockHeight
                Else
                    remaining = remaining - blockHeight - blankHeight
                End If
            End If
        End If
    Next titleRow

    Set ApplyOrderPageBreaks = errors
End Function

Private Function A4LandscapeBodyHeight(ByVal ws As Worksheet) As Double
    Dim titleHeight As Double
    Dim r As Long

    For r = 1 To 3
        titleHeight = titleHeight + ws.Rows(r).RowHeight
    Next r

    A4LandscapeBodyHeight = 595.28 - ws.PageSetup.TopMargin - ws.PageSetup.BottomMargin - titleHeight
End Function

Private Sub FormatProfileRowVba(ByVal ws As Worksheet, ByVal rowNo As Long)
    With ws.Range("A" & rowNo & ":D" & rowNo)
        .UnMerge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(217, 220, 224)
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Locked = True
    End With
    ws.Range("A" & rowNo & ":C" & rowNo).Locked = False
    ws.Range("A" & rowNo).NumberFormat = "0"
    ws.Range("A" & rowNo).HorizontalAlignment = xlCenter
    ws.Range("B" & rowNo).HorizontalAlignment = xlLeft
    ws.Range("C" & rowNo).HorizontalAlignment = xlCenter
    ws.Range("B" & rowNo).WrapText = True
    ws.Range("C" & rowNo).ShrinkToFit = True
    ws.Range("C" & rowNo).Value = "6000 mm"
    ws.Rows(rowNo).RowHeight = __ROW_PROFILE__
    ws.Cells(rowNo, MARKER_COLUMN).Value = MARKER_PROFILE
    ApplyProfileValidation ws, rowNo
End Sub

Private Sub FormatBlankRow(ByVal ws As Worksheet, ByVal rowNo As Long)
    ws.Range("A" & rowNo & ":D" & rowNo).Clear
    ws.Range("A" & rowNo & ":D" & rowNo).Interior.Color = RGB(255, 255, 255)
    ws.Range("A" & rowNo & ":D" & rowNo).Borders.LineStyle = xlNone
    ws.Range("A" & rowNo & ":D" & rowNo).Locked = True
    ws.Rows(rowNo).RowHeight = __ROW_BLANK__
    ws.Cells(rowNo, MARKER_COLUMN).Value = MARKER_BLANK
End Sub

Private Sub FormatOrderBlockVba(ByVal ws As Worksheet, ByVal startRow As Long, ByVal orderNo As Long)
    Dim titleRow As Long
    Dim infoRow As Long
    Dim headerRow As Long
    Dim profileRow As Long

    titleRow = startRow
    infoRow = startRow + 1
    headerRow = startRow + 2
    profileRow = startRow + 3

    ws.Range("A" & titleRow & ":D" & titleRow).UnMerge
    ws.Range("A" & titleRow & ":D" & titleRow).Merge
    ws.Range("A" & titleRow).Value = "Auftrag " & orderNo
    With ws.Range("A" & titleRow & ":D" & titleRow)
        .Interior.Color = RGB(28, 28, 28)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Aptos Display"
        .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .Locked = True
    End With
    ws.Rows(titleRow).RowHeight = __ROW_ORDER_TITLE__
    ws.Cells(titleRow, MARKER_COLUMN).Value = MARKER_ORDER_TITLE

    ws.Range("A" & infoRow).Value = "Kommissionsnummer"
    ws.Range("C" & infoRow).Value = "Strasse / Objekt"
    ws.Range("B" & infoRow).ClearContents
    ws.Range("D" & infoRow).ClearContents
    With ws.Range("A" & infoRow & ":D" & infoRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(215, 217, 220)
        .Font.Name = "Aptos"
        .Font.Size = 10
    End With
    ws.Range("A" & infoRow).Interior.Color = RGB(246, 246, 244)
    ws.Range("C" & infoRow).Interior.Color = RGB(246, 246, 244)
    ws.Range("B" & infoRow).Interior.Color = RGB(255, 255, 255)
    ws.Range("D" & infoRow).Interior.Color = RGB(255, 255, 255)
    ws.Range("B" & infoRow).Font.Size = 11
    ws.Range("D" & infoRow).Font.Size = 11
    ws.Range("B" & infoRow).ShrinkToFit = True
    ws.Range("D" & infoRow).ShrinkToFit = True
    ws.Range("A" & infoRow).Font.Bold = True
    ws.Range("C" & infoRow).Font.Bold = True
    ws.Range("B" & infoRow).Locked = False
    ws.Range("D" & infoRow).Locked = False
    ws.Rows(infoRow).RowHeight = __ROW_ORDER_INFO__
    ws.Cells(infoRow, MARKER_COLUMN).Value = MARKER_ORDER_INFO

    ws.Range("A" & headerRow).Value = "Anzahl"
    ws.Range("B" & headerRow).Value = "Bezeichnung der Profile"
    ws.Range("C" & headerRow).Value = "Länge"
    ws.Range("D" & headerRow).ClearContents
    With ws.Range("A" & headerRow & ":D" & headerRow)
        .Interior.Color = RGB(232, 234, 237)
        .Font.Bold = True
        .Font.Name = "Aptos"
        .Font.Size = 9
        .Font.Color = RGB(52, 55, 60)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(215, 217, 220)
        .Locked = True
    End With
    ws.Rows(headerRow).RowHeight = __ROW_PROFILE_HEADER__
    ws.Cells(headerRow, MARKER_COLUMN).Value = MARKER_PROFILE_HEADER

    FormatProfileRowVba ws, profileRow
End Sub

Private Function ProfileRowHasUserInput(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean
    Dim a As String
    Dim b As String
    Dim c As String
    a = Trim$(CStr(ws.Range("A" & rowNo).Value))
    b = Trim$(CStr(ws.Range("B" & rowNo).Value))
    c = Trim$(CStr(ws.Range("C" & rowNo).Value))
    ProfileRowHasUserInput = (a <> vbNullString Or b <> vbNullString Or (c <> vbNullString And c <> "6000 mm"))
End Function

Private Sub ClearValidationMarks(ByVal ws As Worksheet)
    Dim maxRow As Long
    maxRow = ws.Cells(ws.Rows.Count, MARKER_COLUMN).End(xlUp).Row
    ws.Range("A5:D" & maxRow).Interior.Pattern = xlSolid
    Dim r As Long
    For r = 5 To maxRow
        Select Case MarkerAt(ws, r)
            Case MARKER_ORDER_TITLE
                ws.Range("A" & r & ":D" & r).Interior.Color = RGB(28, 28, 28)
            Case MARKER_ORDER_INFO
                ws.Range("A" & r & ":D" & r).Interior.Color = RGB(246, 246, 244)
            Case MARKER_PROFILE_HEADER
                ws.Range("A" & r & ":D" & r).Interior.Color = RGB(232, 234, 237)
            Case MARKER_PROFILE
                ws.Range("A" & r & ":D" & r).Interior.Color = RGB(255, 255, 255)
            Case MARKER_BLANK
                ws.Range("A" & r & ":D" & r).Interior.Color = RGB(255, 255, 255)
        End Select
    Next r
End Sub

Private Sub MarkCell(ByVal target As Range)
    target.Interior.Color = RGB(255, 210, 210)
End Sub

Private Function CollectionToText(ByVal items As Collection, ByVal maxItems As Long) As String
    Dim i As Long
    Dim text As String
    For i = 1 To items.Count
        If i > maxItems Then
            text = text & "- Weitere Hinweise vorhanden..." & vbCrLf
            Exit For
        End If
        text = text & "- " & CStr(items(i)) & vbCrLf
    Next i
    CollectionToText = text
End Function

Private Function UniquePdfPath(ByVal folderPath As String, ByVal fileName As String) As String
    Dim base As String
    Dim ext As String
    Dim candidate As String
    Dim i As Long

    base = Left$(fileName, Len(fileName) - 4)
    ext = ".pdf"
    candidate = folderPath & Application.PathSeparator & fileName
    If Dir(candidate) = vbNullString Then
        UniquePdfPath = candidate
        Exit Function
    End If

    i = 2
    Do
        candidate = folderPath & Application.PathSeparator & base & "_v" & i & ext
        If Dir(candidate) = vbNullString Then
            UniquePdfPath = candidate
            Exit Function
        End If
        i = i + 1
    Loop
End Function
'@

$workbookCode = @'
Private Sub Workbook_Open()
    WorkbookStartup
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ReleaseWorkbookLock
End Sub
'@

$vbaReplacements = @{
    "__PROTECT_PASSWORD__" = ConvertTo-VbaStringValue $password
    "__MAIL_TO__" = ConvertTo-VbaStringValue $mailTo
    "__MAIL_SALUTATION__" = ConvertTo-VbaStringValue $mailSalutation
    "__MAIL_SUBJECT_PREFIX__" = ConvertTo-VbaStringValue $mailSubjectPrefix
    "__MAIL_CLOSING__" = ConvertTo-VbaStringValue $mailClosing
    "__MAIL_SIGNATURE__" = ConvertTo-VbaStringValue $mailSignature
    "__PDF_MARGIN_CM__" = ConvertTo-InvariantNumber $pdfMarginCm
    "__ROW_PROFILE__" = ConvertTo-InvariantNumber $rowProfile
    "__ROW_BLANK__" = ConvertTo-InvariantNumber $rowBlank
    "__ROW_ORDER_TITLE__" = ConvertTo-InvariantNumber $rowOrderTitle
    "__ROW_ORDER_INFO__" = ConvertTo-InvariantNumber $rowOrderInfo
    "__ROW_PROFILE_HEADER__" = ConvertTo-InvariantNumber $rowProfileHeader
    "__MIN_START_YEAR__" = [string]$minimumStartDate.Year
    "__MIN_START_MONTH__" = [string]$minimumStartDate.Month
    "__MIN_START_DAY__" = [string]$minimumStartDate.Day
}

foreach ($key in $vbaReplacements.Keys) {
    $vbaCode = $vbaCode.Replace($key, [string]$vbaReplacements[$key])
}

if (!(Test-Path -Path $targetDir)) {
    throw "Target directory not found: $targetDir"
}
if (!(Test-Path -Path (Join-Path $targetDir "Bestellt"))) {
    throw "Target subdirectory not found: $(Join-Path $targetDir 'Bestellt')"
}
if (!(Test-Path -Path $logoPath)) {
    throw "Logo not found: $logoPath"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $anleitung = $workbook.Worksheets.Item(1)
    $anleitung.Name = "Anleitung"
    Setup-AnleitungSheet -Sheet $anleitung

    $missing = [System.Type]::Missing
    $stammdaten = $workbook.Worksheets.Add($missing, $anleitung, 1, $missing)
    $stammdaten.Name = "Stammdaten"
    Setup-StammdatenSheet -Sheet $stammdaten -Profiles $profiles

    $profileLastRow = 1 + $profiles.Count
    $workbook.Names.Add("ProfilListe", "='Stammdaten'!`$A`$2:`$A`$$profileLastRow") | Out-Null

    $lastSheet = $stammdaten
    for ($kw = 1; $kw -le 53; $kw++) {
        $ws = $workbook.Worksheets.Add($missing, $lastSheet, 1, $missing)
        $ws.Name = "KW{0:D2}" -f $kw
        Setup-KWSheet -Sheet $ws -Logo $logoPath
        $lastSheet = $ws
    }

    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = "modKindlimann"
    $module.CodeModule.AddFromString($vbaCode)

    $thisWorkbookModule = $null
    foreach ($component in $workbook.VBProject.VBComponents) {
        if ($component.Type -eq 100 -and $component.Name -like "*Workbook*") {
            $thisWorkbookModule = $component
            break
        }
    }
    if ($null -eq $thisWorkbookModule) {
        foreach ($component in $workbook.VBProject.VBComponents) {
            if ($component.Type -eq 100) {
                $thisWorkbookModule = $component
                break
            }
        }
    }
    $thisWorkbookModule.CodeModule.AddFromString($workbookCode)

    foreach ($ws in @($anleitung, $stammdaten)) {
        $ws.Protect($password, $true, $true, $true, $true)
    }

    foreach ($ws in $workbook.Worksheets) {
        if ($ws.Name -like "KW??") {
            $ws.Protect($password, $true, $true, $true, $true)
        }
    }

    $currentDate = Get-Date
    if ($currentDate -lt $minimumStartDate) {
        $currentDate = $minimumStartDate
    }
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $calendar = $culture.Calendar
    $visible = @{}
    $currentWeekSheetName = $null
    for ($i = 0; $i -lt 10; $i++) {
        $d = $currentDate.AddDays($i * 7)
        $week = $calendar.GetWeekOfYear($d, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
        $weekYear = $d.AddDays(3 - (([int]$d.DayOfWeek + 6) % 7)).Year
        $weekSheetName = "KW{0:D2}" -f $week
        if ($i -eq 0) { $currentWeekSheetName = $weekSheetName }
        $visible[$weekSheetName] = $weekYear
    }

    foreach ($ws in $workbook.Worksheets) {
        if ($ws.Name -like "KW??") {
            if ($visible.ContainsKey($ws.Name)) {
                $ws.Visible = -1
                $ws.Unprotect($password)
                $ws.Range("B2").Value2 = [int]$ws.Name.Substring(2, 2)
                $ws.Range("D2").Value2 = [int]$visible[$ws.Name]
                $ws.Protect($password, $true, $true, $true, $true)
            }
            else {
                $ws.Visible = 2
            }
        }
    }

    $stammdaten.Visible = 2
    $anleitung.Visible = -1
    if ($currentWeekSheetName) {
        $workbook.Worksheets.Item($currentWeekSheetName).Activate()
    }
    $workbook.Protect($password, $true)

    if (Test-Path -Path $targetFile) {
        Remove-Item -Path $targetFile -Force
    }

    $workbook.SaveAs($targetFile, 52)
    $workbook.Close($true)
}
finally {
    if ($workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Created workbook: $targetFile"
