#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Configuration ---
Global DB_FILE := "stats.db" ; SQLite database file name
Global SQLITE_CLI := A_ScriptDir . "\sqlite3.exe" ; Path to the sqlite3 command-line tool

; --- Data Storage ---
Global keyCounts := Map()
Global mouseCounts := Map() ; For mouse clicks
Global appCounts := Map()
Global appUsageTime := Map()
Global appLastCountedTime := Map()
Global activeAppInfo := Map("path", "", "startTime", 0)

; --- Initial Setup ---
InitDatabase()
LoadFromDatabase() ; Load previous data on startup

; --- App Hook Setup ---
WM_SHELLHOOKMESSAGE := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK", "UInt")
OnMessage(WM_SHELLHOOKMESSAGE, ShellProc)
DllCall("RegisterShellHookWindow", "Ptr", A_ScriptHwnd)
SetTimer TrackAppUsage, 1000

; --- Database Initialization and Loading ---
InitDatabase() {
    if !FileExist(SQLITE_CLI) {
        MsgBox "错误: sqlite3.exe 未找到。`n请从 sqlite.org 下载并将其放置在脚本目录中。", "初始化失败", 48
        ExitApp
    }
    
    sql_command := "CREATE TABLE IF NOT EXISTS key_stats (date TEXT NOT NULL, key_name TEXT NOT NULL, count INTEGER NOT NULL, PRIMARY KEY (date, key_name));" 
                 . "CREATE TABLE IF NOT EXISTS app_stats (date TEXT NOT NULL, app_name TEXT NOT NULL, launch_count INTEGER NOT NULL, usage_seconds INTEGER NOT NULL, PRIMARY KEY (date, app_name));" 
                 . "CREATE TABLE IF NOT EXISTS mouse_stats (date TEXT NOT NULL, button_name TEXT NOT NULL, count INTEGER NOT NULL, PRIMARY KEY (date, button_name));"

    RunWait('"' SQLITE_CLI '" "' DB_FILE '" "' sql_command '"',, "Hide")
}

LoadFromDatabase() {
    if !FileExist(SQLITE_CLI) {
        return ; Silently fail, InitDatabase will have warned
    }

    currentDate := A_YYYY . "-" . A_MM . "-" . A_DD
    tmpFile := A_Temp . "\stat_load_tmp.csv"

    ; Load Key Stats
    sql_keys := "SELECT key_name, count FROM key_stats WHERE date = '" . currentDate . "';"
    RunWait('cmd /c ""' . SQLITE_CLI . '" -csv "' . DB_FILE . '" "' . sql_keys . '" > "' . tmpFile . '""',, "Hide")
    if (FileExist(tmpFile)) {
        Loop Read, tmpFile {
            try {
                parts := StrSplit(A_LoopReadLine, ",")
                if (parts.Length >= 2) {
                    keyCounts[parts[1]] := Integer(parts[2])
                }
            } catch {
                ; Ignore parsing errors
            }
        }
        FileDelete tmpFile
    }

    ; Load App Stats
    sql_apps := "SELECT app_name, launch_count, usage_seconds FROM app_stats WHERE date = '" . currentDate . "';"
    RunWait('cmd /c ""' . SQLITE_CLI . '" -csv "' . DB_FILE . '" "' . sql_apps . '" > "' . tmpFile . '""',, "Hide")
    if (FileExist(tmpFile)) {
        Loop Read, tmpFile {
            try {
                parts := StrSplit(A_LoopReadLine, ",")
                if (parts.Length >= 3) {
                    appCounts[parts[1]] := Integer(parts[2])
                    appUsageTime[parts[1]] := Integer(parts[3])
                }
            } catch {
                ; Ignore parsing errors
            }
        }
        FileDelete tmpFile
    }

    ; Load Mouse Stats
    sql_mouse := "SELECT button_name, count FROM mouse_stats WHERE date = '" . currentDate . "';"
    RunWait('cmd /c ""' . SQLITE_CLI . '" -csv "' . DB_FILE . '" "' . sql_mouse . '" > "' . tmpFile . '""',, "Hide")
    if (FileExist(tmpFile)) {
        Loop Read, tmpFile {
            try {
                parts := StrSplit(A_LoopReadLine, ",")
                if (parts.Length >= 2) {
                    mouseCounts[parts[1]] := Integer(parts[2])
                }
            } catch {
                ; Ignore parsing errors
            }
        }
        FileDelete tmpFile
    }
}

; --- App Usage Tracking ---
TrackAppUsage() {
    try {
        currentAppPath := WinGetProcessPath("A")
        if (currentAppPath != activeAppInfo["path"]) {
            UpdateUsage()
            activeAppInfo["path"] := currentAppPath
            if (currentAppPath != "") {
                activeAppInfo["startTime"] := A_TickCount
                SplitPath(currentAppPath, &appName)
                if !appUsageTime.Has(appName) {
                    appUsageTime[appName] := 0
                }
            } else {
                activeAppInfo["startTime"] := 0
            }
        }
    } catch {
        UpdateUsage()
        activeAppInfo["path"] := ""
        activeAppInfo["startTime"] := 0
    }
}

UpdateUsage() {
    if (activeAppInfo["path"] != "" && activeAppInfo["startTime"] > 0) {
        durationMs := A_TickCount - activeAppInfo["startTime"]
        durationSec := Floor(durationMs / 1000)
        if (durationSec > 0) {
            SplitPath(activeAppInfo["path"], &appName)
            if (appUsageTime.Has(appName)) {
                appUsageTime[appName] += durationSec
            } else {
                appUsageTime[appName] := durationSec
            }
        }
    }
}

; --- Input Hooking ---
; Keyboard
Loop 26 {
    Hotkey "~*" Chr(A_Index + 96), KeyCounter
}
Loop 10 {
    Hotkey "~*" (A_Index - 1), KeyCounter
}
Loop 12 {
    Hotkey "~*" ("F" . A_Index), KeyCounter
}
otherKeys := ["Space", "Enter", "Tab", "Backspace", "Delete", "Up", "Down", "Left", "Right", "Control", "Alt", "Shift", "Escape", "LWin", "RWin", "Home", "End", "PgUp", "PgDn", "Ins", "PrintScreen", "CapsLock", "NumLock", "ScrollLock", "NumpadDot", "NumpadAdd", "NumpadSub", "NumpadMult", "NumpadDiv"]
for key in otherKeys {
    Hotkey "~*" key, KeyCounter
}

; Mouse
mouseButtons := ["LButton", "RButton", "MButton", "XButton1", "XButton2", "WheelUp", "WheelDown"]
for btn in mouseButtons {
    Hotkey "~*" btn, MouseCounter
}

; --- Input Counting Functions ---
MouseCounter(*) {
    btnName := SubStr(A_ThisHotkey, 3)
    if mouseCounts.Has(btnName) {
        mouseCounts[btnName]++
    } else {
        mouseCounts[btnName] := 1
    }
}

KeyCounter(*) {
    local baseKey := SubStr(A_ThisHotkey, 3)
    local finalKey := ""
    local modifierKeys := ["Control", "Alt", "Shift", "LWin", "RWin"]
    
    for mod in modifierKeys {
        if (baseKey = mod) {
            finalKey := baseKey
            if keyCounts.Has(finalKey) {
                keyCounts[finalKey]++
            } else {
                keyCounts[finalKey] := 1
            }
            return
        }
    }

    local isCtrl := GetKeyState("Control", "P")
    local isAlt := GetKeyState("Alt", "P")
    local isWin := GetKeyState("LWin", "P") || GetKeyState("RWin", "P")
    local isShift := GetKeyState("Shift", "P")

    if (isCtrl || isAlt || isWin) {
        local prefix := ""
        if (isCtrl) {
            prefix .= "Ctrl+"
        }
        if (isAlt) {
            prefix .= "Alt+"
        }
        if (isShift) {
            prefix .= "Shift+"
        }
        if (isWin) {
            prefix .= "Win+"
        }
        
        local keyToCombine := RegExMatch(baseKey, "^[a-z]$") ? StrUpper(baseKey) : baseKey
        finalKey := prefix . keyToCombine
    } else if (isShift) {
        if (RegExMatch(baseKey, "^[a-z]$")) {
            finalKey := StrUpper(baseKey)
        } else if (RegExMatch(baseKey, "^\d$")) {
            return
        } else {
            finalKey := "Shift+" . baseKey
        }
    } else {
        finalKey := baseKey
    }

    if (finalKey != "") {
        if keyCounts.Has(finalKey) {
            keyCounts[finalKey]++
        } else {
            keyCounts[finalKey] := 1
        }
    }
}

; --- Shell Hook for App Tracking ---
ShellProc(wParam, lParam, *) {
    HSHELL_WINDOWCREATED := 1
    if (wParam != HSHELL_WINDOWCREATED) {
        return
    }
    try {
        hwnd := lParam
        processPath := WinGetProcessPath("ahk_id " . hwnd)
        if (processPath = "") {
            return
        }
        SplitPath(processPath, &appName)
        if (appName = "") {
            return
        }
        if (appLastCountedTime.Has(appName) && (A_TickCount - appLastCountedTime[appName] < 2000)) {
            return
        }
        appLastCountedTime[appName] := A_TickCount
        if (appCounts.Has(appName)) {
            appCounts[appName]++
        } else {
            appCounts[appName] := 1
        }
        if (!appUsageTime.Has(appName)) {
            appUsageTime[appName] := 0
        }
    } catch {
        ; Ignore errors
    }
}

; --- Hotkey to Show Stats and Save to DB ---
^+!t::
{
    UpdateUsage()

    if (keyCounts.Count = 0 && appCounts.Count = 0 && mouseCounts.Count = 0) {
        MsgBox "尚未记录任何按键或应用启动。", "统计信息"
        return
    }

    saveNotification := SaveToDatabase()

    keyTable := BuildTable(keyCounts, "按键", "次数")
    mouseTable := BuildTable(mouseCounts, "鼠标按键", "次数")
    appTable := BuildAppTable(appCounts, appUsageTime)
    
    displayOutput := ""
    if (keyTable != "") {
        displayOutput .= "--- 按键统计 ---`n" . keyTable
    }
    if (mouseTable != "") {
        if (displayOutput != "") {
            displayOutput .= "`n`n"
        }
        displayOutput .= "--- 鼠标统计 ---`n" . mouseTable
    }
    if (appTable != "") {
        if (displayOutput != "") {
            displayOutput .= "`n`n"
        }
        displayOutput .= "--- 应用统计 ---`n" . appTable
    }

    finalText := saveNotification . "`n`n" . displayOutput
    ShowStatsGui(finalText, keyCounts, mouseCounts, appCounts, appUsageTime)
}

; --- GUI Display Function ---
ShowStatsGui(text, keyCounts, mouseCounts, appCounts, appUsageTime) {
    StatsGui := Gui("+Resize", "按键与应用统计")
    StatsGui.SetFont("s10", "Consolas") ; Use a monospaced font for alignment
    StatsGui.Add("Edit", "w780 h550 ReadOnly +VScroll", text)

    ExportJsonButton := StatsGui.Add("Button", "x195 w120 h30", "导出为 JSON")
    ExportJsonButton.OnEvent("Click", (*) => ExportToJson(keyCounts, mouseCounts, appCounts, appUsageTime))

    OkButton := StatsGui.Add("Button", "x340 w120 h30 Default", "确定")
    OkButton.OnEvent("Click", (*) => StatsGui.Destroy())

    ExportHtmlButton := StatsGui.Add("Button", "x485 w120 h30", "导出为 HTML")
    ExportHtmlButton.OnEvent("Click", (*) => ExportToHtml(keyCounts, mouseCounts, appCounts, appUsageTime))

    StatsGui.OnEvent("Close", (*) => StatsGui.Destroy())
    StatsGui.Show()
}

; --- Export Functions ---
; --- Export Functions ---
ExportToJson(keyCounts, mouseCounts, appCounts, appUsageTime) {
    currentDate := A_YYYY . "-" . A_MM . "-" . A_DD
    
    data := Map()
    data["date"] := currentDate
    
    stats := Map()
    
    keyData := Map()
    for key, count in keyCounts {
        keyData[key] := count
    }
    stats["keys"] := keyData
    
    mouseData := Map()
    for btn, count in mouseCounts {
        mouseData[btn] := count
    }
    stats["mouse"] := mouseData
    
    appData := []
    for appName, launchCount in appCounts {
        usageSec := appUsageTime.Has(appName) ? appUsageTime[appName] : 0
        appEntry := Map("name", appName, "launchCount", launchCount, "usageSeconds", usageSec)
        appData.Push(appEntry)
    }
    stats["applications"] := appData
    
    data["statistics"] := stats
    
    jsonString := MapToJson(data)
    
filePath := FileSelect("S", A_ScriptDir . "\stats_" . currentDate . ".json", "Save As", "JSON Files (*.json)")
    if (filePath = "") {
        return
    }
    
    try {
        file := FileOpen(filePath, "w", "UTF-8")
        file.Write(jsonString)
        file.Close()
        MsgBox "数据已成功导出到 " . filePath, "导出成功", 64
    } catch {
        MsgBox "导出文件时出错。", "错误", 16
    }
}

MapToJson(map) {
    json := "{"
    first := true
    for k, v in map {
        if (!first) {
            json .= ","
        }
        json .= Chr(34) . EscapeJsonString(k) . Chr(34) . ":" . ValueToJson(v)
        first := false
    }
    json .= "}"
    return json
}

ArrayToJson(arr) {
    json := "["
    first := true
    for _, v in arr {
        if (!first) {
            json .= ","
        }
        json .= ValueToJson(v)
        first := false
    }
    json .= "]"
    return json
}

ValueToJson(val) {
    if (IsObject(val)) {
        if (val.Has(1) || val.Length = 0) { ; Check if it's an array-like map or empty array
            return ArrayToJson(val)
        } else {
            return MapToJson(val)
        }
    } else if (IsInteger(val) || IsFloat(val)) {
        return val
    } else {
        return Chr(34) . EscapeJsonString(String(val)) . Chr(34)
    }
}

EscapeJsonString(str) {
    str := StrReplace(str, "\", "\\")
	str := StrReplace(str, Chr(34), "\" . Chr(34))

    str := StrReplace(str, "/", "\/")
    str := StrReplace(str, "`b", "\b")
    str := StrReplace(str, "`f", "\f")
    str := StrReplace(str, "`n", "\n")
    str := StrReplace(str, "`r", "\r")
    str := StrReplace(str, "`t", "\t")
    return str
}

ExportToHtml(keyCounts, mouseCounts, appCounts, appUsageTime) {
    currentDate := A_YYYY . "-" . A_MM . "-" . A_DD
    
    html := ""
    html .= BuildHtmlHeader(currentDate)
    html .= BuildHtmlBody(keyCounts, mouseCounts, appCounts, appUsageTime)
    html .= BuildHtmlFooter()

filePath := FileSelect("S", A_ScriptDir . "\stats_report_" . currentDate . ".html", "Save As", "HTML Files (*.html)")
    if (filePath = "") {
        return
    }
    
    try {
        file := FileOpen(filePath, "w", "UTF-8")
        file.Write(html)
        file.Close()
        MsgBox "报告已成功导出到 " . filePath, "导出成功", 64
    } catch {
        MsgBox "导出文件时出错。", "错误", 16
    }
}



BuildHtmlHeader(date) {
    return (
"rwa"
)
}

BuildHtmlFooter() {
    return (
    "hello"
    )
}

BuildHtmlBody(keyCounts, mouseCounts, appCounts, appUsageTime) {
    body := ""

    ; Key Stats
    if (keyCounts.Count > 0) {
        body .= "<h2>按键统计</h2><table><tr><th>按键</th><th>次数</th></tr>"
        for key, count in keyCounts {
            body .= "<tr><td>" . EscapeHtml(key) . "</td><td>" . count . "</td></tr>"
        }
        body .= "</table>"
    }

    ; Mouse Stats
    if (mouseCounts.Count > 0) {
        body .= "<h2>鼠标统计</h2><table><tr><th>按钮</th><th>次数</th></tr>"
        for btn, count in mouseCounts {
            body .= "<tr><td>" . EscapeHtml(btn) . "</td><td>" . count . "</td></tr>"
        }
        body .= "</table>"
    }

    ; App Stats
    if (appCounts.Count > 0) {
        body .= "<h2>应用程序统计</h2><table><tr><th>程序</th><th>启动次数</th><th>使用时长</th></tr>"
        for appName, launchCount in appCounts {
            usageSec := appUsageTime.Has(appName) ? appUsageTime[appName] : 0
            usageFormatted := FormatUsageTime(usageSec)
            body .= "<tr><td>" . EscapeHtml(appName) . "</td><td>" . launchCount . "</td><td>" . usageFormatted . "</td></tr>"
        }
        body .= "</table>"
    }
    
    return body
}

EscapeHtml(str) {
    str := StrReplace(str, "&", "&amp;")
    str := StrReplace(str, "<", "&lt;")
    str := StrReplace(str, ">", "&gt;")
    str := StrReplace(str, Chr(34), "&quot;")
    str := StrReplace(str, "'", "&#39;")
    return str
}

; --- Database Saving Function ---
SaveToDatabase() {
    if !FileExist(SQLITE_CLI) {
        return "错误: sqlite3.exe 未找到。"
    }
    
    currentDate := A_YYYY . "-" . A_MM . "-" . A_DD
    
    RunWait('"" . SQLITE_CLI . """ """ . DB_FILE . """ "BEGIN;"',, "Hide")

    for key, count in keyCounts {
        escKey := StrReplace(key, "'", "''")
; sql := "INSERT INTO key_stats (date, key_name, count) VALUES (""" . currentDate . """, """ . escKey . """, " . count . ") ON CONFLICT(date, key_name) DO UPDATE SET count = " . count . ";"
		sql_template := "INSERT INTO key_stats (date, key_name, count) VALUES ('{}', '{}', {}) ON CONFLICT(date, key_name) DO UPDATE SET count = {};"
sql := Format(sql_template, currentDate, escKey, count, count)
        RunWait('"" . SQLITE_CLI . """ """ . DB_FILE . """ """ . sql . """',, "Hide")
    }

    for btn, count in mouseCounts {
        escBtn := StrReplace(btn, "'", "''")
;sql := "INSERT INTO mouse_stats (date, button_name, count) VALUES (""" . currentDate . """, """ . escBtn . """, " . count . ") ON CONFLICT(date, button_name) DO UPDATE SET count = " . count . ";"
        sql_template := "INSERT INTO mouse_stats (date, button_name, count) VALUES ('{}', '{}', {}) ON CONFLICT(date, button_name) DO UPDATE SET count = {};"
		sql := Format(sql_template, currentDate, escBtn, count, count)
		RunWait('"" . SQLITE_CLI . """ """ . DB_FILE . """ """ . sql . """',, "Hide")
    }

    for appName, launchCount in appCounts {
        usage := appUsageTime.Has(appName) ? appUsageTime[appName] : 0
        escAppName := StrReplace(appName, "'", "''")
;sql := "INSERT INTO app_stats (date, app_name, launch_count, usage_seconds) VALUES (""" . currentDate . """, """ . escAppName . """, " . launchCount . ", " . usage . ") ON CONFLICT(date, app_name) DO UPDATE SET launch_count = " . launchCount . ", usage_seconds = " . usage . ";"
        sql_template := "INSERT INTO app_stats (date, app_name, launch_count, usage_seconds) VALUES ('{}', '{}', {}, {}) ON CONFLICT(date, app_name) DO UPDATE SET launch_count = {}, usage_seconds = {};"
		sql := Format(sql_template, currentDate, escAppName, launchCount, usage, launchCount, usage)
		RunWait('"" . SQLITE_CLI . """ """ . DB_FILE . """ """ . sql . """',, "Hide")
    }
    
    RunWait('"" . SQLITE_CLI . """ """ . DB_FILE . """ "COMMIT;"',, "Hide")

    return "统计信息已保存到数据库 " . DB_FILE
}


; --- Helper functions for display ---
BuildTable(map, header1, header2) {
    if (map.Count = 0) {
        return ""
    }
    maxLen := StrLen(header1)
    for key in map {
        if (StrLen(key) > maxLen) {
            maxLen := StrLen(key)
        }
    }

    col1Width := maxLen + 2
    header := Pad(header1, col1Width) . "| " . header2 . "`n"
    separator := ""
    Loop (col1Width + StrLen(header2) + 2) {
        separator .= "-"
    }
    separator .= "`n"

    content := ""
    for key, count in map {
        content .= Pad(key, col1Width) . "| " . count . "`n"
    }

    return header . separator . content
}

BuildAppTable(appCounts, appUsageTime) {
    if (appCounts.Count = 0) {
        return ""
    }
    maxLen := StrLen("程序")
    for appName in appCounts {
        if (StrLen(appName) > maxLen) {
            maxLen := StrLen(appName)
        }
    }

    col1Width := maxLen + 2
    header1 := "程序", header2 := "启动次数", header3 := "使用时长"
    
    output := Pad(header1, col1Width) . "| " . Pad(header2, 12) . "| " . header3 . "`n"
    
    separator := ""
    headerLine := Pad(header1, col1Width) . "| " . Pad(header2, 12) . "| " . header3
    Loop StrLen(headerLine) {
        separator .= "-"
    }
    output .= separator . "`n"

    for appName, launchCount in appCounts {
        usageSec := appUsageTime.Has(appName) ? appUsageTime[appName] : 0
        usageFormatted := FormatUsageTime(usageSec)
        output .= Pad(appName, col1Width) . "| " . Pad(launchCount, 12) . "| " . usageFormatted . "`n"
    }
    return output
}

Pad(str, len) {
    return str . SubStr("                                                                      ", 1, len - StrLen(str))
}

FormatUsageTime(seconds) {
    if (seconds < 60) {
        return seconds . " 秒"
    } else if (seconds < 3600) {
        return Round(seconds / 60, 1) . " 分钟"
    } else {
        return Round(seconds / 3600, 2) . " 小时"
    }
}
