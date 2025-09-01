#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Configuration ---
Global DB_FILE := "stats.db" ; SQLite database file name
Global SQLITE_CLI := A_ScriptDir . "\sqlite3.exe" ; Path to the sqlite3 command-line tool

; --- Data Storage ---
Global keyCounts := Map()
Global appCounts := Map()
Global appUsageTime := Map()
Global appLastCountedTime := Map()
Global activeAppInfo := Map("path", "", "startTime", 0)

; --- Initial Setup ---
InitDatabase()

; --- App Hook Setup ---
WM_SHELLHOOKMESSAGE := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK", "UInt")
OnMessage(WM_SHELLHOOKMESSAGE, ShellProc)
DllCall("RegisterShellHookWindow", "Ptr", A_ScriptHwnd)
SetTimer TrackAppUsage, 1000

; --- Database Initialization ---
InitDatabase() {
    if !FileExist(SQLITE_CLI) {
        MsgBox "错误: sqlite3.exe 未找到。`n请从 sqlite.org 下载并将其放置在脚本目录中。", "初始化失败", 48
        ExitApp
    }
    
    sql_command := "CREATE TABLE IF NOT EXISTS key_stats (date TEXT NOT NULL, key_name TEXT NOT NULL, count INTEGER NOT NULL, PRIMARY KEY (date, key_name));" 
                 . "CREATE TABLE IF NOT EXISTS app_stats (date TEXT NOT NULL, app_name TEXT NOT NULL, launch_count INTEGER NOT NULL, usage_seconds INTEGER NOT NULL, PRIMARY KEY (date, app_name));"

    RunWait('"' SQLITE_CLI '" "' DB_FILE '" "' sql_command '"',, "Hide")
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

; --- Key Hooking ---
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

; --- Key Counting Function ---
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

    if (keyCounts.Count = 0 && appCounts.Count = 0) {
        MsgBox "尚未记录任何按键或应用启动。", "统计信息"
        return
    }

    saveNotification := SaveToDatabase()

    keyTable := BuildTable(keyCounts, "按键", "次数")
    appTable := BuildAppTable(appCounts, appUsageTime)
    displayOutput := ""
    if (keyTable != "") {
        displayOutput .= "--- 按键统计 ---`n" . keyTable
    }
    if (appTable != "") {
        if (displayOutput != "") {
            displayOutput .= "`n`n"
        }
        displayOutput .= "--- 应用统计 ---`n" . appTable
    }

    finalMsgBoxText := saveNotification . "`n`n" . displayOutput
    MsgBox finalMsgBoxText, "按键与应用统计"
}

; --- Database Saving Function ---
SaveToDatabase() {
    if !FileExist(SQLITE_CLI) {
        return "错误: sqlite3.exe 未找到。"
    }
    
    currentDate := A_YYYY . "-" . A_MM . "-" . A_DD
    
    RunWait('"' SQLITE_CLI '" "' DB_FILE '" "BEGIN;"',, "Hide")

    for key, count in keyCounts {
        escKey := StrReplace(key, "'", "''")
        sql := "INSERT INTO key_stats (date, key_name, count) VALUES ('" . currentDate . "', '" . escKey . "', " . count . ") ON CONFLICT(date, key_name) DO UPDATE SET count = " . count . ";"
        RunWait('"' SQLITE_CLI '" "' DB_FILE '" "' sql '"',, "Hide")
    }

    for appName, launchCount in appCounts {
        usage := appUsageTime.Has(appName) ? appUsageTime[appName] : 0
        escAppName := StrReplace(appName, "'", "''")
        sql := "INSERT INTO app_stats (date, app_name, launch_count, usage_seconds) VALUES ('" . currentDate . "', '" . escAppName . "', " . launchCount . ", " . usage . ") ON CONFLICT(date, app_name) DO UPDATE SET launch_count = " . launchCount . ", usage_seconds = " . usage . ";"
        RunWait('"' SQLITE_CLI '" "' DB_FILE '" "' sql '"',, "Hide")
    }
    
    RunWait('"' SQLITE_CLI '" "' DB_FILE '" "COMMIT;"',, "Hide")

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