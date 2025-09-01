
; AutoHotkey v2 script
#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Data Storage ---
; 使用 Map 对象来存储每个按键的计数值
Global keyCounts := Map()
; 使用 Map 对象来存储每个应用启动的计数值
Global appCounts := Map()
; 使用 Map 对象来存储每个应用的使用时长 (秒)
Global appUsageTime := Map()
Global appLastCountedTime := Map() ; For debouncing app counts
Global activeAppInfo := Map("path", "", "startTime", 0) ; For tracking the current active app

; --- App Hook Setup ---
WM_SHELLHOOKMESSAGE := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK", "UInt")
OnMessage(WM_SHELLHOOKMESSAGE, ShellProc)
DllCall("RegisterShellHookWindow", "Ptr", A_ScriptHwnd)
SetTimer TrackAppUsage, 1000 ; Check active app every second for usage tracking

; --- App Usage Tracking ---
TrackAppUsage() {
    try {
        currentAppPath := WinGetProcessPath("A")

        ; If the active app has changed from what we were tracking
        if (currentAppPath != activeAppInfo["path"]) {
            ; Finalize and save the usage time for the previous app
            UpdateUsage()

            ; Reset and start tracking the new app
            activeAppInfo["path"] := currentAppPath
            if (currentAppPath != "") {
                activeAppInfo["startTime"] := A_TickCount
                SplitPath(currentAppPath, &appName)
                ; Ensure the app has an entry in the usage map
                if !appUsageTime.Has(appName) {
                    appUsageTime[appName] := 0
                }
            } else {
                ; No active app (e.g., desktop), so reset start time
                activeAppInfo["startTime"] := 0
            }
        }
    } catch {
        ; On error, just stop tracking the current app
        UpdateUsage()
        activeAppInfo["path"] := ""
        activeAppInfo["startTime"] := 0
    }
}

UpdateUsage() {
    ; Check if there was a valid app being tracked
    if (activeAppInfo["path"] != "" && activeAppInfo["startTime"] > 0) {
        durationMs := A_TickCount - activeAppInfo["startTime"]
        durationSec := Floor(durationMs / 1000) ; Use Floor to get whole seconds

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


; --- 要追踪的按键列表 ---
; 字母
Loop 26 {
    key := Chr(A_Index + 96)
    Hotkey "~*" key, KeyCounter
}
; 数字
Loop 10 {
    key := A_Index - 1
    Hotkey "~*" key, KeyCounter
}
; F1-F12 功能键
Loop 12 {
    key := "F" . A_Index
    Hotkey "~*" key, KeyCounter
}
; 其他常用按键
otherKeys := [
    "Space", "Enter", "Tab", "Backspace", "Delete", "Up", "Down", "Left", "Right", 
    "Control", "Alt", "Shift", "Escape", "LWin", "RWin",
    "Home", "End", "PgUp", "PgDn", "Ins", "PrintScreen",
    "CapsLock", "NumLock", "ScrollLock",
    "NumpadDot", "NumpadAdd", "NumpadSub", "NumpadMult", "NumpadDiv"
]
for key in otherKeys {
    Hotkey "~*" key, KeyCounter
}


; --- 按键计数函数 ---
; 每当上面定义的任意一个热键被触发时，此函数就会被调用
KeyCounter(*) {
    ; 从 A_ThisHotkey (例如 "~*c") 中提取基础按键名 ("c")
    local baseKey := SubStr(A_ThisHotkey, 3)
    local finalKey := ""

    ; 首先，判断触发的键本身是否是修饰键。如果是，则直接统计并返回。
    local modifierKeys := ["Control", "Alt", "Shift", "LWin", "RWin"]
    for mod in modifierKeys {
        if (baseKey = mod) {
            finalKey := baseKey
            ; 更新统计并立即返回
            if keyCounts.Has(finalKey) {
                keyCounts[finalKey]++
            } else {
                keyCounts[finalKey] := 1
            }
            return
        }
    }

    ; 如果代码执行到这里，说明按下的不是修饰键本身，需要检查组合情况。
    local isCtrl := GetKeyState("Control", "P")
    local isAlt := GetKeyState("Alt", "P")
    local isWin := GetKeyState("LWin", "P") || GetKeyState("RWin", "P")
    local isShift := GetKeyState("Shift", "P")

    ; 情况一: 存在 Ctrl, Alt, 或 Win 的“真正”组合键 (例如 Ctrl+Shift+T)
    if (isCtrl || isAlt || isWin) {
        local prefix := ""
        if isCtrl
            prefix .= "Ctrl+"
        if isAlt
            prefix .= "Alt+"
        if isShift ; Shift键只有在和其他修饰键组合时才在这里处理
            prefix .= "Shift+"
        if isWin
            prefix .= "Win+"
        
        local keyToCombine := baseKey
        if (RegExMatch(keyToCombine, "^[a-z]$")) {
            keyToCombine := StrUpper(keyToCombine)
        }
        finalKey := prefix . keyToCombine
    }
    ; 情况二: 只存在 Shift 的“输入型”组合
    else if (isShift) {
        ; Shift + 字母: 我们将其视为输入大写字母
        if (RegExMatch(baseKey, "^[a-z]$")) {
            finalKey := StrUpper(baseKey)
        }
        ; Shift + 数字: 我们视其为输入符号，并根据要求过滤掉，不统计
        else if (RegExMatch(baseKey, "^\d$")) {
            return ; 直接退出，不进行任何统计
        }
        ; 其他 Shift 组合 (例如 Shift+Tab, Shift+Enter): 仍然视为有效组合
        else {
            finalKey := "Shift+" . baseKey
        }
    }
    ; 情况三: 没有任何修饰键，只是一个普通的按键
    else {
        finalKey := baseKey
    }

    ; 如果 finalKey 有效，则更新统计
    if (finalKey != "") {
        if keyCounts.Has(finalKey) {
            keyCounts[finalKey]++
        } else {
            keyCounts[finalKey] := 1
        }
    }
}

; --- Shell Hook Callback for App Tracking ---
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

        ; 过滤掉空的或无效的程序名
        if (appName = "") {
            return
        }

        ; 去重: 2秒内同一个程序只计数一次
        if (appLastCountedTime.Has(appName) && (A_TickCount - appLastCountedTime[appName] < 2000)) {
            return
        }
        appLastCountedTime[appName] := A_TickCount

        ; 增加计数
        if (appCounts.Has(appName)) {
            appCounts[appName]++
        } else {
            appCounts[appName] := 1
        }
        ; Initialize usage time if not present
        if (!appUsageTime.Has(appName)) {
            appUsageTime[appName] := 0
        }
    } catch {
        ; 忽略在获取窗口或进程信息时可能发生的错误
    }
}


; --- 显示统计信息并保存到文件的全局热键 ---
; ^: Ctrl, !: Alt, +: Shift
^+!t::
{
    ; First, update usage for the last active app before showing stats
    UpdateUsage()

    ; 如果没有任何记录，则只提示，不生成文件
    if (keyCounts.Count = 0 && appCounts.Count = 0) {
        MsgBox "尚未记录任何按键或应用启动。", "统计信息"
        return
    }

    ; --- 1. 构建JSON字符串并写入文件 ---
    keyJson := BuildJson(keyCounts)
    
    ; Build app JSON with launches and usage
    local appJson := "{"
    local isFirstApp := true
    for appName, launchCount in appCounts {
        if !isFirstApp
            appJson .= ","
        
        usage := appUsageTime.Has(appName) ? appUsageTime[appName] : 0
        
        local escapedAppName := StrReplace(appName, "\", "\\")
        escapedAppName := StrReplace(escapedAppName, '"', '"')

        appJson .= '"' . escapedAppName . '":{"launches":' . launchCount . ',"usage_seconds":' . usage . '}'
        isFirstApp := false
    }
    appJson .= "}"

    jsonString := '{"keys":' . (keyJson ? keyJson : '{}') . ',"apps":' . (appJson ? appJson : '{}') . '}'

    fileName := A_YYYY . A_MM . A_DD . ".json"
    saveNotification := ""
    try {
        file := FileOpen(fileName, "w", "UTF-8")
        file.Write(jsonString)
        file.Close()
        saveNotification := "统计信息已保存到 " . fileName
    } catch {
        saveNotification := "错误：无法写入文件 " . fileName
    }

    ; --- 2. 美化输出为表格 ---
    keyTable := BuildTable(keyCounts, "按键", "次数")
    appTable := BuildAppTable(appCounts, appUsageTime)

    displayOutput := ""
    if (keyTable != "") {
        displayOutput .= "--- 按键统计 ---" . "`n" . keyTable
    }
    if (appTable != "") {
        if (displayOutput != "") {
            displayOutput .= "`n`n" ; Add space between tables
        }
        displayOutput .= "--- 应用统计 ---" . "`n" . appTable
    }

    ; --- 3. 弹出最终结果 ---
    finalMsgBoxText := saveNotification . "`n`n" . displayOutput
    MsgBox finalMsgBoxText, "按键与应用统计"
}

; --- 用于生成统计信息的辅助函数 ---
BuildJson(map) {
    if (map.Count = 0) {
        return ""
    }
    local json := "{"
    local isFirst := true
    for key, count in map {
        if !isFirst {
            json .= ","
        }
        local escapedKey := StrReplace(key, "\", "\\")
        escapedKey := StrReplace(escapedKey, '"', '"')
        json .= '"' . escapedKey . '":' . count
        isFirst := false
    }
    json .= "}"
    return json
}

BuildTable(map, header1, header2) {
    if (map.Count = 0) {
        return ""
    }
    ; a. 找到最长的键名长度以确定列宽
    local maxLen := StrLen(header1)
    for key, count in map {
        len := StrLen(key)
        if (len > maxLen) {
            maxLen := len
        }
    }

    ; b. 构建表格头部
    local col1Width := maxLen + 2
    local spaces := "                                        " ; 用于填充的空格
    local header := header1 . SubStr(spaces, 1, col1Width - StrLen(header1)) . "| " . header2 . "`n"
    local separator := ""
    Loop (col1Width + StrLen(header2) + 2) { ; 分隔符长度
        separator .= "-"
    }
    separator .= "`n"

    ; c. 构建表格主体内容
    local content := ""
    for key, count in map {
        paddedKey := key . spaces
        paddedKey := SubStr(paddedKey, 1, col1Width)
        content .= paddedKey . "| " . count . "`n"
    }

    return header . separator . content
}

BuildAppTable(appCounts, appUsageTime) {
    if (appCounts.Count = 0) {
        return ""
    }
    local maxLen := StrLen("程序")
    for appName in appCounts {
        if (StrLen(appName) > maxLen)
            maxLen := StrLen(appName)
    }

    local col1Width := maxLen + 2
    local header1 := "程序", header2 := "启动次数", header3 := "使用时长"
    
    local output := Pad(header1, col1Width) . "| " . Pad(header2, 12) . "| " . header3 . "`n"
    
    local separator := ""
    local headerLine := Pad(header1, col1Width) . "| " . Pad(header2, 12) . "| " . header3
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
