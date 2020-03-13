#NoEnv
#MaxHotkeysPerInterval 99000000
#HotkeyInterval 99000000
#KeyHistory 0
ListLines Off
Process, Priority, , A
SetBatchLines, -1
SetKeyDelay, -1, -1
SetMouseDelay, -1
SetDefaultMouseSpeed, 0
SetWinDelay, -1  
SetControlDelay, -1
SendMode Input
SetWorkingDir %A_ScriptDir%
#Include, AutoXYWH.ahk
#NoTrayIcon
#SingleInstance off
OnExit, RunBeforeExitApp

global objExcel
global vPG_1
global vPG_2
global vED_results
global pp1Max
global pp2Max
global resultsTxt      

Gui, +Resize +MinSize250x480
gui, add, text, ym+3 w120 h20 Right, 文件夹地址(拖拽修改)
gui, add, text, y+5 wp h100 Right, 搜索内容`r`n(多行时批量搜索)
gui, add, text, y+5 wp h20 Right, 搜索范围(留空时最大)
gui, add, text, y+5 wp hp Right, 匹配模式   ;LookAt
gui, add, text, y+5 wp hp Right, 范围类型   ;LookIn
gui, add, text, y+5 wp hp Right, 区分大小写   ;MatchCase
gui, add, text, y+5 wp hp Right, 区分全角/半角   ;MatchByter
gui, add, text, y+5 wp hp Right, 搜索Excel文件数量
gui, add, text, y+5 wp hp Right, 搜索成功数量
gui, add, edit, Section x+5 ym w480 h20 ReadOnly vvED_path, % A_ScriptDir  ;文件夹地址
gui, add, edit, y+5 wp h100 vvED_whats,
gui, add, ComboBox, y+5 wp vvCB_range, % "A1:A500|A1:IV65536"     ;搜索范围
gui, add, DDL, y+5 wp AltSubmit choose2 vvDL_lookAt, 全词匹配|包含   ;LookAt
gui, add, DDL, y+5 wp AltSubmit choose1 vvDL_lookIn, 公式|值|批注   ;LookIn
gui, add, Checkbox, y+5 wp h20 vvCB_MatchCase, 区分大小写   ;MatchCase
gui, add, Checkbox, y+5 wp h20 vvCB_MatchByter, 区分全角/半角   ;MatchByter
gui, add, Progress, y+5 w430 hp Border vvPG_1, 100
gui, add, Progress, y+5 wp hp Border vvPG_2, 100
gui, add, text, x+0 yp-22 w50 hp Center vvTX_PG1,
gui, add, text, y+5 wp hp Center vvTX_PG2,
gui, add, Button, xm w500 h30 vvBT_search ggBT_search, 开始搜索
gui, add, Button, x+5 w100 h30 vvBT_copy ggBT_copy, 复制结果
gui, add, edit, xm y+5 w605 h300 ReadOnly vvED_results, 搜索结果
gui, add, StatusBar,
gui, show
SB_SetText("文件夹中包含Excel文件数量:" GetExcelCount(A_ScriptDir))
return


;搜索
gBT_search:
GuiControl, Disable, vBT_search
Gui, Submit, NoHide
;搜索字符串处理
Whats := {}
Loop, parse, vED_whats, % "`r`n"
{
	if (A_LoopField != "" and Whats[A_LoopField] != true)
		Whats[A_LoopField] := true
}
if (Whats.Count() = 0)
{
    GuiControl, Enable, vBT_search
    return
}
GuiControlGet, BT_name,, vBT_search
GuiControl,, vBT_search, 搜索中...
;搜索成功进度条初始化
GuiControl,, vPG_2, 0   ;进度条归零
pp2 := 0
pp2Max := Whats.Count()
GuiControl, % "+Range0-" pp2Max, vPG_2   ;进度条范围重设
GuiControl,, vTX_PG2, % "0/" pp2Max
;搜索文件数量进度条初始化
GuiControl,, vPG_1, 0   ;进度条归零
pp1 := 0
pp1Max := GetExcelCount(vED_path)
GuiControl, % "+Range0-" pp1Max, vPG_1   ;进度条范围重设
GuiControl,, vTX_PG1, % "0/" pp1Max
;范围类型LookIn  公式|值|批注;     -4144 批注; -4184 Threaded 批注; -4123 公式; -4163 值
Switch vDL_lookIn
{
Case 1: LookIn := -4123   ;公式
Case 2: LookIn := -4163   ;值
Case 3: LookIn := -4144   ;批注
Default:
}
;输出结果变更
resultsTxt := "结果可直接粘贴到EXCEL中`r`n"
resultsTxt .= "搜索值`t序号`t文件路径`t表格`t位置`t值`r`n"
GuiControl,, vED_results, % resultsTxt
;开始搜索
results := {}
StrArrayFindInExcels(results, Whats, vED_path, false, vCB_range, LookIn, vDL_lookAt, vCB_MatchCase, vCB_MatchByter)
;输出结果重新整理
resultsTxt := "结果可直接粘贴到EXCEL中`r`n"
resultsTxt .= "搜索值`t序号`t文件路径`t表格`t位置`t值`r`n"
for k, result in results
{
	for i, v in result
	{
		resultsTxt .= k "`t" i "`t" v.excelPath "`t" v.worksheetName "`t" v.Address "`t" v.value "`r`n"
	}
}
GuiControl,, vED_results, % resultsTxt
GuiControl,, vBT_search, % BT_name
GuiControl, Enable, vBT_search
SB_SetText("搜索完成")
return


;复制到剪贴板
gBT_copy:
GuiControlGet, resultsTxtCopy,, vED_results
Clipboard := resultsTxtCopy
SB_SetText("已将搜索结果复制到剪贴板")
return


;重设尺寸
GuiSize:
If (A_EventInfo = 1)
    Return
AutoXYWH("w", "vED_path") 
AutoXYWH("w", "vED_whats") 
AutoXYWH("w", "vCB_range") 
AutoXYWH("w", "vDL_lookAt") 
AutoXYWH("w", "vDL_lookIn") 
AutoXYWH("w", "vPG_1") 
AutoXYWH("w", "vPG_2") 
AutoXYWH("x", "vTX_PG1") 
AutoXYWH("x", "vTX_PG2") 
AutoXYWH("w", "vBT_search")
AutoXYWH("x", "vBT_copy")
AutoXYWH("wh", "vED_results")
return


;拖进文件夹
GuiDropFiles:
FileGetAttrib, Attributes, % A_GuiEvent
if InStr(Attributes, "D")
{
	GuiControl,, vED_path, % A_GuiEvent	
    SB_SetText("文件夹中包含Excel文件数量:" GetExcelCount(A_GuiEvent))
}
return


;gui关闭退出app
GuiClose:
GuiEscape:
ExitApp


;退出前执行
RunBeforeExitApp:
objExcel.Quit
ExitApp



;=============================================================================================================
;获取文件夹内excel数量
GetExcelCount(path)
{
	c := 0
	loop, Files, % path "\*", FR
	{
        if (A_LoopFileExt = "xls" or A_LoopFileExt = "xlsx")
            c++
	}
    return c
}

;在一个包含excel的文件夹中搜索字符数组(对象)
StrArrayFindInExcels(ByRef results, Whats, folderPath, onlyFirst := false, range := "", LookIn := "", LookAt := "", MatchCase := "", MatchByte := "")
{
	;正式开找
    results := IsObject(results)?results:{}
    loop, Files, % folderPath "\*", FR
    {
        if (onlyFirst and results.Count() = Whats.Count())
            break
        if (A_LoopFileExt != "xls" and A_LoopFileExt != "xlsx")
            continue
        StrArrayFindInExcel(results, Whats, A_LoopFileLongPath, onlyFirst, range, LookIn, LookAt, MatchCase, MatchByte)
		;找到数量进度条变化
		pp2 := 0
		for k, value in results
		{
			if value.Count()
				pp2++
		}
		GuiControl,, vPG_2, % pp2 ;进度条2
		GuiControl,, vTX_PG2, % pp2 "/" pp2Max
		;搜索过文件数量进度条变化
		pp1++
		GuiControl,, vPG_1, % pp1 ;进度条1
		GuiControl,, vTX_PG1, % pp1 "/" pp1Max
    }
    return results
}


;在一个excel的全部表格中搜索字符数组(对象)
StrArrayFindInExcel(ByRef results, Whats, excelPath, onlyFirst := false, range := "", LookIn := "", LookAt := "", MatchCase := "", MatchByte := "")
{
    results := IsObject(results)?results:{}
    objExcel := ComObjCreate("Excel.Application")
    try objWorkbook := objExcel.Workbooks.Open(excelPath)
    for What in Whats
    {
        results[What] := IsObject(results[What])?results[What]:[]
        if (onlyFirst and results[What].Count() >= 1)
            continue
        FindInExcel(results[What], What, objExcel, onlyFirst, range, LookIn, LookAt, MatchCase, MatchByte)
    }
    objExcel.Quit
    return results
}

;在一个excel的全部表格中搜索字符
FindInExcel(ByRef result, What, objExcel, onlyFirst := false, range := "", LookIn := "", LookAt := "", MatchCase := "", MatchByte := "")
{
    Loop % objExcel.Worksheets.count
    {
		if (onlyFirst and result.Count() >= 1)
            break
        result := FindInWorksheet(result, What, objExcel.Worksheets(A_index), onlyFirst, range, LookIn, LookAt, MatchCase, MatchByte)
    }
    return result
}

;在一个Worksheet表格中搜索字符
FindInWorksheet(ByRef result, What, worksheet, onlyFirst := false, range := "", LookIn := "", LookAt := "", MatchCase := "", MatchByte := "")
{
    LookIn := LookIn?LookIn:-4123   ;-4144 批注; -4184 Threaded 批注; -4123 公式; -4163 值
    LookAt := LookAt?LookAt:1   ;1 全词匹配; 2 包含
    MatchCase := MatchCase?MatchCase:0   ;1 区分大小写; 0 不区分大小写
    MatchByte := MatchByte?MatchByte:1   ;1 双字节字符仅匹配双字节字符; 0 双字节字符匹配其单字节等效字符
    r := range ? worksheet.Range(range) : worksheet.UsedRange
    if range
        r := worksheet.Range(range)
    else
        r := worksheet.UsedRange
    c := r.Find(What, r.Cells(1), LookIn, LookAt, 2, 1, MatchCase, MatchByte) 
    if IsObject(c)
    {
		result := IsObject(result)?result:[]
        firstAddress := c.Address
        i := result.push({Address:c.Address, excelPath:worksheet.Parent.FullName, worksheetName:worksheet.Name, value:c.Value})
        resultsTxt .= What "`t" i "`t" result[i].excelPath "`t" worksheet.Name "`t" c.Address "`t" c.Value "`r`n" ;gui变动
        GuiControl,, vED_results, % resultsTxt   ;gui变动
        if onlyFirst
            return result
        loop
        {
            c := r.FindNext(c)
            if IsObject(c) and (c.Address != firstAddress)
            {
                i := result.push({Address:c.Address, excelPath:worksheet.Parent.FullName, worksheetName:worksheet.Name, value:c.Value})
                resultsTxt .= What "`t" i "`t" result[i].excelPath "`t" worksheet.Name "`t" c.Address "`t" c.Value "`r`n" ;gui变动
                GuiControl,, vED_results, % resultsTxt   ;gui变动
            }
            else
                break
        }
    }
    return result
}