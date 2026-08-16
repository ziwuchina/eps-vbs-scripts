' ============================================================
' 导出全部属性.vbs（全属性版 v4）
' 功能：导出所选要素的全部属性（基本属性 + 全量扩展属性）
' 支持：选择集 / 全部范围，POINT / LINE / AREA / NOTE / 全部
' 扩展属性：硬编码全量清单 + MemoData 动态枚举合并
' 取值：优先 MemoData（不带括号），无则 GetSelGeoValue（带括号）
' 说明：仅用最基础 VBScript 语法，不依赖任何 COM 对象
' ============================================================

Sub OnClick()
    MsgBox "脚本已启动，开始执行"
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "要素类型", "全部", 0, "POINT,LINE,AREA,NOTE,全部", ""
    SSProcess.AddInputParameter "导出范围", "选择集", 0, "选择集,全部", ""
    SSProcess.AddInputParameter "输出文件", "C:\全部属性导出.csv", 0, "", ""
    res = SSProcess.ShowInputParameterDlg("导出全部属性")
    If res = 0 Then
        Exit Sub
    End If
    SSProcess.UpdateScriptDlgParameter 1
    objType = SSProcess.GetInputParameter("要素类型")
    outRange = SSProcess.GetInputParameter("导出范围")
    outPath = SSProcess.GetInputParameter("输出文件")
    If outPath = "" Then
        MsgBox "请填写输出文件路径"
        Exit Sub
    End If

    ' 选择集模式：直接使用当前系统选择，不清空
    If outRange = "选择集" Then
        SSProcess.UpdateSysSelection 0
    Else
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        If objType = "全部" Then
            SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT,LINE,AREA,NOTE"
        Else
            SSProcess.SetSelectCondition "SSObj_Type", "=", objType
        End If
        SSProcess.SelectFilter
    End If

    geoCount = SSProcess.GetSelGeoCount()
    noteCount = SSProcess.GetSelNoteCount()
    totalCount = geoCount + noteCount
    If totalCount = 0 Then
        MsgBox "没有选中任何要素，请先选择要素（选择集模式）或检查要素类型（全部模式）"
        Exit Sub
    End If

    ' 全量扩展属性清单（硬编码，来自 EPS 属性表截图 + 脚本目录经验）
    Dim allAttrStr
    allAttrStr = "FeatureGUID|ZDGUID|ZRZGUID|LCGUID|LJZGUID|HGUID|LJZH|CH|SJC|MYC|HH|SHBW|DYH|HXK|MTDMC|PWLX|SYGN|QSXZ|KZMJ|MJXS|JZMJ|PTGHMC|JJMJLX|GHMJXS|GHJZMJ|MJTJGUID|FMMC|YYGZ|JGGN|JLX|XSTJ|JSCC|SNQMJ|CG|JGJG|ZDCG|ZDGD|ZCGC|ZGGB|JSJG|ZLDG|GD|ZGIDG|IDGD|PMTKL|KH|JGLX|HD|LMTKL|SFJR|JRXS|ZJJRXS|JRMJ|BJRMJ|ZJRMJ|PCXSFS|GHXSFS|JRXSFS|ZJJXFSFS|HNGZ|FSCH|SZ|JSMJSM|GHGHMC|GHGHLX|YeWID|YeWBH|ZL|JZWMC|CeLDW|CeLY|HuiTY|JianCY|BBH|FWJG|FWJGNAME|SCFTJZMJ|FTTDMJ|DYTDMJ|ZCS_LJZ|CLDW|ShenHY|ShenHRQ|CeLRQ|JianCRQ|ZCS_ZRZ|ZCS_ZRZ_LJZ|DJH|TFH|QLR|FangCH|JZDH|TongZH|ZDMJ|JianZZDMJ|JianZZMJ|TuDLYLB|FaZQK|ZongGD|DiEJGD|WaiDMGD|ZRZH|ZDDM|ZhiTRQ|ZhiTY|ShenHR|版本号|地籍号|图幅号|使用权人|土地座落|房产号|原宗地号|统字号|发证面积|建基面积|建筑总面积|地类编码比率|发证情况|宗地号|第二级高度|外地面高度|打印比例|打印纸张|打印范围左上X坐标|纸张方向|纸宽|纸高|上边距|左边距|打印宽|打印高|打印范围左上Y坐标|工程类型|DYBLC|DYFX|DYZZK|DYZZG|DYFWK|DYFWG|DYZSX|DYZSY"

    ' 动态枚举 MemoData 中的额外属性，合并进全量清单
    Dim i
    For i = 0 To geoCount - 1
        memoData = SSProcess.GetSelGeoValue(i, "SSObj_MemoData")
        allAttrStr = CollectAttrNames(memoData, allAttrStr)
    Next
    For i = 0 To noteCount - 1
        memoData = SSProcess.GetSelNoteValue(i, "SSObj_MemoData")
        allAttrStr = CollectAttrNames(memoData, allAttrStr)
    Next

    Dim attrNames, attrCount
    attrNames = Split(allAttrStr, "|")
    attrCount = UBound(attrNames) + 1

    Dim fso, outputFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set outputFile = fso.CreateTextFile(outPath, True)

    Dim header, j
    header = "ID,编码,图层,类型,颜色,面积(㎡),长度(m),注记文字"
    For j = 0 To attrCount - 1
        header = header & "," & attrNames(j)
    Next
    outputFile.WriteLine header

    Dim line
    For i = 0 To geoCount - 1
        line = CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_ID")) & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_Code")) & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_LayerName")) & "," & "地物" & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_Color")) & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_Area")) & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_Length")) & "," & ""
        memoData = SSProcess.GetSelGeoValue(i, "SSObj_MemoData")
        line = line & BuildExtLine(i, memoData, attrNames, attrCount, 0)
        outputFile.WriteLine line
    Next
    For i = 0 To noteCount - 1
        line = CsvEscape(SSProcess.GetSelNoteValue(i, "SSObj_ID")) & "," & CsvEscape(SSProcess.GetSelNoteValue(i, "SSObj_Code")) & "," & CsvEscape(SSProcess.GetSelNoteValue(i, "SSObj_LayerName")) & "," & "注记" & "," & "" & "," & "" & "," & "" & "," & CsvEscape(SSProcess.GetSelNoteValue(i, "SSObj_FontString"))
        memoData = SSProcess.GetSelNoteValue(i, "SSObj_MemoData")
        line = line & BuildExtLine(i, memoData, attrNames, attrCount, 1)
        outputFile.WriteLine line
    Next

    outputFile.Close
    MsgBox "导出完成！共处理 " & totalCount & " 个要素（地物 " & geoCount & "，注记 " & noteCount & "），导出 " & attrCount & " 个属性列，输出文件：" & outPath
End Sub

Function CollectAttrNames(memoData, attrNamesStr)
    If memoData <> "" Then
        Dim arrKV, k, name
        arrKV = Split(memoData, "|")
        For k = 0 To UBound(arrKV) Step 2
            name = arrKV(k)
            If name <> "" Then
                If InStr("|" & attrNamesStr & "|", "|" & name & "|") = 0 Then
                    attrNamesStr = attrNamesStr & "|" & name
                End If
            End If
        Next
    End If
    CollectAttrNames = attrNamesStr
End Function

Function BuildExtLine(idx, memoData, attrNames, attrCount, isNote)
    Dim line, j, v
    line = ""
    For j = 0 To attrCount - 1
        v = GetAttrValue(memoData, attrNames(j))
        If v = "" Then
            If isNote = 0 Then
                v = SSProcess.GetSelGeoValue(idx, "[" & attrNames(j) & "]")
            Else
                v = SSProcess.GetSelNoteValue(idx, "[" & attrNames(j) & "]")
            End If
        End If
        line = line & "," & CsvEscape(v)
    Next
    BuildExtLine = line
End Function

Function GetAttrValue(memoData, attrName)
    GetAttrValue = ""
    If memoData <> "" Then
        Dim arrKV, k
        arrKV = Split(memoData, "|")
        For k = 0 To UBound(arrKV) Step 2
            If arrKV(k) = attrName Then
                If k + 1 <= UBound(arrKV) Then
                    GetAttrValue = arrKV(k + 1)
                End If
                Exit Function
            End If
        Next
    End If
End Function

Function CsvEscape(v)
    If InStr(v, ",") > 0 Or InStr(v, """") > 0 Or InStr(v, Chr(13)) > 0 Or InStr(v, Chr(10)) > 0 Then
        CsvEscape = """" & Replace(v, """", """""") & """"
    Else
        CsvEscape = v
    End If
End Function
