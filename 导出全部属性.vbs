' ============================================================
' 导出全部属性.vbs（全属性版 v5）
' 功能：导出所选要素的全部属性（基本属性 + 全量扩展属性 + 楼层表属性）
' 支持：选择集 / 全部范围，POINT / LINE / AREA / NOTE / 全部
' v4：硬编码全量清单 + MemoData 动态枚举 + 双通道取值
' v5：新增通过 ADO 读取当前工程 EDB 的 [FC_楼层信息属性表]，
'      按 FeatureGUID 匹配合并楼层信息（109 字段），自动去重
' 说明：仅用最基础 VBScript 语法，不依赖 Scripting.Dictionary 等
' ============================================================

' ---- 楼层表全局数据（LoadFloorTable 填充）----
Dim floorLoaded
Dim floorCount
Dim floorFieldCount
Dim floorOutCount
Dim floorFieldNames(200)
Dim floorGuids(300)
Dim floorValues(300, 200)
Dim floorOutNames(200)
Dim floorOutCols(200)

Sub OnClick()
    MsgBox "脚本已启动，开始执行（v5 支持楼层表合并）"
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "要素类型", "全部", 0, "POINT,LINE,AREA,NOTE,全部", ""
    SSProcess.AddInputParameter "导出范围", "选择集", 0, "选择集,全部", ""
    SSProcess.AddInputParameter "合并楼层表属性", "是", 0, "是,否", ""
    SSProcess.AddInputParameter "输出文件", "C:\全部属性导出.csv", 0, "", ""
    res = SSProcess.ShowInputParameterDlg("导出全部属性")
    If res = 0 Then
        Exit Sub
    End If
    SSProcess.UpdateScriptDlgParameter 1
    objType = SSProcess.GetInputParameter("要素类型")
    outRange = SSProcess.GetInputParameter("导出范围")
    mergeFloor = SSProcess.GetInputParameter("合并楼层表属性")
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

    ' v5：合并楼层表属性（成功则填充 floorOutNames / floorOutCount）
    If mergeFloor = "是" Then
        LoadFloorTable allAttrStr
    End If

    Dim fso, outputFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set outputFile = fso.CreateTextFile(outPath, True)

    ' 表头：基本属性 + 扩展属性 + 楼层表字段
    Dim header, j
    header = "ID,编码,图层,类型,颜色,面积(㎡),长度(m),注记文字"
    For j = 0 To attrCount - 1
        header = header & "," & attrNames(j)
    Next
    If floorLoaded = 1 Then
        For j = 0 To floorOutCount - 1
            header = header & "," & floorOutNames(j)
        Next
    End If
    outputFile.WriteLine header

    Dim line, featureGuid
    For i = 0 To geoCount - 1
        line = CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_ID")) & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_Code")) & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_LayerName")) & "," & "地物" & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_Color")) & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_Area")) & "," & CsvEscape(SSProcess.GetSelGeoValue(i, "SSObj_Length")) & "," & ""
        memoData = SSProcess.GetSelGeoValue(i, "SSObj_MemoData")
        line = line & BuildExtLine(i, memoData, attrNames, attrCount, 0)
        featureGuid = GetAttrValue(memoData, "FeatureGUID")
        If featureGuid = "" Then
            featureGuid = SSProcess.GetSelGeoValue(i, "[FeatureGUID]")
        End If
        line = line & BuildFloorLine(featureGuid)
        outputFile.WriteLine line
    Next
    For i = 0 To noteCount - 1
        line = CsvEscape(SSProcess.GetSelNoteValue(i, "SSObj_ID")) & "," & CsvEscape(SSProcess.GetSelNoteValue(i, "SSObj_Code")) & "," & CsvEscape(SSProcess.GetSelNoteValue(i, "SSObj_LayerName")) & "," & "注记" & "," & "" & "," & "" & "," & "" & "," & CsvEscape(SSProcess.GetSelNoteValue(i, "SSObj_FontString"))
        memoData = SSProcess.GetSelNoteValue(i, "SSObj_MemoData")
        line = line & BuildExtLine(i, memoData, attrNames, attrCount, 1)
        featureGuid = GetAttrValue(memoData, "FeatureGUID")
        If featureGuid = "" Then
            featureGuid = SSProcess.GetSelNoteValue(i, "[FeatureGUID]")
        End If
        line = line & BuildFloorLine(featureGuid)
        outputFile.WriteLine line
    Next

    outputFile.Close
    MsgBox "导出完成！共处理 " & totalCount & " 个要素（地物 " & geoCount & "，注记 " & noteCount & "），扩展属性 " & attrCount & " 列，楼层字段 " & floorOutCount & " 列，输出文件：" & outPath
End Sub

' v5：通过 ADO 读取当前工程 EDB 的 [FC_楼层信息属性表]，加载到内存数组
Sub LoadFloorTable(currentAttrStr)
    Dim projFile, sysPath, fsoCheck, strConnect, adoConnection, adoRs
    Dim ff, fr, fv, fo
    floorLoaded = 0
    floorCount = 0
    floorFieldCount = 0
    floorOutCount = 0

    ' 获取当前工程文件（.edb）；若返回不含路径分隔符则用系统路径拼接
    projFile = SSProcess.GetProjectFileName()
    If InStr(projFile, "\") = 0 And InStr(projFile, "/") = 0 Then
        sysPath = SSProcess.GetSysPathName(5)
        If Right(sysPath, 1) <> "\" Then
            sysPath = sysPath & "\"
        End If
        projFile = sysPath & projFile
    End If

    Set fsoCheck = CreateObject("Scripting.FileSystemObject")
    If Not fsoCheck.FileExists(projFile) Then
        MsgBox "未找到工程文件：" & projFile & Chr(13) & "跳过楼层表属性合并，继续导出"
        Exit Sub
    End If

    ' 连接 EDB（Access 格式），读取 [FC_楼层信息属性表]
    strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & projFile & ";"
    Set adoConnection = CreateObject("ADODB.Connection")
    adoConnection.Open strConnect
    Set adoRs = CreateObject("ADODB.Recordset")

    ' 第一次：SELECT TOP 1 * 获取字段名
    adoRs.Open "SELECT TOP 1 * FROM [FC_楼层信息属性表]", adoConnection, 0
    If adoRs.Fields.Count > 200 Then
        floorFieldCount = 200
    Else
        floorFieldCount = adoRs.Fields.Count
    End If
    For ff = 0 To floorFieldCount - 1
        floorFieldNames(ff) = adoRs.Fields(ff).Name
    Next
    adoRs.Close

    ' 第二次：SELECT * 读取全部记录到内存数组
    Set adoRs = Nothing
    Set adoRs = CreateObject("ADODB.Recordset")
    adoRs.Open "SELECT * FROM [FC_楼层信息属性表]", adoConnection, 0
    floorCount = 0
    Do While Not adoRs.Eof
        If floorCount < 300 Then
            floorGuids(floorCount) = NormalizeGuid(CStr(adoRs.Fields("FeatureGUID").Value))
            For fr = 0 To floorFieldCount - 1
                fv = adoRs.Fields(fr).Value
                If IsNull(fv) Then
                    floorValues(floorCount, fr) = ""
                Else
                    floorValues(floorCount, fr) = CStr(fv)
                End If
            Next
            floorCount = floorCount + 1
        End If
        adoRs.MoveNext
    Loop
    adoRs.Close
    adoConnection.Close
    Set adoRs = Nothing
    Set adoConnection = Nothing

    ' 字段去重：跳过与现有清单重复的字段（FeatureGUID、ZDGUID 等），构建输出列
    floorOutCount = 0
    For fo = 0 To floorFieldCount - 1
        If InStr("|" & currentAttrStr & "|", "|" & floorFieldNames(fo) & "|") = 0 Then
            floorOutNames(floorOutCount) = floorFieldNames(fo)
            floorOutCols(floorOutCount) = fo
            floorOutCount = floorOutCount + 1
        End If
    Next

    floorLoaded = 1
End Sub

' v5：按 FeatureGUID 匹配楼层表记录，追加楼层字段值（含前导逗号）
Function BuildFloorLine(featureGuid)
    Dim line, guid, rowIdx, c, k
    line = ""
    If floorLoaded = 1 Then
        guid = NormalizeGuid(featureGuid)
        rowIdx = -1
        For k = 0 To floorCount - 1
            If floorGuids(k) = guid Then
                rowIdx = k
                Exit For
            End If
        Next
        For c = 0 To floorOutCount - 1
            If rowIdx >= 0 Then
                line = line & "," & CsvEscape(floorValues(rowIdx, floorOutCols(c)))
            Else
                line = line & ","
            End If
        Next
    End If
    BuildFloorLine = line
End Function

' v5：GUID 归一化为小写无花括号格式
Function NormalizeGuid(s)
    NormalizeGuid = LCase(Replace(Replace(s, "{", ""), "}", ""))
End Function

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
