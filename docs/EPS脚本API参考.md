---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 95ce5665a2763fbde4f792b4da682b54_9c49c3a2998211f19467525400287e28
    ReservedCode1: FCACNBfkkuYVKPDuY5l1fDtjs9JXB4UOjJGKgov/YjoYdb37eCxYUDrOfwOv+pC+uTwHKNTWbQzTqqV+rJL8/vYtLhj1Skh+L7sVAJqAE056Qf1qsKBUJnGZAVpXB7lNqxg+SBGLn9OgdJYp3pWu+TAcB/SODhD+PzxVL+AlYJRG3XjmWYP6s/FYnMM=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 95ce5665a2763fbde4f792b4da682b54_9c49c3a2998211f19467525400287e28
    ReservedCode2: FCACNBfkkuYVKPDuY5l1fDtjs9JXB4UOjJGKgov/YjoYdb37eCxYUDrOfwOv+pC+uTwHKNTWbQzTqqV+rJL8/vYtLhj1Skh+L7sVAJqAE056Qf1qsKBUJnGZAVpXB7lNqxg+SBGLn9OgdJYp3pWu+TAcB/SODhD+PzxVL+AlYJRG3XjmWYP6s/FYnMM=
---

# EPS 脚本语言 API 参考

> 来源：EPS 官方 CHM 帮助文档（EpsScipt-20161030.chm）
>
> 本文档从 CHM 解压的 611 个 htm 页面（GBK 编码）中提取整理，覆盖 **8 个对象、571 个 API 条目**，按官方文档分类组织。
>
> 每个条目包含：函数名、功能说明、语法、参数说明（含属性字段表）、详细说明、示例代码。

---

## 目录

### SSProcess 对象参考（428 条）

- [图形编辑操作](#图形编辑操作)（110 条）
- [地模处理](#地模处理)（8 条）
- [工程管理函数](#工程管理函数)（9 条）
- [角度函数](#角度函数)（14 条）
- [数据检查](#数据检查)（9 条）
- [数据库操作](#数据库操作)（24 条）
- [数据整理](#数据整理)（62 条）
- [数据转换](#数据转换)（12 条）
- [图幅操作](#图幅操作)（18 条）
- [外部函数](#外部函数)（15 条）
- [系统设置](#系统设置)（88 条）
- [选择集操作](#选择集操作)（17 条）
- [坐标转换函数](#坐标转换函数)（22 条）
- [数学函数](#数学函数)（10 条）
- [系统设置函数](#系统设置函数)（6 条）
- [系统函数](#系统函数)（4 条）

### Function 对象参考（30 条）

- [系统消息函数](#系统消息函数)（30 条）

### SSView 对象参考（23 条）

- [屏幕绘制函数](#屏幕绘制函数)（23 条）

### SSArray 对象参考（10 条）

- [数组函数](#数组函数)（10 条）

### SSParmeter 对象参考（11 条）

- [共享参数函数](#共享参数函数)（11 条）

### SSFunc 对象参考（36 条）

- [扩展函数库](#扩展函数库)（36 条）

### SSProject 对象参考（32 条）

- [工程管理函数](#工程管理函数)（32 条）

### 常用脚本函数 对象参考（1 条）

- [获取指定目录下指定扩展名的所有文件](#获取指定目录下指定扩展名的所有文件)（1 条）

---

## SSProcess 对象参考

### 图形编辑操作

#### AddClipBoardObjToMap 粘贴板对象加入到地图

**函数**：`AddClipBoardObjToMap`

**功能**：粘贴板对象加入到地图。

**语法**：

```vbs
SSProcess.AddClipBoardObjToMap(insertPointX , insertPointY)
```

**参数**：

- insertPointX 插入位置Y的增量(米)
- insertPointY 插入位置X的增量(米)

**说明**：

使用此函数可以将粘贴板对象添加到地图中，并且指定偏移量，使对象偏移到指定位置。

**示例**：

```vbs
Sub OnClick()
'选择地物
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
'将选择集对象拷贝到粘贴板
SSProcess.SelectionObjToClipBoard
'粘帖板对象添加到地图
SSProcess.AddClipBoardObjToMap 100, 100
End Sub
```


#### AddNewObjPoint 添加空间点到新图形对象

**函数**：`AddNewObjPoint`

**功能**：添加空间点到新图形对象。

**语法**：

```vbs
SSProcess.AddNewObjPoint( x, y, z, point type, name )
```

**参数**：

- x 空间点X坐标
- y 空间点Y坐标
- z 空间点Z坐标
| 参数 | 说明 |
|---|---|
| pointtype | 空间点类型 |
| 0 | 普通点 |
| 1 | 实测点 |
| 2 | 标注点 |
| 4 | 建模点 |
| 16 | 转向点 |
| 32 | 断点 |
| 64 | 直线起点 |
| 130 | 曲线起点 |
| 192 | 圆弧中点 |
| 2048 | 分段点 |
| 4096 | 依附点 |
| 16384 | 平滑点 |
| -32768 | 特征点 |
- name 空间点名

**说明**：

执行该函数，将添加空间点到当前创建的对象。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateNewObj 0
SSProcess.SetNewObjValue "SSObj_Code", "0"
SSProcess.SetNewObjValue "SSObj_LayerName", "DEFAULT"
SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,0,0)"
SSProcess.SetNewObjValue "SSObj_LineType", "0"
SSProcess.AddNewObjPoint 1000.01, 1020.34, 23.45, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### AddNewObjToSaveObjList 添加新创建图形对象到存库列表

**函数**：`AddNewObjToSaveObjList`

**功能**：添加新创建图形对象到存库列表。

**语法**：

```vbs
SSProcess.AddNewObjToSaveObjList()
```

**说明**：

执行该函数，将把当前新创建对象添加到存库地物列表中。执行该函数后,需再增加SaveBufferObjToDatabase函数,才能把新创建对象保存到数据库。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateNewObjByClass "0"
SSProcess.SetNewObjValue "SSObj_FontString", "上地路"
SSProcess.AddNewObjPoint 1000.01, 1000.34, 22.45, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### AddNewObjToSelObjList 添加新创建图形对象到选择集列表

**函数**：`AddNewObjToSelObjList`

**功能**：添加新创建对象到选择集列表

**语法**：

```vbs
SSProcess.AddNewObjToSelObjList()
```

**说明**：

执行该函数将新创建的图形添加到选择集列表

**示例**：

```vbs
Sub OnClick()
'创建新的图形
SSProcess.CreateNewObj 0
SSProcess.SetNewObjValue "SSObj_Code", "0"
SSProcess.SetNewObjValue "SSObj_LayerName", "DEFAULT"
SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,0,0)"
SSProcess.SetNewObjValue "SSObj_LineType", "0"
SSProcess.AddNewObjPoint 1000.01, 1020.34, 23.45 , 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
'将新创建的图形添加到选择集列表
SSProcess.AddNewObjToSelObjList
GeoCount=SSProcess.GetSelGeoCount()
End Sub
```


#### AddSelGeoToSaveGeoList 添加当前选择集指定地物到存库列表

**函数**：`AddSelGeoToSaveGeoList`

**功能**：添加当前选择集指定地物到存库列表。

**语法**：

```vbs
SSProcess.AddSelGeoToSaveGeoList( index )
```

**参数**：

- index 选择集地物的索引号,从0开始

**说明**：

执行该函数，将把编辑修改过的选择集对象添加到存库地物列表中。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
If geocount > 0 Then
'锁定数据库
SSProcess.SetMapStatus 1, 2
Dim x, y, z, pointtype, name
For i=0 To geocount-1
'修改基本信息
SSProcess.SetSelGeoValue i, "SSObj_Code", "0"
SSProcess.SetSelGeoValue i, "SSObj_LayerName","DEFAULT"
'加入存库对象列表
SSProcess.AddSelGeoToSaveGeoList i
Next
'开锁数据库
SSProcess.SetMapStatus 0, 2
'存库对象列表保存到数据库
SSProcess.SaveBufferObjToDatabase
End If
End Sub
```


#### AddSelNoteToSaveNoteList 添加当前选择集指定注记到存库列表

**函数**：`AddSelNoteToSaveNoteList`

**功能**：添加当前选择集指定注记到存库列表。

**语法**：

```vbs
SSProcess.AddSelNoteToSaveNoteList( index )
```

**参数**：

- index 选择集注记的索引号,从0开始

**说明**：

执行该函数，将把编辑修改过的选择集注记添加到存库注记列表中。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount()
If notecount > 0 Then
'锁定数据库
SSProcess.SetMapStatus 1, 2
For i=0 To notecount-1
'修改基本信息
SSProcess.SetSelNoteValue i, "SSObj_FontClass", "0"
SSProcess.SetSelNoteValue i, "SSObj_LayerName","DEFAULT"
'加入存库对象列表
SSProcess.AddSelNoteToSaveNoteList i
Next
'开锁数据库
SSProcess.SetMapStatus 0, 2
'存库对象列表保存到数据库
SSProcess.SaveBufferObjToDatabase
End If
End Sub
```


#### AddSeqRecord 增加对象序列属性

**函数**：`AddSeqRecord`

**功能**：增加对象序列属性

**语法**：

```vbs
SSProcess.AddSeqRecord(geoID,fieldNames,filedValues,seqTableName)
```

**参数**：

- geoID 地物ID
- fieldNames 序列属性表字段名称
- fieldValues 序列属性表字段属性值
- seqTableName 序列属性表名称

**说明**：

执行该函数，增加地物 ID指定序列表指定字段记录

**示例**：

```vbs
Sub OnClick()
’执行下面函数将对 ID为 1的地物 ZD_权利人序列表增加 QLRMC为张三的一条记录
SSProcess.AddSeqRecord 1, "QLRMC","张三","ZD_权利人"
End Sub
```


#### ChangeSelectionObj 执行修改选择集操作

**函数**：`ChangeSelectionObj`

**功能**：执行修改选择集操作。

**语法**：

```vbs
SSProcess.ChangeSelectionObj()
```

**说明**：

该函数根据SetChangeSelectionParameter函数设定的选择集修改条件，执行实际修改操作。
ChangeSelectionObjAttr函数一次只能修改一个属性字段，而ChangeSelectionObj函数则可以一次执行多个属性字段的修改，执行效率要高于ChangeSelectionObjAttr函数。

**示例**：

```vbs
下面示例把选择集中[结构]=11,[层数]=12的对象的颜色改为3号颜色，把[结构]=10,[层数]=11的对象的颜色改为2号颜色，把[结构]=8,[层数]=9的对象的颜色改为5号颜色。
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
SSProcess.ClearChangeSelectionParameter
SSProcess.SetChangeSelectionObjParameter "[结构]=11,[层数]=12", "SSObj_Color= COLORNO(3)"
SSProcess.SetChangeSelectionObjParameter "[结构]=10,[层数]=11", "SSObj_Color= COLORNO(2)"
SSProcess.SetChangeSelectionObjParameter "[结构]=8,[层数]=9", "SSObj_Color= COLORNO(5)"
SSProcess.ChangeSelectionObj
End Sub
```


#### ChangeSelectionObjAttr 修改选择集内对象的属性

**函数**：`ChangeSelectionObjAttr`

**功能**：修改选择集内对象的属性。

**语法**：

```vbs
SSProcess.ChangeSelectionObjAttr( attrField, attrValue)
```

**参数**：

- attrField 要修改的属性名称，见下表，如果要修改关联地物的属性，则用SSSubObj_或SSSubExtAttr表示，如（SSSubObj_Code,[SSSubExtAttr_层数]）
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_GroupID | 对象组 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 0-256 颜色号、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_Angle | 角度 , 以弧度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_FontWidth | 字宽，以图上 0.01 毫米为单位 |
| SSObj_FontHeight | 字高，以图上 0.01 毫米为单位 |
| SSObj_FontName | 字体名 |
| SSObj_FontClass | 注记分类号 |
| SSObj_FontAlignment | 对齐方式，包括 0 （中心）、 1 （左上角）、 2 （左下角）、 3 （左中）、 4 （右上角）、 5 （右中）、 6 （右下角）、 7 （上中）、 8 （下中） |
| SSObj_FontPosType | 注记排列方式 , 包括 0 （均分）、 1 （散列） |
| SSObj_FontDirection | 字头朝向，包括 0( 正东 ) 、 1( 正北 ) 、 2( 正西 ) 、 3( 正南 ) 、 4( 切线方向 ) 、 5( 法线方向 ) |
| SSObj_FontWordAngle | 字角度 , 以度为单位 |
| SSObj_FontStringAngle | 串角度 , 以度为单位 |
| SSObj_FontWeight | 字重 |
| SSObj_FontIlaticAngle | 倾斜，包括 0 （不倾斜）、 -1 （左斜）、 1 （右斜） |
| SSObj_FontDownAngle | 耸肩，包括 0 （不耸肩）、 -1 （左耸肩）、 1 （右耸肩） |
| SSObj_FontUnderLine | 下划线，包括 0 （无下划线）、 1 （有下划线） |
| SSObj_FontInterval | 字隔，以图上 0.01 毫米为单位 |
| SSObj_FontString | 注记内容 |
| SSObj_FontStringCount | 注记内容字符个数（汉字算两个字符） |
| SSObj_X | 空间点 X 坐标 |
| SSObj_Y | 空间点 Y 坐标 |
| SSObj_Z | 空间点 Z 坐标 |
| SSObj_PointName | 空间点点名 |
| SSObj_PointType | 空间点首点点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_X(index) | 指定索引空间点 X 坐标 |
| SSObj_Y(index) | 指定索引空间点 Y 坐标 |
| SSObj_Z(index) | 指定索引空间点 Z 坐标 |
| SSObj_PointName(index) | 指定索引空间点点名 |
| SSObj_PointType(index) | 指定索引空间点点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_Reverse | 地物方向， 0 （正向）、 1 （反向） |
| SSObj_DrawAreaLabelStatus | 面心点标注状态 , 0( 不注 ) 1( 标注 ) |
| SSObj_Status | 地物状态 , 正值为设置 , 负值为取消 e_Valid = 0x0001, // 有效 e_Explode = 0x0002, // 是否已打散 e_Screen = 0x0004, // 是否被屏蔽 e_Duplicate = 0x0008, // 是否被复制品 e_IsQuote = 0x0010, // 是否索引 e_Hidden = 0x0020, // 是否隐藏 e_IsSelected = 0x0040, // 是否已选择 e_Editable = 0X0080, // 是否可编辑 e_Highlight = 0X0100, // 是否加亮 e_Gray = 0X0200, // 是否变灰 e_Edited = 0X0400, // 是否被修改 e_Refrence = 0X0800, // 参考图形 e_Opaque = 0X1000, // 是否压盖 |
| SSObj_GraphicInfo(flag) | 图形特征 ,flag 可以是以下值 : 25 面心标志点显示开关及相对多边形焦点偏移量 , 格式 :" 是否显示 (0 不显示 ,1 显示 ), 东偏 , 北偏 ",( 米为单位 ) 26 个性化符号描述 , 格式 :" 是否替换现在符号描述 (0 不替换 ,1 替换 ); 符号描述语句 ( 多个语句用 ; 号分隔 )" 101 图形缩放比例 , 格式 : "X 向比例 ,Y 向比例 ,Z 向比例 " ( 只支持整数 , 单位为 0.00001 米 ) |
| SSObj_HatchStyle | 面填充符号样式 0 符号填充 1 不填充 2 颜色填充 |
| [ExtAttr] | 扩展属性字段 |
- attrValue 修改值
- 修改值可以是任意值或attrField中指定的属性,或二者的组合(如: SSObj_LayerName123)，如果是替换值中的某部分内容，则可加{替换值,替换新值}标志(如: SSObj_FontString{砖,砼}, 替换注记中的砖为砼字) 。
- 修改值同时支持回调函数方式,由脚本定义处理函数来控制修改，格式为： CallBackFunc_函数名称(参数)，（如："CallBackFunc_TrimString(SSObj_FontString)"），如果同时有拓扑关联对象,则允许指定关联对象的属性， 以SSSubObj_和SSSubExtAttr_开头,如"CallBackFunc_CheckString(SSObj_ FontString, SSSubObj_FontString)"，修改值的来源可以是某个选择条件组的结果中的属性内容，格式为：SelectionGroup_选择组名称(参数) （如："SelectionGroup_房屋内属性点(CallBackFunc_TrimString([层数]))"） 。

**说明**：

执行该函数，将修改当前选择集对象为指定属性值，该函数允许连续多次执行，执行一次修改一次属性值。如果使用了回调函数，则必须在脚本中定义回调函数的实现。
如果修改的内容是SSObj_Code或SSObj_LayerName,则选择定一次，只能执行一次修改，修改完后必须马上执行清空选择集操作。

**示例**：

```vbs
下面示例把含有“砖”字的注记中的“砖”字替换为砼字,并把颜色修改为RGB(255,255,0)：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_FontString", "LIKE", "砖"
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
SSProcess.ChangeSelectionObjAttr "SSObj_FontString", "砼"
SSProcess.ChangeSelectionObjAttr "SSObj_Color", "RGB(255,255,0)"
End Sub
```


#### ClearChangeSelectionParameter 清空选择集修改参数

**函数**：`ClearChangeSelectionParameter`

**功能**：清空选择集修改参数。

**语法**：

```vbs
SSProcess.ClearChangeSelectionParameter()
```

**说明**：

执行该函数，将清空由函数SetChangeSelectionObjParameter设置的择集修改参数。


#### ClearSelectionObjOnFramePoint 删除选择集对象中落在图幅边线上接边点

**函数**：`ClearSelectionObjOnFramePoint`

**功能**：删除选择集对象中落在图幅边线上接边点。

**语法**：

```vbs
SSProcess.ClearSelectionObjOnFramePoint (onlineLimit,offset)
```

**参数**：

- onlineLimit 限差
- Offset

**说明**：

该函数删除选择集对象中落在图幅边线上的接边点。


#### CloneSelGeo 复制当前选择集指定地物

**函数**：`CloneSelGeo`

**功能**：复制当前选择集指定地物

**语法**：

```vbs
SSProcess.CloneSelGeo(index)
```

**参数**：

- index 选择集中地物索引号

**说明**：

该函数用于复制当前选择集指定地物。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "=", "0"
SSProcess.SelectFilter
Geocount = SSProcess.GetSelGeoCount
If Geocount > 0 Then
SSProcess. CloneSelGeo 0
End If
End Sub
```


#### CloneSelNote 复制当前选择集指定注记

**函数**：`CloneSelNote`

**功能**：复制当前选择集指定注记

**语法**：

```vbs
SSProcess.CloneSelNote(index)
```

**参数**：

- index 对象序号

**说明**：

该函数用于复制当前选择集指定注记。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==" , "NOTE"
SSProcess.SelectFilter
Notecount = SSProcess.GetSelNoteCount()
If Notecount > 0 Then
SSProcess.CloneSelNote 0
End If
End Sub
```


#### CopyAreaAttrToLabel 面属性转换到 LABEL 点属性

**函数**：`CopyAreaAttrToLabel`

**功能**：面属性转换到LABEL点属性。

**语法**：

```vbs
SSProcess.CopyAreaAttrToLabel( areaLayer, labelLayer)
```

**参数**：

- areaLayer 面图层名
- labelLayer LABEL点图层名

**说明**：

执行该函数，将把面属性转换到LABEL点属性，转换时，只转换LABEL点属性表与面属性表字段名称相同的字段。

**示例**：

```vbs
下面示例把“ZDBZM”层的面属性转换到"ZDBZD"层的点属性：
Sub OnClick()
SSProcess.CopyAreaAttrToLabel "ZDBZM", "ZDBZD"
End Sub
```


#### CopyLabelAttrToArea LABEL 点属性转换到面属性

**函数**：`CopyLabelAttrToArea`

**功能**：LABEL点属性转换到面属性。

**语法**：

```vbs
SSProcess.CopyLabelAttrToArea( labelLayer, areaLayer)
```

**参数**：

- labelLayer LABEL点图层名
- areaLayer 面图层名

**说明**：

执行该函数，将把LABEL点属性转换到面属性，转换时，只转换LABEL点属性表与面属性表字段名称相同的字段。

**示例**：

```vbs
下面示例把"ZDBZD"层的点属性转换到“ZDBZM”层的面属性：
Sub OnClick()
SSProcess.CopyLabelAttrToArea "ZDBZD", "ZDBZM"
End Sub
```


#### CreateNewObj 创建新图形对象

**函数**：`CreateNewObj`

**功能**：创建新图形对象。

**语法**：

```vbs
SSProcess.CreateNewObj( type )
```

**参数**：

- type 创建新对象类型，包括 0（点） 1（线） 2（面） 3（注记）

**说明**：

执行该函数，将创建一个指定类型的新图形对象作为当前操作的对象。随后可执行SetNewObjValue函数设置新对象的属性值，执行AddNewObjPoint函数添加新对象的空间点，执行AddNewObjToSaveObjList函数把新对象加到存库对象列表，执行SaveBufferObjToDatabase函数保存存库列表到数据库。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateNewObj 0
SSProcess.SetNewObjValue "SSObj_Code", "0"
SSProcess.SetNewObjValue "SSObj_LayerName", "DEFAULT"
SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,0,0)"
SSProcess.SetNewObjValue "SSObj_LineType", "0"
SSProcess.AddNewObjPoint 1020.01, 1030.34, 22.45, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### CreateNewObjByClass 根据注记分类号创建新图形对象

**函数**：`CreateNewObjByClass`

**功能**：根据注记分类号创建新图形对象。

**语法**：

```vbs
SSProcess.CreateNewObjByClass( fontclass )
```

**参数**：

- fontclass 创建注记的分类号

**说明**：

执行该函数，将创建一个指定分类号的新注记对象作为当前操作的对象，新建的注记会根据Notetemplate表的注记分类号定义自动填充层名、颜色、字高、字宽等属性。随后可执行SetNewObjValue函数设置新对象的属性值，执行AddNewObjPoint函数添加新对象的空间点，执行AddNewObjToSaveObjList函数把新对象加到存库对象列表，执行SaveBufferObjToDatabase函数保存存库列表到数据库。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateNewObjByClass "1"
SSProcess.SetNewObjValue "SSObj_FontString", "上地路"
SSProcess.AddNewObjPoint 1000.01, 1000.34, 22.45, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### CreateNewObjByCode 根据地物编码创建新图形对象

**函数**：`CreateNewObjByCode`

**功能**：根据地物编码创建新图形对象。

**语法**：

```vbs
SSProcess.CreateNewObjByCode( code )
```

**参数**：

- code 创建地物的编码

**说明**：

执行该函数，将创建一个指定编码的新图形对象作为当前操作的对象，新建的对象会根据FeatureCode表的编码定义自动填充层名、颜色、线型、线宽等属性。随后可执行SetNewObjValue函数设置新对象的属性值，执行AddNewObjPoint函数添加新对象的空间点，执行AddNewObjToSaveObjList函数把新对象加到存库对象列表，执行SaveBufferObjToDatabase函数保存存库列表到数据库。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateNewObjByCode 0
SSProcess.AddNewObjPoint 1000.01, 1000.34, 22.45, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### DeleteLayer 删除图层

**函数**：`DeleteLayer`

**功能**：删除图层。

**语法**：

```vbs
SSProcess.DeleteLayer(layerName)
```

**参数**：

- layerName 要删除的图层名称

**说明**：

执行该函数，将先删除图层内的所有对象，后删除图层。

**示例**：

```vbs
下面示例删除"居民地"图层：
Sub OnClick()
SSProcess.DeleteLayer "居民地"
End Sub
```


#### DeleteNewObj 删除新创建图形对象

**函数**：`DeleteNewObj`

**功能**：DeltetNewObj函数

**语法**：

```vbs
SSProcess.DeleteNewObj()
```

**说明**：

执行该函数，将删当前新创建对象。

**示例**：

```vbs
下面示例添加新对象到数据库：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateNewObj 0
SSProcess.SetNewObjValue "SSObj_Code", "0"
SSProcess.SetNewObjValue "SSObj_LayerName", "DEFAULT"
SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,0,0)"
SSProcess.SetNewObjValue "SSObj_LineType", "0"
SSProcess.AddNewObjPoint 1000.01, 1020.34, 23.45, 0, ""
SSProcess.DeleteNewObj
End Sub
```


#### DeleteObject 删除图形对象

**函数**：`DeleteObject`

**功能**：删除图形对象。

**语法**：

```vbs
SSProcess.DeleteObject (geoID)
```

**参数**：

- geoID 要删除的图形ID

**说明**：

执行该函数，将删除指定ID的地物。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
If geocount > 0 Then
For i=0 To geocount-1
ids = SSProcess.GetSelGeoValue(i, "SSObj_ID")
SSProcess.DeleteObject ids
SSProcess.RefreshView
Next
End If
End Sub
```


#### DeleteObjectPoint 删除图形对象的空间点

**函数**：`DeleteObjectPoint`

**功能**：删除图形对象的空间点。

**语法**：

```vbs
SSProcess.DeleteObjectPoint (geoID,pointIndex,delPointCount)
```

**参数**：

- geoID 图形对象ID
- pointIndex 图形对象的空间点索引号，从0开始
- delPointCount 要删除的空间点数

**说明**：

该函数为图形编辑操作函数，执行后将从指定图形对象的第pointIndex个空间点开始向后删除delPointCount个空间点。

**示例**：

```vbs
下面示例是获取所有线状地物的空间点数，并从线地物的第1个空间点开始往后删除1个空间点：
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
If geocount > 0 Then
For i=0 To geocount-1
ids = SSProcess.GetSelGeoValue(i, "SSObj_ID")
pCount =SSProcess.GetSelGeoPointCount(i)
SSProcess.DeleteObjectPoint ids , 0, 1
SSProcess.RefreshView
p =SSProcess.GetSelGeoPointCount(i)
Next
End If
End Sub
```


#### DeleteSelectionObj 删除选择集对象

**函数**：`DeleteSelectionObj`

**功能**：删除选择集内的所有对象。

**语法**：

```vbs
SSProcess.DeleteSelectionObj()
```

**说明**：

执行该函数，将删除选择集内的所有对象。

**示例**：

```vbs
下面示例删除编码等于0并且点数少于2个点的对象：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "0"
SSProcess.SetSelectCondition "SSObj_PointCount", "<", "2"
SSProcess.SelectFilter
SSProcess.DeleteSelectionObj
End Sub
```


#### DeleteSelGeoPoint 删除当前选择集指定地物的空间点

**函数**：`DeleteSelGeoPoint`

**功能**：删除当前选择集指定地物的空间点。

**语法**：

```vbs
SSProcess.DeleteSelGeoPoint( index,pointindex)
```

**参数**：

- index 选择集地物的索引号,从0开始
- pointindex 该地物空间点索引号，从0开始

**说明**：

执行该函数，将删除当前选择集对象中指定地物的第index+1个空间点。

**示例**：

```vbs
下面示例删除当前选择集对象中指定地物的第2个空间点：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount-1
pointcount = SSProcess.DeleteSelGeoPoint(i,1)
Next
End Sub
```


#### DeleteSelNotePoint 删除当前选择集指定注记的空间点

**函数**：`DeleteSelNotePoint`

**功能**：删除当前选择集指定注记的空间点。

**语法**：

```vbs
SSProcess.DeleteSelNotePoint( index,pointindex )
```

**参数**：

- index 选择集注记的索引号,从0开始
- pointindex 该注记的空间点号索引

**说明**：

执行该函数，将删除当前选择集对象中指定注记的第index+1个空间点。

**示例**：

```vbs
下面示例删除当前选择集对象中指定注记的第2个空间点：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
Notecount = SSProcess.GetSelNoteCount
For i=0 To Notecount -1
SSProcess.DeleteSelNotePoint i,1
Next
End Sub
```


#### DeleteSeqRecord 删除对象序列属性

**函数**：`DeleteSeqRecord`

**功能**：删除对象序列属性。

**语法**：

```vbs
SSProcess.DeleteSeqRecord(geoID,seqID,seqTableName)
```

**参数**：

- geoID 地物ID
- seqID 对象序列属性表记录行ID
- seqTableName 对象序列属性表名

**说明**：

该函数为图形编辑操作函数，用于删除指定对象的序列属性。


#### DelSelGeo 删除当前选择集指定地物

**函数**：`DelSelGeo`

**功能**：删除当前选择集指定地物。

**语法**：

```vbs
SSProcess.DelSelGeo(index)
```

**参数**：

- index 选择集地物的索引号,从0开始

**说明**：

执行该函数，将删除当前选择集中指定的地物。

**示例**：

```vbs
下面示例删除当前选择集指定地物：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount-1
SSProcess.DelSelGeo i
Next
SSProcess.ClearSelection
End Sub
```


#### DelSelNote 删除当前选择集指定注记

**函数**：`DelSelNote`

**功能**：删除当前选择集指定注记。

**语法**：

```vbs
SSProcess.DelSelNote( index )
```

**参数**：

- index 选择集注记的索引号,从0开始

**说明**：

执行该函数，将删除当前选择集中指定的注记。

**示例**：

```vbs
下面示例删除当前选择集指定注记：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount
For i=0 To notecount-1
SSProcess.DelSelNote i
Next
End Sub
```


#### ExplodeObj 对象打散

**函数**：`ExplodeObj`

**功能**：对象打散。

**语法**：

```vbs
SSProcess.ExplodeObj(objID,explodeWay,delsrcObj,callbackFunctionName)
```

**参数**：

- objID 对象ID
- explodeWay 符号打散方式 0（自动打散）、1（根据编码表设定打散）
- delSrcObj 是否删除原地物 0（不删除）、1（删除）
- callbackFunctionName 回调函数（可为空值）

**说明**：

执行该函数，可打散指定地物对象

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "4503023"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
For i=0 To geocount- 1
'写基本信息
sid = SSProcess.GetSelGeoValue(i, "SSObj_ID")
'符号打散方式。 0（自动打散）、 1（根据编码表设定打散）
SSProcess.ExplodeObj sid, 0, 1, ""
Next
End Sub
```


#### ExplodeSelectionObj 打散选择集对象

**函数**：`ExplodeSelectionObj`

**功能**：打散选择集内的所有对象。

**语法**：

```vbs
SSProcess.ExplodeSelectionObj( explodeWay, delSrcObj, callbackFunctionName)
```

**参数**：

- explodeWay 打散方式，0（按图形打散）、1（按编码表设置进行打散）
- delSrcObj 是否删除原地物 0（不删除）、1（删除）
- callbackFunctionName 地物执行打散后，自动执行的回调函数名称，如果不需回调处理，该参数可以为空值

**说明**：

执行该函数，将打散选择集内的所有对象。

**示例**：

```vbs
下面示例打散编码等于0并且点数少于2个点的对象。
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "0"
SSProcess.SetSelectCondition "SSObj_PointCount", "<", "2"
SSProcess.SelectFilter
SSProcess.ExplodeSelectionObj 1,0, ""
End Sub
```


#### ExportDataToStream 数据输出到二进制流

**函数**：`ExportDataToStream`

**功能**：数据输出到二进制流。

**语法**：

```vbs
SSProcess.ExportDataToStream(byteObjList)
```

**参数**：

- byteObjList 二进制对象列表

**说明**：

执行该函数，数据输出到二进制流。


#### ExportDataToStream4House 房产库数据输出到二进制流

**函数**：`ExportDataToStream4House`

**功能**：房产数据输出到二进制流。

**语法**：

```vbs
SSProcess.ExportDataToStream4House(dataType，fromSelection,byteObjList)
```

**参数**：

- dateType 数据类型
- fromSelection 是否从选择集中输出（1：选择集中输出，0：全部输出）
- byteObjList 数据对象列表

**说明**：

此函数可以实现从二进制流导入房产库数据。


#### FilterSelectionObjVertex 选择集内对象节点抽稀

**函数**：`FilterSelectionObjVertex`

**功能**：选择集内对象节点抽稀。

**语法**：

```vbs
SSProcess.FilterSelectionObjVertex(miniDistV, miniDistH ,saveHeightMark, saveSurveyPoint, saveTurnPoint, saveFramePoint)
```

**参数**：

- miniDistV 垂距,米为单位
- delSrcObj 弦距,米为单位
- saveHeightMark 保留高程注记点 0(不保留) 1(保留)
- saveSurveyPoint 保留实测点 0(不保留) 1(保留)
- saveTurnPoint 保留转向点 0(不保留) 1(保留)
- saveFramePoint 保留特征点 0(不保留) 1(保留)

**说明**：

执行该函数，将对选择集内的所有对象进行节点抽稀处理。

**示例**：

```vbs
下面示例对编码等于4410对象进行抽稀处理：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "4410"
SSProcess.SelectFilter
SSProcess.FilterSelectionObjVertex 0.01,0.05,1,1,1,1
End Sub
```


#### FindFeatureCode 根据 SQL 条件查找编码

**函数**：`FindFeatureCode`

**功能**：根据SQL条件查找编码

**语法**：

```vbs
SSProcess.FindFeatureCode(featureTBName,matchConditions)
```

**参数**：

- featureTBName 模板feature表名称
- matchConditions 查找条件

**说明**：

执行此函数将根据SQL条件查找对象编码。

**示例**：

```vbs
Sub Onclick()
SSProcess.ClearSysSelection
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "Line"
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount
For i= 0 To geoCount - 1
layername = SSProcess.GetSelgeoValue (i, "SSObj_LayerName" )
byname = SSProcess.GetSelgeoValue (i, "[GB]" )
findConditionss="FeatureCodeTB_500.ByName='"&byname&"'And FeatureCodeTB_500.LayerName='" & layername &"'"
code=SSProcess.FindFeatureCode ("FeatureCodeTB_500", findConditionss)
Msgbox code
Next
End Sub
```


#### FindFeatureCodeByname 根据 Byname 查找编码

**函数**：`FindFeatureCodeByname`

**功能**：根据Bymane查找编码。

**语法**：

```vbs
SSProcess.FindFeatureCodeByname(byname)
```

**参数**：

- byname 别名

**说明**：

使用此函数获取对应 byname的编码。
返回值即为指定 byname所对应的编码。

**示例**：

```vbs
Sub OnClick()
A = SSProcess.FindFeatureCodeByname (129)
Msgbox A
End Sub
```


#### FindNoteClass 根据 SQL 条件查找注记分类号

**函数**：`FindNoteClass`

**功能**：根据SQL条件查找注记分类号。

**语法**：

```vbs
SSProcess.FindNoteClass(notetemplateTBName,matchConditions)
```

**参数**：

- notetemplateTBName 注记模板名称
- matchConditions 查找条件

**说明**：

执行此函数将根据SQL条件查找注记分类号。

**示例**：

```vbs
Sub Onclick()
SSProcess.ClearSysSelection
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
noteCount = SSProcess.GetSelNoteCount
For i= 0 To noteCount - 1
layername1 = SSProcess.GetSelNoteValue (i, "SSObj_LayerName" )
byname1 = "4190000"
'设置要查找的注记的byname
objtype1 = SSProcess.GetSelNoteValue (i, "SSObj_Type")
findConditionss="NoteTemplateTB_ArcInfo.ByName='"&byname1&"'And NoteTemplateTB_ArcInfo.LayerName='" & layername1 &"'"
FontClass=SSProcess.FindNoteClass("NoteTemplateTB_ArcInfo", findConditionss)
Msgbox FontClass
Next
End Sub
```


#### FindNoteClassByname 根据 Byname 查找注记分类号

**函数**：`FindNoteClassByname`

**功能**：根据Byname查找注记分类号。

**语法**：

```vbs
SSProcess.FindNoteClassByname(byname)
```

**参数**：

- byname 别名

**说明**：

执行该函数，将 根据Byname查找注记分类号。

**示例**：

```vbs
Sub Onclick()
SSProcess.ClearSysSelection
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
noteCount = SSProcess.GetSelNoteCount
For i= 0 To noteCount - 1
layername1 = SSProcess.GetSelNoteValue (i, "SSObj_LayerName" )
byname1 = "4190000"
'设置要查找的注记的byname
objtype1 = SSProcess.GetSelNoteValue (i, "SSObj_Type")
findConditionss="NoteTemplateTB_ArcInfo.ByName='"&byname1&"'And NoteTemplateTB_ArcInfo.LayerName='" & layername1 &"'"
FontClass =SSProcess.FindNoteClassByname (byname1)
Msgbox FontClass
Next
End Sub
```


#### GetDistDir 求点到折线集的距离和左右方向

**函数**：`GetDistDir`

**功能**：求点到折线集的距离和左右方向。

**语法**：

```vbs
SSProcess.GetDistDir(geoID,X,Y,Irec,Flag)
```

**参数**：

- geoID 折线集的ID
- X 点到X坐标
- Y 点的Y坐标
- Irec 到直线最近点的点序列号
- Flag 点在直线的方位，0为右，1为左

**说明**：

执行该函数，获取点到直线最近点的距离和方向，同时返回最近点的序列号。

**示例**：

```vbs
Sub OnClick()
pointID =8 '设定点地物ID
lineID =7  '设定直线的ID
'获取指定 ID点地物的坐标
SSProcess.GetObjectPoint pointID , 0, x ,y, z, ptype, name
'获取点到直线最近点的距离
Dist= SSProcess.GetDistDir (lineID , x, y , irec, flag )
Msgbox "点到直线最近点的距离为："&Dist, 64,"消息"
End Sub
```


#### GetNearDist 查找指定地物最近边索引

**函数**：`GetNearDist`

**功能**：查找指定地物最近边索引。

**语法**：

```vbs
SSProcess.GetNearDist(geoID,x,y,index)
```

**参数**：

- geoID 被查找地物ID
- x 查找位置（点）的X坐标
- y 查找位置（点）的Y坐标
- index 返回值，索引边号

**说明**：

执行该函数，将获取从给定地物离参考点最近的边索引。

**示例**：

```vbs
Sub OnClick()
PointID =2
SSProcess.GetObjectPoint PointID, 0, x, y, z, ptype, name
geoID =3
SSProcess.GetNearDist geoID, x, y, index
MsgBox index
End Sub
```


#### GetNearPointIndex 查找指定地物最近点索引

**函数**：`GetNearPointIndex`

**功能**：查找指定地物最近点索引。

**语法**：

```vbs
SSProcess.GetNearPointIndex( geoID, x, y)
```

**参数**：

- geoID 地物ID
- x 查找参考点X坐标
- y 查找参考点Y坐标

**说明**：

执行该函数，将获取从给定地物离参考点最近的点索引。
返回值为最近点索引值,从0开始

**示例**：

```vbs
Sub OnClick()
PointID =2
SSProcess.GetObjectPoint PointID, 0, x, y, z, ptype, name
geoID =3
pnt=SSProcess.GetNearPointIndex（geoID, x, y）
End Sub
```


#### GetNewObjHandle 获取新创建对象句柄

**函数**：`GetNewObjHandle`

**功能**：获取新创建对象句柄。

**语法**：

```vbs
SSProcess.GetNewObjHandle ()
```

**说明**：

使用此函数将获取新创建对象句柄

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateNewObj 0
SSProcess.SetNewObjValue "SSObj_Code", "3805011"
SSProcess.SetNewObjValue "SSObj_LayerName", "居民地点"
SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
SSProcess.SetNewObjValue "SSObj_LineType", "0"
SSProcess.AddNewObjPoint 168.930,902.449, 23.45, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
p = SSProcess.GetNewObjHandle()
Msgbox "NewObjHandle=" & p
End Sub
```


#### GetObjectAttr 获取对象属性

**函数**：`GetObjectAttr`

**功能**：获取对象属性。

**语法**：

```vbs
SSProcess.GetObjectAttr( geoID, attrField)
```

**参数**：

- geoID 地物ID
- attrField 要获取的属性字段名称,见下表
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 0-256 颜色号、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_Angle | 角度 , 以度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_Area | 面积，以平方米为单位 |
| SSObj_EllipsoidArea1 | 椭球面积 ( 内插点计算 ) ，以平方米为单位 |
| SSObj_EllipsoidArea2 | 椭球面积 ( 不内插点计算 ) ，以平方米为单位 |
| SSObj_Length | 2D 长度，以米为单位 |
| SSObj_3DLength | 3D 长度，以米为单位 |
| SSObj_PointCount | 空间点数 |
| SSObj_X | 空间点首点 X 坐标 |
| SSObj_Y | 空间点首点 Y 坐标 |
| SSObj_Z | 空间点首点 Z 坐标 |
| SSObj_PointName | 空间点首点点名 |
| SSObj_PointType | 空间点首点点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_X(index) | 指定索引空间点 X 坐标 |
| SSObj_Y(index) | 指定索引空间点 Y 坐标 |
| SSObj_Z(index) | 指定索引空间点 Z 坐标 |
| SSObj_PointName(index) | 指定索引空间点点名 |
| SSObj_PointType(index) | 指定索引空间点 点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_HatchStyle | 面填充符号样式 0 符号填充 1 不填充 2 颜色填充 |
| [ExtAttr] | 扩展属性字段 |

**说明**：

执行该函数，将获取指定地物的指定字段的值。
返回值即为指定字段所对应的属性值

**示例**：

```vbs
下面示例把选择集对象输出到外部文件：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
If geocount > 0 Then
Dim fso, tf, i, j
Dim chline, code, layer, color
Set fso = CreateObject("Scripting.FileSystemObject")
Set tf = fso.CreateTextFile("c:.txt", True)
For i=0 To geocount-1
'写基本信息
id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
code = SSProcess.GetObjectAttr id, "SSObj_Code")
layer = SSProcess.GetObjectAttr id, "SSObj_LayerName")
color= SSProcess.GetObjectAttr id, "SSObj_Color")
chline = code & "," & layer & "," & color
tf.WriteLine (chline)
Next
tf.Close
End If
End Sub
```


#### GetObjectFocusPoint 获取多边形的焦点

**函数**：`GetObjectFocusPoint`

**功能**：获取多边形的焦点。

**语法**：

```vbs
SSProcess.GetObjectFocusPoint( geoID, x, y)
```

**参数**：

- geoID 地物ID
- x 返回的多边形焦点x坐标
- y 返回的多边形焦点y坐标

**说明**：

执行该函数，将获取给定多边形的焦点坐标。

**示例**：

```vbs
下面是获取多边形焦点坐标的示例：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "AREA"
SSProcess.SetSelectCondition "SSObj_Code", "==", "2"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
If geocount > 0 Then
For i=0 To geocount-1
geoid = SSProcess.GetSelGeoValue(i, "SSObj_ID")
SSProcess.GetObjectFocusPoint geoid, x, y
Next
End If
End Sub
```


#### GetObjectFocusPointInFrame 获取多边形在图幅内的焦点

**函数**：`GetObjectFocusPointInFrame`

**功能**：获取多边形在图幅内的焦点。

**语法**：

```vbs
SSProcess.GetObjectFocusPointInFrame(geoID,frmCenterX,frmCenterY,x,y)
```

**参数**：

- GeoID 地物ID
- frmCenterX 中心坐标X
- frmCenterY 中心坐标Y
- x 返回的多边形焦点X坐标
- y 返回的多边形焦点Y坐标

**说明**：

执行该函数，将获取给定多边形在图幅内的焦点坐标。


#### GetObjectPoint 获取指定地物的空间点信息

**函数**：`GetObjectPoint`

**功能**：获取指定地物空间点信息。

**语法**：

```vbs
SSProcess.GetObjectPoint(id,pointindex,x,y,z,ptype,name)
```

**参数**：

- id 实体对象ID
- pointindex 实体对象节点的索引号
- x 节点的x坐标（返回值）
- y 节点的y坐标（返回值）
- z 节点的z坐标（返回值）
- ptype 节点的特征值（返回值）
- name 节点的点名（返回值）。

**说明**：

执行该函数，可获取指定地物的空间点信息。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount-1
pointcount = SSProcess.GetSelGeoPointCount(i)
objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
For j=0 To pointcount-1
Dim x, y, z, pointtype, name
SSProcess.GetObjectPoint objID, j, x, y, z, pointtype, name
x = x + 1000
y = y + 1000
SSProcess.SetObjectPoint objID, j, x, y, z, pointtype, name,1
Next
Next
End Sub
```


#### GetPointHeight 获取指定位置点高程值

**函数**：`GetPointHeight`

**功能**：获取指定位置点高程。

**语法**：

```vbs
SSProcess.GetPointHeight(x,y,fromMode)
```

**参数**：

- x 位置点屏幕坐标x
- y 位置点屏幕坐标y
- fromMode 高程取值方式，0为周边高程数据取值法；1为三角网取值法；2为格网取值法

**说明**：

执行该函数，将获取指定位置点高程。

**示例**：

```vbs
Sub OnClick()
SSProcess.GetObjectPoint 3, 0, x, y, z, ptype, name
z = SSProcess.GetPointHeight (x , y, 0)
Msgbox z
End Sub
```


#### GetSelGeoCount 获取当前选择集地物个数

**函数**：`GetSelGeoCount`

**功能**：获取当前选择集地物个数。

**语法**：

```vbs
SSProcess.GetSelGeoCount()
```

**说明**：

执行该函数，将获取执行SelectFilter函数后当前选择集对象中地物对象个数。

**示例**：

```vbs
下面示例当前选择集对象中地物对象个数：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
End Sub
```


#### GetSelGeoHandle 获取选择集图形句柄

**函数**：`GetSelGeoHandle`

**功能**：获取选择集图形句柄。

**语法**：

```vbs
SSProcess.GetSelGeoHandle (index)
```

**参数**：

- index 选择集地物的索引号，从0开始

**说明**：

使用此函数获取选择集图形句柄。
返回值即为指定字段所对应的值。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark '设置Undo 标志，支持后面操作的回退
SSProcess.ClearSelection '清空选择集
SSProcess.ClearSelectCondition '清空选择条件
SSProcess.ClearSelectConditionGroups '清空设定的所有分组选择条件
SSProcess.SelectFilter '完成选择过滤（执行选择过滤）
GeoCount = SSProcess.GetSelGeoCount() '得到对象个数
For i= 0 To GeoCount - 1
A = SSProcess.GetSelGeoHandle (i)
Next
End Sub
```


#### GetSelGeoPoint 获取当前选择集指定地物的空间点信息

**函数**：`GetSelGeoPoint`

**功能**：获取当前选择集指定地物的空间点信息。

**语法**：

```vbs
SSProcess.GetSelGeoPoint( index, pointindex, x, y, z, point type, name )
```

**参数**：

- index 选择集地物的索引号,从0开始
- pointindex 地物点索引号,从0开始
- x 空间点X坐标
- y 空间点Y坐标
- z 空间点Z坐标
- pointtype 空间点类型
- name 空间点名

**说明**：

执行该函数，将获取当前选择集对象中指定地物第pointindex+1个空间点的信息。

**示例**：

```vbs
下面示例获取当前选择集指定地物的空间点坐标：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount-1
pointcount = SSProcess.GetSelGeoPointCount(i)
For j=0 To pointcount-1
Dim x, y, z, pointtype, name
SSProcess.GetSelGeoPoint i, j, x, y, z, pointtype, name
Next
Next
End Sub
```


#### GetSelGeoPointCount 获取当前选择集指定地物的点数

**函数**：`GetSelGeoPointCount`

**功能**：获取当前选择集指定地物的点数。

**语法**：

```vbs
SSProcess.GetSelGeoPointCount( index )
```

**参数**：

- index 选择集地物的索引号,从0开始

**说明**：

执行该函数，将获取当前选择集对象中指定地物的空间点个数。

**示例**：

```vbs
下面示例获取当前选择集对象中指定地物的空间点个数：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount-1
pointcount = SSProcess.GetSelGeoPointCount(i)
Next
End Sub
```


#### GetSelGeoValue 获取当前选择集指定地物的属性

**函数**：`GetSelGeoValue`

**功能**：获取当前选择集指定地物的属性。

**语法**：

```vbs
SSProcess.GetSelGeoValue( index, attrField)
```

**参数**：

- index 选择集地物的索引号,从0开始
- attrField 要获取的属性字段名称,见下表
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 0-256 颜色号、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_Angle | 角度 , 以弧度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_Area | 平面面积，以平方米为单位 |
| SSObj_EllipsoidArea1 | 椭球面积 ( 内插点计算 ) ，以平方米为单位 |
| SSObj_EllipsoidArea2 | 椭球面积 ( 不内插点计算 ) ，以平方米为单位 |
| SSObj_Length | 2D 长度，以米为单位 |
| SSObj_3DLength | 3D 长度，以米为单位 |
| SSObj_PointCount | 空间点数 |
| SSObj_X | 空间点首点 X 坐标 |
| SSObj_Y | 空间点首点 Y 坐标 |
| SSObj_Z | 空间点首点 Z 坐标 |
| SSObj_PointName | 空间点首点点名 |
| SSObj_PointType | 空间点首点点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_X(index) | 指定索引空间点 X 坐标 |
| SSObj_Y(index) | 指定索引空间点 Y 坐标 |
| SSObj_Z(index) | 指定索引空间点 Z 坐标 |
| SSObj_PointName(index) | 指定索引空间点点名 |
| SSObj_PointType(index) | 指定索引空间点 点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_HatchStyle | 面填充符号样式 0 符号填充 1 不填充 2 颜色填充 |
| [ExtAttr] | 扩展属性字段 |

**说明**：

执行该函数，将获取执行SelectFilter函数后当前选择集第index+1个地物的指定字段的值。
返回值即为指定字段所对应的属性值

**示例**：

```vbs
下面示例把选择集对象输出到外部文件：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=, "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
If geocount > 0 Then
Dim fso, tf, i, j
Dim chline, code, layer, color, pointcount
Dim x, y, z, pointtype, name
Set fso = CreateObject("Scripting.FileSystemObject")
Set tf = fso.CreateTextFile("c:.txt", True)
For i=0 To geocount-1
'写基本信息
code = SSProcess.GetSelGeoValue(i, "SSObj_Code")
layer = SSProcess.GetSelGeoValue(i, "SSObj_LayerName")
color= SSProcess.GetSelGeoValue(i, "SSObj_Color")
pointcount = SSProcess.GetSelGeoPointCount(i)
chline = code & "," & layer & "," & color & "," & pointcount
tf.WriteLine (chline)
'写点位信息
For j=0 To pointcount-1
SSProcess.GetSelGeoPoint i, j, x, y, z, pointtype, name
chline = x & "," & y & "," & z & "," & pointtype & "," & name
tf.WriteLine (chline)
Next
Next
tf.Close
End If
End Sub
```


#### GetSelNoteCount 获取当前选择集注记个数

**函数**：`GetSelNoteCount`

**功能**：获取当前选择集注记个数。

**语法**：

```vbs
SSProcess.GetSelNoteCount()
```

**说明**：

执行该函数，将获取执行SelectFilter函数后当前选择集对象中注记对象个数。

**示例**：

```vbs
下面示例当前选择集对象中注记对象个数：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount
Msgbox notecount
End Sub
```


#### GetSelNoteHandle 获取选择集注记句柄

**函数**：`GetSelNoteHandle`

**功能**：获取选择集注记句柄。

**语法**：

```vbs
SSProcess.GetSelNoteHandle(index)
```

**参数**：

- index 选择集地物的索引号,从0开始

**说明**：

此函数的功能为获取选择集注记句柄。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==" , "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount
For  i=0 To notecount-1
kkk= SSProcess.GetSelNoteHandle(i)
Msgbox kkk
Next
End Sub
```


#### GetSelNotePoint 获取当前选择集指定注记的空间点信息

**函数**：`GetSelNotePoint`

**功能**：获取当前选择集指定注记的空间点信息。

**语法**：

```vbs
SSProcess.GetSelNotePoint( index, pointindex, x, y, z, point type, name )
```

**参数**：

- index 选择集注记的索引号,从0开始
- pointindex 注记点索引号,从0开始
- x 空间点X坐标
- y 空间点Y坐标
- z 空间点Z坐标
- pointtype 空间点类型
- name 空间点名

**说明**：

执行该函数，将获取当前选择集对象中指定注记第pointindex+1个空间点的信息。

**示例**：

```vbs
下面示例获取当前选择集指定注记的空间点坐标：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount
For i=0 To notecount-1
pointcount = SSProcess.GetSelNotePointCount(i)
For j=0 To pointcount-1
Dim x, y, z, pointtype, name
SSProcess.GetSelNotePoint i, j, x, y, z, pointtype, name
Next
Next
End Sub
```


#### GetSelNotePointCount 获取当前选择集指定注记的点数

**函数**：`GetSelNotePointCount`

**功能**：获取当前选择集指定注记的点数。

**语法**：

```vbs
SSProcess.GetSelNotePointCount( index )
```

**参数**：

- index 选择集注记的索引号,从0开始

**说明**：

执行该函数，将获取当前选择集对象中指定注记的空间点个数。

**示例**：

```vbs
下面示例获取当前选择集对象中指定注记的空间点个数：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount
For i=0 To notecount-1
pointcount = SSProcess.GetSelNotePointCount(i)
Next
End Sub
```


#### GetSelNoteValue 获取当前选择集指定注记的属性

**函数**：`SetSelNoteValue`

**功能**：设置当前选择集指定注记的属性。

**语法**：

```vbs
SSProcess.SetSelNoteValue( index, attrField, attrValue)
```

**参数**：

- index 选择集注记的索引号,从0开始
- attrField 要设置的属性字段名称,见下表：
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 0-256 颜色号、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_Angle | 角度 , 以弧度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_FontWidth | 字宽，以图上 0.01 毫米为单位 |
| SSObj_FontHeight | 字高，以图上 0.01 毫米为单位 |
| SSObj_FontName | 字体名 |
| SSObj_FontClass | 注记分类号 |
| SSObj_FontAlignment | 对齐方式，包括 0 （中心）、 1 （左上角）、 2 （左下角）、 3 （左中）、 4 （右上角）、 5 （右中）、 6 （右下角）、 7 （上中）、 8 （下中） |
| SSObj_FontDirection | 字头朝向，包括 0( 正东 ) 、 1( 正北 ) 、 2( 正西 ) 、 3( 正南 ) 、 4( 切线方向 ) 、 5( 法线方向 ) |
| SSObj_FontWordAngle | 字角度 , 以度为单位 |
| SSObj_FontStringAngle | 串角度 , 以度为单位 |
| SSObj_FontWeight | 字重 |
| SSObj_FontIlaticAngle | 倾斜，包括 0 （不倾斜）、 -1 （左斜）、 1 （右斜） |
| SSObj_FontDownAngle | 耸肩，包括 0 （不耸肩）、 -1 （左耸肩）、 1 （右耸肩） |
| SSObj_FontUnderLine | 下划线，包括 0 （无下划线）、 1 （有下划线） |
| SSObj_FontInterval | 字隔，以图上 0.01 毫米为单位 |
| SSObj_FontString | 注记内容 |
| [ExtAttr] | 扩展属性字段 |
- attrValue 要设置的属性值

**说明**：

执行该函数，将修改当前选择集第index+1个注记的指定字段的值。
函数只是修改对象在内存中的值,并不真正修改数据库,如果要同时改库,则需与AddSelNoteToSaveNoteList和SaveBufferObjToDatabase函数配合使用,只有执行以上两个函数后,才会真正的修改数据数据。
下面示例修改当前选择集对象：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount()
If notecount > 0 Then
'锁定数据库
SSProcess.SetMapStatus 1, 2
For i=0 To notecount-1
'修改基本信息
SSProcess.SetSelNoteValue i, "SSObj_FontClass", "0"
SSProcess.SetSelNoteValue i, "SSObj_LayerName","DEFAULT"
'加入存库对象列表
SSProcess.AddSelNoteToSaveNoteList i
Next
'开锁数据库
SSProcess.SetMapStatus 0, 2
'存库对象列表保存到数据库
SSProcess.SaveBufferObjToDatabase
End If
End Sub


#### GetSelObjHandle 获取选择集内容句柄

**函数**：`GetSelObjHandle`

**功能**：获取选择集内容句柄。

**语法**：

```vbs
SSProcess.GetSelObjHandle()
```

**说明**：

执行该函数，将获取选择集的内容句柄。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
haddle = SSProcess.GetSelObjHandle
End Sub
```


#### GetSelObjList 提取选择集内容

**函数**：`GetSelObjList`

**功能**：提取选择集内容。

**语法**：

```vbs
SSProcess.GetSelObjList(Handle)
```

**参数**：

- Handle 数据源句柄

**说明**：

执行该函数，将提取选择集内容。


#### GetSelPolygonFocus 获取当前选择集指定地物的面心点坐标

**函数**：`GetSelPolygonFocus`

**功能**：获取当前选择集指定地物的面心点坐标。

**语法**：

```vbs
SSProcess.GetSelPolygonFocus(index,x,y,z)
```

**参数**：

- index 地物索引号，从0开始
- x X坐标
- y Y坐标
- z Z坐标

**说明**：

执行该函数，将获取指定地物的面心坐标。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "8104003"
SSProcess.SetSelectCondition "<Close>", "==", "1"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
For i=0 To geocount- 1
SSProcess.GetSelPolygonFocus i , x, y , z
Next
End Sub
```


#### GetSeqIDList 获取对象序列 ID 列表

**函数**：`GetSeqIDList`

**功能**：获取对象序列ID列表。

**语法**：

```vbs
SSProcess.GetSeqIDList(geoID, seqIDList, seqTableName)
```

**参数**：

- geoID 地物ID
- seqIDList 序列ID列表
- seqTableName 属性表名

**说明**：

执行该函数，将获取对象序列ID列表。


#### GetSeqRecord 获取对象序列属性

**函数**：`GetSeqRecord`

**功能**：获取对象序列属性。

**参数**：

- geoID 对象ID
- seqID 序列ID
- fieldNames 属性字段名称
- fieldValues 属性字段值
- seqTableName 属性表名

**说明**：

执行该函数，将获取对象序列属性。

**示例**：

```vbs
Sub OnClick()
Dim fieldName(10),fieldNamecount
Dim fieldValue(10),fieldValuecount
geoID =2
seqID =0
seqTableName ="房屋权利人扩展属性表"
'获得属性记录数目
seqcount=SSProcess.GetSeqRecordCount( geoID,seqTableName)
For i_seq =0 To seqcount -1
fieldNames =""
fieldValues =""
SSProcess.GetSeqRecord geoID, i_seq, fieldNames, fieldValues, seqTableName
SSFunc.ScanString fieldNames,"|",fieldName,fieldNamecount
SSFunc.ScanString fieldValues,"|",fieldValue,fieldValuecount
QLR=getvalue("权利人",fieldName,fieldValue,fieldValuecount)
JZMJ=getvalue("建筑面积",fieldName,fieldValue,fieldValuecount)
fieldN ="分摊系数"
fieldV=12 +i_seq
SSProcess.SetSeqRecord geoID, i_seq, fieldN, fieldV, seqTableName
Msgbox i_seq& chr(13) & fieldNames & chr(13) & fieldValues & chr(13) & QLR & "，" & JZMJ
Next
End Sub
'获得指定属性字段值
Function getvalue(field,fields,values,vcnt)
For i=0 To vcnt-1
If fields(i)=field Then
getvalue= values(i)
Exit Function
End If
Next
End Function
```


#### GetSeqRecordCount 获取对象序列属性个数

**函数**：`GetSeqRecordCount`

**功能**：获取对象序列属性个数。

**语法**：

```vbs
SSProcess.GetSeqRecordCount(geoID,seqTableName)
```

**参数**：

- geoID 对象ID
- seqTableName 属性表名

**说明**：

执行该函数，将获取对象序列属性个数。

**示例**：

```vbs
Sub OnClick()
geoID =2
seqTableName ="RESA属性表"
'获得属性记录数目
seqcount=SSProcess.GetSeqRecordCount( geoID,seqTableName)
Msgbox seqcount
End Sub
```


#### ImportDataFromStream 从二进制流导入数据

**函数**：`ImportDataFromStream`

**功能**：从二进制流导入数据。

**语法**：

```vbs
SSProcess.ImportDataFromStream(byteObjList)
```

**参数**：

- byteObjList 数据对象列表

**说明**：

执行该函数，将从二进制流导入数据。


#### ImportDataFromStream4House 从二进制流导入房产数据

**函数**：`ImportDataFromStream4House`

**功能**：从二进制流导入房产数据。

**说明**：

执行该函数，将从二进制流导入房产数据。


#### InsertSeqRecord 插入对象序列属性

**函数**：`InsertSeqRecord`

**功能**：插入对象序列属性。

**语法**：

```vbs
SSProcess.InsertSeqRecord(geoID,seqID,fieldNames,fieldValues,seqTableName)
```

**参数**：

- geoID 地物ID
- seqID 序列表ID
- fieldNames 文件名
- fieldValues 文件值
- seqTableName 序列表名

**说明**：

执行该函数，将插入对象序列属性。

**示例**：

```vbs
下面示例获取当前选择集对象中指定地物的空间点个数：
Sub OnClick()
Dim fieldName(10),fieldNamecount
Dim fieldValue(10),fieldValuecount
geoID =1
seqID =0
seqTableName ="房屋权利人扩展属性表"
'获得属性记录数目
seqcount=SSProcess.GetSeqRecordCount( geoID ,seqTableName)
For i_seq =0 To seqcount - 1
fieldNames =""
fieldValues =""
SSProcess.GetSeqRecord geoID , i_seq, fieldNames, fieldValues , seqTableName
SSFuc.ScanString fieldNames ,"|" ,fieldName,fieldNamecount
SSFuc.ScanString fieldValues ,"|" ,fieldValue,fieldValuecount
QLR =getvalue( "权利人" ,fieldName,fieldValue,fieldValuecount )
JZMJ =getvalue( "建筑面积" ,fieldName,fieldValue,fieldValuecount )
fieldN ="分摊系数"
fieldV =12 + i_seq
SSProcess.InsertSeqRecord geoID , i_seq, fieldN, fieldV , seqTableName
Msgbox i_seq & chr( 13) & fieldNames & chr (13) & fieldValues & chr(13) & QLR & "," & JZMJ
Next
End Sub
```


#### IsClockwise 是否为顺时针

**函数**：`IsClockwise`

**功能**：是否为顺时针。

**语法**：

```vbs
SSProcess.IsClockwise(geoID)
```

**参数**：

- geoID 地物ID

**说明**：

执行该函数，判断该地物是否为顺时针。
返回 0 逆时针 1顺时针

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "3103003"
SSProcess.SelectFilter
geocount= SSProcess.GetSelGeoCount()
If geocount> 0 Then
For i=0 To geocount - 1
geoID = SSProcess.GetSelGeoValue(i ,"SSObj_ID" )
IsClockwise = SSProcess.IsClockwise(geoID ) '判断地物是否为顺时针
Msgbox "IsClockwise="&IsClockwise
Next
End If
End Sub
```


#### IsExistentFeatureCode 判断编码是否存在

**函数**：`IsExistentFeatureCode`

**功能**：判断编码是否存在。

**语法**：

```vbs
SSProcess.IsExistentFeatureCode( code )
```

**参数**：

- code 要判断的编码

**说明**：

执行该函数可判断指定编码在编码表定义中是否存在
返回 0 不存在 1 存在

**示例**：

```vbs
下面示例判断编码2110是否存在：
Sub OnClick()
isexist = SSProcess.IsExistentFeatureCode("2110")
If isexist = 0 Then
Msgbox "编码2110不存在"
End If
End Sub
```


#### IsExistentFontClass 判断分类号是否存在

**函数**：`IsExistentFontClass`

**功能**：判断注记分类号是否存在。

**语法**：

```vbs
SSProcess.IsExistentFontClass( code )
```

**参数**：

- code 要判断的注记分类号

**说明**：

执行该函数可判断指定分类号在注记分类表定义中是否存在
返回 0 不存在 1 存在

**示例**：

```vbs
下面示例判断分类号1是否存在：
Sub OnClick()
isexist = SSProcess.IsExistentFontClass("1")
If isexist = 0 Then
Msgbox "分类号1不存在"
End If
End Sub
```


#### LineCrack 线分解

**函数**：`LineCrack`

**功能**：线分解。

**语法**：

```vbs
SSProcess.LineCrack(geoID,longcrackMode)
```

**参数**：

- geoID 地物ID
- longcrackMode 分解模式 0：分解折线；1：分解多义线

**说明**：

执行该函数，可以对指定线地物进行分解。

**示例**：

```vbs
Sub OnClick()
'选择所有数据
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount -1
'获取地物的ID
geoID = SSProcess.GetSelGeoValue( i , "SSObj_ID" )
'对每个地物进行分解
SSProcess.LineCrack geoID ,0
Next
End Sub
```


#### LineParallelDist 距离平行线

**函数**：`LineParallelDist`

**功能**：距离平行线。

**语法**：

```vbs
SSProcess.LineParallelDist(geoID,dist,direction,wholeLine,moveOrCopy)
```

**参数**：

- geoID 地物ID
- dist 距离
- direction 方向  0：左侧，1：右侧
- wholeLine 全线/一段 0：全段，1：一段
- moveOrCopy 平移/复制 0：平移，1：复制

**说明**：

根据ID=1的线地物生成平行线。

**示例**：

```vbs
Sub OnClick()
geoID=1
SSProcess.LineParallelDist geoID ,10,0,0,1
End Sub
```


#### LineParallelScale 比例平行线

**函数**：`LineParallelScale`

**功能**：比例平行线。

**语法**：

```vbs
SSProcess.LineParallelScale(geoID,side,newDist,moveOrCopy)
```

**参数**：

- geoID 地物 ID
- side 新创建或复制地物的位置
- newDist 新创建或者复制线与原地物的距离
- moveOrCopy 复制或者移动

**说明**：

按比例生成平行线

**示例**：

```vbs
Sub OnClick()
geoID=1
SSProcess.LineParallelScale geoID,0,10,0
End Sub
```


#### LockObjectPoint 锁定（解锁）指定地物的点列

**函数**：`LockObjectPoint`

**功能**：锁定(解锁)指定地物的点列。

**语法**：

```vbs
SSProcess.LockObjectPoint(geoID status)
```

**参数**：

- geoID 地物ID
- status 状态 0：解锁 1：锁定

**说明**：

执行该函数，在对数据坐标转换时锁定数据点位信息，在执行完所有操作后再解锁地物点列信息，一次性将数据点列信息进行更新，提高数据处理的速度。

**示例**：

```vbs
Sub OnClick()
'选择所有数据
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount -1
'获取地物的点数
ID= SSProcess.GetSelGeoValue( i , "SSObj_ID" )
pointCount = SSProcess.GetSelGeoPointCount (i )
'锁定地物点列
SSProcess.LockObjectPoint ID, 1
'对于每个点进行数据坐标转换
For j=0 To pointCount -1
SSProcess.GetSelGeoPoint i, j, x , y, z , ptype, pname
SSProcess.LongiLatiToxy54 126, y, x, x0, y0
SSProcess.SetSelGeoPoint i, j, x0 , y0, z , ptype, pname
Next
SSProcess.UpdateSelGeoPoint i
'解锁地物点列
SSProcess.LockObjectPoint ID, 0
Next
End Sub
```


#### LockSelGeoPoint 锁定（解锁）当前选择集指定地物的点列

**函数**：`LockSelGeoPoint`

**功能**：LockSelGeoPoint 函数

**语法**：

```vbs
SSProcess.LockSelGeoPoint( index, status )
```

**参数**：

- index 选择集地物的索引号,从0开始
- status 锁定或解锁状态 0 解锁内存点列 1 锁定内存点列

**说明**：

执行该函数，将解锁或者锁定当前选择地物的点列到内存,以提高对选择对象点列的编辑效率。
锁定和解锁需配对执行,一旦设置了锁定状态,在编辑结束后,必须执行一次解锁.

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount- 1
SSProcess.LockSelGeoPoint i , 1
SSProcess.GetSelGeoPoint i , 0, y , x, z , pointtype, name
xx = x+10:yy=y+10
SSProcess.SetSelGeoPoint i , 0, yy , xx, z , pointtype, name
SSProcess.UpdateSelGeoPoint i
SSProcess.LockSelGeoPoint i , 0
SSProcess.AddSelGeoToSaveGeoList i
Next
SSProcess.SaveBufferObjToDatabase
SSProcess.RefreshView
End Sub
```


#### LockSelNotePoint 锁定（解锁）当前选择集指定注记的点列

**函数**：`LockSelNotePoint`

**功能**：LockSelNotePoint 函数

**语法**：

```vbs
SSProcess.LockSelNotePoint(index, status)
```

**参数**：

- index 选择集注记的索引号,从0开始
- status 锁定或解锁状态 0 解锁内存点列 1 锁定内存点列

**说明**：

执行该函数，将解锁或者锁定当前选择注记的点列到内存,以提高对选择对象点列的编辑效率。
锁定和解锁需配对执行,一旦设置了锁定状态,在编辑结束后,必须执行一次解锁.

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectCondition "SSObj_ PointCount ", "<", "2"
SSProcess.SetSelectCondition "SSObj_Type", "==", "Note"
SSProcess.SelectFilter
notecount = SSProcess.GetSelnoteCount
For i=0 To notecount- 1
SSProcess.LockSelNotePoint i , 1
SSProcess.GetSelNotePoint i , 0, y , x, z , pointtype, name
xx = x+10:yy=y+10
SSProcess.SetSelNotePoint i , 0, yy , xx, z , pointtype, name
SSProcess.UpdateSelNotePoint i
SSProcess.LockSelNotePoint i , 0
SSProcess.AddSelNoteToSaveNoteList i
Next
SSProcess.SaveBufferObjToDatabase
SSProcess.RefreshView
End Sub
```


#### MergeObjByCondition 带条件线面合并

**函数**：`MergeObjByCondition`

**功能**：带条件线面合并。

**语法**：

```vbs
SSProcess.MergeObjByCondition(mergeType,srcObjCodes,maxMergeDist,isMergeByRange,rangeCodes,isMergeGroupByAttr,fieldNames, isMergeOnFrameLine,getObjMode)
```

**参数**：

- mergeType 合并对象类型 1 线 2 面
- srcObjCodes 合并目标编码
- maxMergeDist 合并最大间距
- isMergeByRange 是否只在区域内合并 0 否 1 是
- rangeCodes 区域编码，当isMergeByRange =1时只对rangeCodes编码指定的地物范围内合并
- isMergeGroupByAttr 是否属性值相等时合并 0 否 1 是
- fieldNames 要判断相等的字段名称
- isMergeOnFrameLine 是否在分幅边界上合并 0 否 1 是
- getObjMode 获取处理对象方式 0 脚本选择集 1 全图

**说明**：

使用此函数将根据设定的条件将面或线地物进行合并

**示例**：

```vbs
Sub OnClick()
mergeType = 2 : srcObjCodes = "2" : maxMergeDist = 0.5 : isMergeByRange = 0 : rangeCodes = "" : isMergeGroupByAttr = 1 : fieldNames = "Note" : isMergeOnFrameLine = 0 : getObjMode = 1
SSProcess.MergeObjByCondition mergeType,srcObjCodes, maxMergeDist, isMergeByRange, rangeCodes, isMergeGroupByAttr, fieldNames, isMergeOnFrameLine, getObjMode
End Sub
```


#### MergeObjPointList 合并指定 ID 列表地物

**函数**：`MergeObjPointList`

**功能**：合并指定ID列表地物。

**语法**：

```vbs
SSProcess.MergeObjPointList(objIDs,limit,delLinkObj)
```

**参数**：

- objIDs 要合并的地物ID列表，用逗号分隔
- limit 间隙最小距，米为单位
- delLinkObj 是否删除原地物

**说明**：

执行该函数，将合并指定ID列表的地物为一新地物。

**示例**：

```vbs
'合并线段
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount
SSProcess.LockGeoPointsInMemory 1
geoids = ""
For i=0 To geoCount- 1
geoid = SSProcess.GetSelGeoValue (i, "SSObj_ID" )
If geoids = "" Then
geoids = geoid
Else
geoids = geoids&","& geoid
End If
Next
SSProcess.MergeObjPointList geoids , 0.001, 1
SSProcess.LockGeoPointsInMemory 0
SSProcess.RefreshView
End Sub
```


#### OffsetSelectionObj 平移选择集对象

**函数**：`OffsetSelectionObj`

**功能**：平移选择集内的所有对象。

**语法**：

```vbs
SSProcess.OffsetSelectionObj( x, y, z)
```

**参数**：

- x X坐标平移量
- y Y坐标平移量
- z Z坐标平移量

**说明**：

执行该函数，将平行移动选择内的所有对象，平移量均为数学坐标系，X为东方向，Y为北方向。

**示例**：

```vbs
下面示例平移编码等于0的对象：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
SSProcess.OffsetSelectionObj 10,20,1.2
End Sub
```


#### PushUndoMark 设置 Undo 标志

**函数**：`PushUndoMark`

**功能**：设置Undo标志

**语法**：

```vbs
SSProcess.PushUndoMark()
```

**说明**：

执行该函数，将为往后的编辑操作设置Undo标志，支持后面操作的回退。


#### RemoveSelectionObjPoint 删除选择集对象的空间点

**函数**：`RemoveSelectionObjPoint`

**功能**：删除选择集对象的空间点。

**语法**：

```vbs
SSProcess.RemoveSelectionObjPoint( index, count)
```

**参数**：

- index 空间点索引号，从0开始
- count 删除点数

**说明**：

执行该函数，将删除选择集对象中的指定空间点。

**示例**：

```vbs
下面示例把删除选择集对象的第二个空间点：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
SSProcess.RemoveSelectionObjPoint 1,1
End Sub
```


#### RemoveSelGeo 移除当前选择集指定地物

**函数**：`RemoveSelGeo`

**功能**：移除当前选择集指定地物。

**语法**：

```vbs
SSProcess.RemoveSelGeo(index)
```

**参数**：

- index 地物索引号，从0开始

**说明**：

执行该函数，将移除当前选择集指定索引序号的地物。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSysSelection
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "3103013"
SSProcess.SelectFilter
SSProcess.RemoveSelGeo 0
End Sub
```


#### RemoveSelNote 移除当前选择集指定注记

**函数**：`RemoveSelNote`

**功能**：移除当前选择集指定索引序号的注记。

**语法**：

```vbs
SSProcess.RemoveSelNote(index)
```

**参数**：

- index 注记索引号，从0开始

**说明**：

执行该函数，将移除当前选择集指定注记。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
SSProcess.RemoveSelNote 0
End Sub
```


#### ResetSelGeoByCode 根据编码重置当前选择集指定地物的属性

**函数**：`ResetSelGeoByCode`

**功能**：根据编码重置当前选择集指定地物的属性。

**语法**：

```vbs
SSProcess.ResetSelGeoByCode(index,code)
```

**参数**：

- index 地物索引号，从0开始，脚本选择集地物索引号
- code 重置后的地物编码

**说明**：

使用指定的编码（重置后的地物编码）重置地物要素的编码、图层、颜色等基本属性，多用于数据导入及编码匹配后的地物重置。

**示例**：

```vbs
'道路颜色重置
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT,LINE,AREA "
SSProcess.SetSelectCondition "SSObj_Color","=" ,"RGB(0,255,0)"
SSProcess.SelectFilter
geocount= SSProcess.GetSelGeoCount()
For i =0 To geocount- 1
code = SSProcess.GetSelGeoValue(i ,"SSObj_Code" )
SSProcess.ResetSelGeoByCode i , code
SSProcess.AddSelGeoToSaveGeoList i
Next
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### ResetSelNoteByFontClass 根据注记分类号重置当前选择集指定注记的属性

**函数**：`ResetSelNoteByFontClass`

**功能**：根据注记分类号重置当前选择集指定注记的属性。

**语法**：

```vbs
SSProcess.ResetSelNoteByFontClass(index,fontClass)
```

**参数**：

- index 注记序号
- fontClass 注记分类号

**说明**：

执行该函数，根据注记分类号重置当前选择集指定注记的属性。

**示例**：

```vbs
'道路颜色重置
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectCondition "SSObj_Type", "=", "NOTE"
SSProcess.SetSelectCondition "SSObj_Color","=" ,"RGB(0,255,0)"
SSProcess.SelectFilter
notecount= SSProcess.GetSelNoteCount()
For i =0 To notecount- 1
fontClass = SSProcess.GetSelNoteValue(i ,"SSObj_FontClass" )
SSProcess.ResetSelNoteByFontClass i ,fontClass
SSProcess.AddSelNoteToSaveNoteList i
Next
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### SaveBufferObjToDatabase 保存存库列表对象到数据库

**函数**：`SaveBufferObjToDatabase`

**功能**：保存存库列表对象到数据库。

**语法**：

```vbs
SSProcess.SaveBufferObjToDatabase()
```

**说明**：

执行该函数，将把当前存库列表中的地物和注记对象保存到数据库中。

**示例**：

```vbs
下面示例保存存库列表对象到数据库：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateNewObjByClass "1"
SSProcess.SetNewObjValue "SSObj_FontString", "上地路"
SSProcess.AddNewObjPoint 1000.01, 1000.34, 22.45, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### SelectionObjClip 选择对象多边形裁剪

**函数**：`SelectionObjClip`

**功能**：选择对象多边形裁剪。

**语法**：

```vbs
SSProcess.SelectionObjClip(clipPolygonID,clipFlag,extendDist)
```

**参数**：

- clipPolygonID 选择的多边形ID ，裁剪依据的多边形
- clipFlag 裁剪方式0 表示裁内， 1 表示裁外
- extendDist 延伸距离，裁剪的时候根据范围线往外扩的距离（这个范围线就是选中的多边型）

**说明**：

使用此函数裁剪与选中多边形相关联的地物。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "LINE"
SSProcess.SetSelectCondition "SSObj_Code", "=", "1"
SSProcess.SelectFilter
clipPolygonID = 56
clipFlag = 0
extendDist = 1
SSProcess.SelectionObjClip clipPolygonID ,clipFlag, extendDist
End Sub
```


#### SelectionObjClip1 选择对象多边形裁剪

**函数**：`SelectionObjClip1`

**功能**：选择对象多边形裁剪。

**语法**：

```vbs
SSProcess.SelectionObjClip1(clipPolygonXYCoords,clipFlag)
```

**参数**：

- clipPolygonXYCoords 所要选择裁剪对象的范围面
- clipFlag （0：表示裁范围面内要素，1：表示裁范围面外要素）

**说明**：

使用此函数可以对范围面内的对象进行裁剪。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark '设置Undo标志，支持回退
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups '清空设定的所有分组选择条件
SSProcess.SetSelectCondition "SSObj_Code", "=", "1"
SSProcess.SelectFilter
yBL = "431500"
xBL = "2886200"
yBR = "431550"
xBR = "2886200"
yTR = "431550"
xTR = "2886250"
yTL = "431500"
xTL = "2886250" '定义坐标对，构成裁剪面
clipPolygonXYCoords=xBL+","+ yBL+","+ xBR+","+ yBR+","+ xTR+","+ yTR+","+ xTL+","+ yTL+","+ xBL+","+ yBL '裁剪范围
clipFlag = 0
SSProcess.SelectionObjClip1 clipPolygonXYCoords ,clipFlag
End Sub
```


#### SelectionObjPerpendLize 选择集对象直角化处理

**函数**：`SelectionObjPerpendLize`

**功能**：选择集对象直角化处理。

**语法**：

```vbs
SSProcess.SelectionObjPerpendLize(minAngle,maxAngle,maxDistance)
```

**参数**：

- minAngle 直角化最小角度
- maxAngle 直角化最大角度
- maxDistance 直角化角度最大距离

**说明**：

执行该函数，将将选择集对象进行直角化处理

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "=", "1"
SSProcess.SelectFilter
SSProcess.SelectionObjPerpendLize 60,80,10
End Sub
```


#### SelectionObjSplitToSingleArea 选择集对象分割成单环面

**函数**：`SelectionObjSplitToSingleArea`

**功能**：选择集对象分割成单环面。

**语法**：

```vbs
SSProcess.SelectionObjSplitToSingleArea()
```

**说明**：

执行该函数，将选择集对象中的岛面分割成单环面。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
SSProcess.SelectFilter
SSProcess.SelectionObjSplitToSingleArea
End Sub
```


#### SelectionObjToClipBoard 选择集对象拷贝到粘贴板

**函数**：`SelectionObjToClipBoard`

**功能**：SelectionObjToClipBoard 函数

**语法**：

```vbs
SSProcess.SelectionObjToClipBoard()
```

**说明**：

执行该函数，将选择集对象添加到粘帖板，通过与SSProcess.AddClipBoardObjToMap配合，将选择集的对象添加到图形中。

**示例**：

```vbs
Sub OnClick()
'选择地物
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
SSProcess.SelectFilter
'将选择集对象拷贝到粘贴板
SSProcess.SelectionObjToClipBoard
'粘帖板对象添加到地图
SSProcess.AddClipBoardObjToMap 100, 100
End Sub
```


#### SelGeoGotoPoints 当前选择集对象折线化

**函数**：`SelGeoGotoPoints`

**功能**：当前选择集对象折线化。

**语法**：

```vbs
SSProcess.SelGeoGotoPoints(index,mode)
```

**参数**：

- index 地物索引号
- mode 是否折线化 0 不折线化；1 折线化

**说明**：

执行该函数，将选择集对象进行折线化

**示例**：

```vbs
Sub OnClick()
'选择地物
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "=", "1"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
For i= 0 To  geocount -1
'将对象进行折线化
SSProcess.SelGeoGotoPoints i,1
Next
End Sub
```


#### SelNoteGotoPoints 当前选择集注记折线化

**函数**：`SelNoteGotoPoints`

**功能**：当前选择集对象折线化。

**语法**：

```vbs
SSProcess.SelNoteGotoPoints(index,mode)
```

**参数**：

- index 注记索引号,从0开始
- mode 折线化形式

**示例**：

```vbs
Sub OnClick()
'选择地物
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "NOTE"
SSProcess.SelectFilter
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount()
For i= 0 To  notecount -1
'将对象进行折线化
SSProcess.SelNoteGotoPoints i,1
Next
End Sub
```


#### SetChangeSelectionObjParameter 设置选择集修改参数

**函数**：`SetChangeSelectionObjParameter`

**功能**：设置选择集修改参数。

**语法**：

```vbs
SSProcess.SetChangeSelectionObjParameter(matchConditions,updateSetAttrs)
```

**参数**：

- matchConditions 选择集修改条件，参数格式"属性字段1=值1,属性字段2=值2,..."，属性字段请参看SetSelectCondition函数说明
- updateSetAttrs 修改选择内容，参数格式"属性字段1=值1,属性字段2=值2,..."，属性字段请参看ChangeSelectionObjAttr函数说明

**说明**：

该函数执行一次，设定一条选择集修改记录，函数允许连续多次执行，以设定多个修改记录。
执行该函数，不进行实际的修改处理，只有执行ChangeSelection函数后，才会真正开始修改选择集内容。
ChangeSelectionObjAttr函数一次只能修改一个属性字段，而ChangeSelection函数则可以一次执行多个属性字段的修改，执行效率要高于ChangeSelectionObjAttr函数。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
SSProcess.ClearChangeSelectionParameter
SSProcess.SetChangeSelectionObjParameter "[结构]=11,[层数]=12", "SSObj_Color=COLORNO(3)"
SSProcess.SetChangeSelectionObjParameter "[结构]=10,[层数]=11", "SSObj_Color=COLORNO(2)"
SSProcess.SetChangeSelectionObjParameter "[结构]=8,[层数]=9", "SSObj_Color=COLORNO(5)"
SSProcess.ChangeSelectionObj
End Sub
```


#### SetNewObjValue 设置新图形对象属性

**函数**：`SetNewObjValue`

**功能**：设置新图形对象属性。

**语法**：

```vbs
SSProcess.SetNewObjValue( attrField, attrValue)
```

**参数**：

- attrField 要设置的属性字段名称,见下表
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 0-256 颜色号、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_Angle | 角度 , 以弧度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_FontWidth | 字宽，以图上 0.01 毫米为单位 |
| SSObj_FontHeight | 字高，以图上 0.01 毫米为单位 |
| SSObj_FontName | 字体名 |
| SSObj_FontClass | 注记分类号 |
| SSObj_FontAlignment | 对齐方式，包括 0 （中心）、 1 （左上角）、 2 （左下角）、 3 （左中）、 4 （右上角）、 5 （右中）、 6 （右下角）、 7 （上中）、 8 （下中） |
| SSObj_FontDirection | 字头朝向，包括 0( 正东 ) 、 1( 正北 ) 、 2( 正西 ) 、 3( 正南 ) 、 4( 切线方向 ) 、 5( 法线方向 ) |
| SSObj_FontWordAngle | 字角度 , 以度为单位 |
| SSObj_FontStringAngle | 串角度 , 以度为单位 |
| SSObj_FontWeight | 字重 |
| SSObj_FontIlaticAngle | 倾斜，包括 0 （不倾斜）、 -1 （左斜）、 1 （右斜） |
| SSObj_FontDownAngle | 耸肩，包括 0 （不耸肩）、 -1 （左耸肩）、 1 （右耸肩） |
| SSObj_FontUnderLine | 下划线，包括 0 （无下划线）、 1 （有下划线） |
| SSObj_FontInterval | 字隔，以图上 0.01 毫米为单位 |
| SSObj_FontString | 注记内容 |
| [ExtAttr] | 扩展属性字段 |
- attrValue 要设置的属性值

**说明**：

执行该函数，将设置当前新建对象指定字段的值。
函数只是设置对象在内存中的值,并不真正修改数据库,如果要同时改库,则需与AddNewObjToSaveObjList和SaveBufferObjToDatabase函数配合使用,只有执行以上两个函数后,才会真正的修改数据数据。

**示例**：

```vbs
下面示例设置当前新建对象的属性值：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateNewObjByClass "1"
SSProcess.SetNewObjValue "SSObj_FontString", "上地路"
SSProcess.AddNewObjPoint 1000.01, 1000.34, 22.45, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.CreateNewObjByClass "2"
SSProcess.SetNewObjValue "SSObj_FontString", "清河路"
SSProcess.AddNewObjPoint 1020.01, 1030.34, 22.45, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### SetObjectAttr 修改对象属性值

**函数**：`SetObjectAttr`

**功能**：修改对象属性值。

**语法**：

```vbs
SSProcess.SetObjectAttr( geoID, attrField, attrValue)
```

**参数**：

- geoID 地物ID
- attrField 要获取的属性字段名称,见下表
- attrValue 要修改字段的新属性值
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 0-256 颜色号、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_Angle | 角度 , 以弧度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_PointCount | 空间点数 |
| SSObj_X | 空间点首点 X 坐标 |
| SSObj_Y | 空间点首点 Y 坐标 |
| SSObj_Z | 空间点首点 Z 坐标 |
| SSObj_PointName | 空间点首点点名 |
| SSObj_PointType | 空间点首点点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_X(index) | 指定索引空间点 X 坐标 |
| SSObj_Y(index) | 指定索引空间点 Y 坐标 |
| SSObj_Z(index) | 指定索引空间点 Z 坐标 |
| SSObj_PointName(index) | 指定索引空间点点名 |
| SSObj_PointType(index) | 指定索引空间点 点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_HatchStyle | 面填充符号样式 0 符号填充 1 不填充 2 颜色填充 |
| [ExtAttr] | 扩展属性字段 |

**说明**：

执行该函数，将修改指定地物的指定字段值。

**示例**：

```vbs
下面示例修改选择集对象的编码：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
If geocount > 0 Then
For i=0 To geocount-1
id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
SSProcess.SetObjectAttr id, "SSObj_Code", "0"
Next
End If
End Sub
```


#### SetObjectPoint 设置指定地物的空间点信息

**函数**：`SetObjectPoint`

**功能**：设置指定地物的空间点信息。

**语法**：

```vbs
SSProcess.SetObjectPoint(id,pointindex,x,,y,,z,,ptype,,name,savetodatabase)
```

**参数**：

- id 实体对象ID
- pointindex 实体对象节点的索引号
- x 节点的x坐标（返回值）
- y 节点的y坐标（返回值）
- z 节点的z坐标（返回值）
- ptype 节点的特征值（返回值）
- name 节点的点名（返回值）。
- savetodatabase 是否保存到数据库 0：不保存；1：保存。

**说明**：

执行该函数，可设置指定地物的空间点信息。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount-1
pointcount = SSProcess.GetSelGeoPointCount(i)
objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
For j=0 To pointcount-1
Dim x, y, z, pointtype, name
SSProcess.GetObjectPoint objID, j, x, y, z, pointtype, name
x = x + 1000
y = y + 1000
SSProcess.SetObjectPoint objID, j, x, y, z, pointtype, name,1
Next
Next
End Sub
```


#### SetSelGeoPoint 设置当前选择集指定地物的空间点信息

**函数**：`SetSelGeoPoint`

**功能**：设置当前选择集指定地物的空间点信息。

**语法**：

```vbs
SSProcess.SetSelGeoPoint( index, pointindex, x, y, z, point type, name )
```

**参数**：

- index 选择集地物的索引号,从0开始
- pointindex 地物点索引号,从0开始
- x 空间点X坐标
- y 空间点Y坐标
- z 空间点Z坐标
- pointtype 空间点类型
- name 空间点名

**说明**：

执行该函数，将设置当前选择集对象中指定地物第pointindex+1个空间点的信息。
如果pointindex大于已有点数，则为新加点。

**示例**：

```vbs
下面示例设置当前选择集指定地物的空间点坐标：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount-1
pointcount = SSProcess.GetSelGeoPointCount(i)
For j=0 To pointcount-1
Dim x, y, z, pointtype, name
SSProcess.GetSelGeoPoint i, j, x, y, z, pointtype, name
x = x + 1000
y = y + 1000
SSProcess.SetSelGeoPoint  i, j, x, y, z, pointtype, name
Next
SSProcess.AddSelGeoToSaveGeoList i
Next
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### SetSelGeoValue 设置当前选择集指定地物的属性

**函数**：`SetSelGeoValue`

**功能**：设置当前选择集指定地物的属性。

**语法**：

```vbs
SSProcess.SetSelGeoValue( index, attrField, attrValue)
```

**参数**：

- index 选择集地物的索引号,从0开始
- attrField 要设置的属性字段名称,见下表
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 0-256 颜色号、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_Angle | 角度 , 以弧度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_HatchStyle | 面填充符号样式 0 符号填充 1 不填充 2 颜色填充 |
| [ExtAttr] | 扩展属性字段 |
- attrValue 要设置的属性值

**说明**：

执行该函数，将修改当前选择集第index+1个地物的指定字段的值。
函数只是修改对象在内存中的值,并不真正修改数据库,如果要同时改库,则需与AddSelGeoToSaveGeoList和SaveBufferObjToDatabase函数配合使用,只有执行以上两个函数后,才会真正的修改数据数据。

**示例**：

```vbs
下面示例修改当前选择集对象：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount()
If geocount > 0 Then
'锁定数据库
SSProcess.SetMapStatus 1, 2
Dim x, y, z, pointtype, name
For i=0 To geocount-1
'修改基本信息
SSProcess.SetSelGeoValue i, "SSObj_Code", "0"
SSProcess.SetSelGeoValue i, "SSObj_LayerName","DEFAULT"
'加入存库对象列表
SSProcess.AddSelGeoToSaveGeoList i
Next
'开锁数据库
SSProcess.SetMapStatus 0, 2
'存库对象列表保存到数据库
SSProcess.SaveBufferObjToDatabase
End If
End Sub
```


#### SetSelNotePoint 设置当前选择集指定注记的空间点信息

**函数**：`SetSelNotePoint`

**功能**：设置当前选择集指定注记的空间点信息。

**语法**：

```vbs
SSProcess.SetSelNotePoint( index, pointindex, x, y, z, point type, name )
```

**参数**：

- index 选择集注记的索引号,从0开始
- pointindex 注记点索引号,从0开始
- x 空间点X坐标
- y 空间点Y坐标
- z 空间点Z坐标
- pointtype 空间点类型
- name 空间点名

**说明**：

执行该函数，将设置当前选择集对象中指定注记第pointindex+1个空间点的信息。
如果pointindex大于已有点数，则为新加点。

**示例**：

```vbs
下面示例设置当前选择集指定注记的空间点坐标：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount
For i=0 To notecount-1
pointcount = SSProcess.GetSelNotePointCount(i)
For j=0 To pointcount-1
SSProcess.GetSelNotePoint i, j, x, y, z, pointtype, name
x = x+1000
y = y+1000
SSProcess.SetSelNotePoint  i, j, x, y, z, pointtype, name
Next
SSProcess.AddSelGeoToSaveGeoList i
Next
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### SetSelNoteValue 设置当前选择集指定注记的属性

**函数**：`SetSelNoteValue`

**功能**：设置当前选择集指定注记的属性。

**语法**：

```vbs
SSProcess.SetSelNoteValue( index, attrField, attrValue)
```

**参数**：

- index 选择集注记的索引号,从0开始
- attrField 要设置的属性字段名称,见下表：
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 0-256 颜色号、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_Angle | 角度 , 以弧度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_FontWidth | 字宽，以图上 0.01 毫米为单位 |
| SSObj_FontHeight | 字高，以图上 0.01 毫米为单位 |
| SSObj_FontName | 字体名 |
| SSObj_FontClass | 注记分类号 |
| SSObj_FontAlignment | 对齐方式，包括 0 （中心）、 1 （左上角）、 2 （左下角）、 3 （左中）、 4 （右上角）、 5 （右中）、 6 （右下角）、 7 （上中）、 8 （下中） |
| SSObj_FontDirection | 字头朝向，包括 0( 正东 ) 、 1( 正北 ) 、 2( 正西 ) 、 3( 正南 ) 、 4( 切线方向 ) 、 5( 法线方向 ) |
| SSObj_FontWordAngle | 字角度 , 以度为单位 |
| SSObj_FontStringAngle | 串角度 , 以度为单位 |
| SSObj_FontWeight | 字重 |
| SSObj_FontIlaticAngle | 倾斜，包括 0 （不倾斜）、 -1 （左斜）、 1 （右斜） |
| SSObj_FontDownAngle | 耸肩，包括 0 （不耸肩）、 -1 （左耸肩）、 1 （右耸肩） |
| SSObj_FontUnderLine | 下划线，包括 0 （无下划线）、 1 （有下划线） |
| SSObj_FontInterval | 字隔，以图上 0.01 毫米为单位 |
| SSObj_FontString | 注记内容 |
| [ExtAttr] | 扩展属性字段 |
- attrValue 要设置的属性值

**说明**：

执行该函数，将修改当前选择集第index+1个注记的指定字段的值。
函数只是修改对象在内存中的值,并不真正修改数据库,如果要同时改库,则需与AddSelNoteToSaveNoteList和SaveBufferObjToDatabase函数配合使用,只有执行以上两个函数后,才会真正的修改数据数据。
下面示例修改当前选择集对象：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount()
If notecount > 0 Then
'锁定数据库
SSProcess.SetMapStatus 1, 2
For i=0 To notecount-1
'修改基本信息
SSProcess.SetSelNoteValue i, "SSObj_FontClass", "0"
SSProcess.SetSelNoteValue i, "SSObj_LayerName","DEFAULT"
'加入存库对象列表
SSProcess.AddSelNoteToSaveNoteList i
Next
'开锁数据库
SSProcess.SetMapStatus 0, 2
'存库对象列表保存到数据库
SSProcess.SaveBufferObjToDatabase
End If
End Sub


#### SetSelObjHandle 设置选择集内容句柄

**函数**：`SetSelObjHandle`

**功能**：设置选择集内容句柄。

**语法**：

```vbs
SSProcess.SetSelObjHandle(Handle)
```

**参数**：

- Handle 句柄

**说明**：

执行该函数，将设置选择集内容句柄。


#### SetSeqRecord 设置对象序列属性

**函数**：`SetSeqRecor`

**功能**：SetSeqRecord函数

**语法**：

```vbs
SSProcess.SetSeqRecord(geoID,seqID,fieldNames,fieldValues,seqTableName)
```

**参数**：

- geoID 地物ID
- seqID 信息表ID
- fieldNames 字段名
- filedValues 字段值
- seqTableName 字段所属表名

**说明**：

执行此函数设置对象序列属性。

**示例**：

```vbs
Sub OnClick()
Dim fieldName(10),fieldNamecount
Dim fieldValue(10),fieldValuecount
geoID =2
seqID =0
seqTableName ="房屋权利人扩展属性表"
'获得属性记录数目
seqcount=SSProcess.GetSeqRecordCount( geoID,seqTableName)
For i_seq =0 To seqcount -1
fieldNames =""
fieldValues =""
SSProcess.GetSeqRecord geoID, i_seq, fieldNames, fieldValues, seqTableName
SSFunc.ScanString fieldNames,"|",fieldName,fieldNamecount
ScanString fieldValues,"|",fieldValue,fieldValuecount
QLR=getvalue("权利人",fieldName,fieldValue,fieldValuecount)
JZMJ=getvalue("建筑面积",fieldName,fieldValue,fieldValuecount)
fieldN ="分摊系数"
fieldV=12 +i_seq
SSProcess.SetSeqRecord geoID, i_seq, fieldN, fieldV, seqTableName
Msgbox i_seq& chr(13) & fieldNames & chr(13) & fieldValues & chr(13) & QLR & "," & JZMJ
Next
End Sub
'获得指定属性字段值
Function getvalue(field,fields,values,vcnt)
For i=0 To vcnt-1
If fields(i)=field Then
getvalue= values(i)
Exit Function
End If
Next
End Function
```


#### TransArcNotePointNote 弧度注记转单点注记

**函数**：`TransArcNotePointNote`

**功能**：弧段注记转单点注记。

**语法**：

```vbs
SSProcess.TransArcNotePointNote(gnoteID)
```

**参数**：

- gnoteID 注记ID

**说明**：

执行该函数，将弧段注记转单点注记。

**示例**：

```vbs
Sub OnClick()
gnoteID=126 '弧段注记ID
SSProcess.TransArcNotePointNote gnoteID
End Sub
```


#### TransAttrToObj 属性转图形

**函数**：`TransAttrToObj`

**功能**：属性转图形。

**语法**：

```vbs
SSProcess.TransAttrToObj(geoID, fieldname, insertPointX, insertPointY)
```

**参数**：

- geoID 存储图形的ID号
- fieldName 存储图形的字段名称（二进制文件）
- insertPointX 插入点X坐标
- insertPointY 插入点Y坐标

**说明**：

使用此函数目的是将存储在指定字段中的属性释放转换成图形。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "=", "9999613"
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount
If geoCount=0 Then
Msgbox "数据中‘图廓范围线-9999613’不存在，已取消操作！" :Exit Sub
elseIf geoCount>1 Then
Msgbox "数据中‘图廓范围线-9999613’有多个，已取消操作！" :Exit Sub
End If
geoID=SSProcess.GetSelGeoValue( 0, "SSObj_ID" )
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT,LINE,AREA,NOTE"
SSProcess.SetSelectCondition "SSObj_Code", "<>", "9999613"
SSProcess.SelectFilter
SSProcess.TransSelectionObjToAttr geoID , "附注", 1 '图形转属性
SSProcess.TransAttrToObj 2, "附注", 0, 0 '属性转图形
End Sub
```


#### TransPointNoteToArcNote 单点注记转弧度注记

**函数**：`TransPointNoteToArcNote`

**功能**：单点注记转弧段注记。

**语法**：

```vbs
SSProcess.TransPointNoteToArcNote(noteID)
```

**参数**：

- noteID 注记ID

**说明**：

将指定ID的点注记转为弧段注记

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "Note"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount()
For i=0 To notecount- 1
'写基本信息
gnoteID = SSProcess.GetSelNoteValue(i , "SSObj_ID") ' 给定要转换的点注记的 ID
zjnr =SSProcess.GetObjectAttr (gnoteID,"SSObj_FontString" )
If  zjnr<>replace(zjnr ," " ,"") THEN
wbzjnr =replace( zjnr," ", "" )
If len (wbzjnr)> 1 Then
'将点注记转换为线注记
SSProcess.TransPointNoteToArcNote gnoteID
End If
End If
Next
End Sub
```


#### UpdateNoteAttrByNoteTemplate 根据注记模板表修改注记属性

**函数**：`UpdateNoteAttrByNoteTemplate`

**功能**：根据NoteTemplate注记模板表修改注记属性。

**语法**：

```vbs
SSProcess.UpdateNoteAttrByNoteTemplate(notetemplateTBName, matchConditions, updateSetAttrs)
```

**参数**：

- notetemplateTBName NoteTemplateTB注记模板表名称
- matchConditions 匹配条件参数, 格式为标准SQL语法，"NoteTemplate属性 = 注记属性 and NoteTemplate属性 = 注记属性 and ...",如果 NoteTemplate字段名的值是字符型, 则需要把注记属性值用''号括起来 如 NoteTemplate.Byname='SSObj_Byname'updateSetAttrs 重置注记属性参数, 格式为 "注记属性=NoteTemplate属性,注记属性=NoteTemplate属性,..."
- NoteTemplate属性字段见下表：
| 参数 | 说明 |
|---|---|
| NoteTemplate.FontClass | 分类号 |
| NoteTemplate.FontColor | 文字颜色 |
| NoteTemplate.StringAngle | 字串角度 |
| NoteTemplate.FontName | 字体名 |
| NoteTemplate.FontWidth | 字宽 |
| NoteTemplate.FontHeight | 字高 |
| NoteTemplate.FontInterval | 字隔 |
| NoteTemplate.FontWeight | 字重 |
| NoteTemplate.IlaticAngle | 斜体 |
| NoteTemplate.WordAngle | 字角度 |
| NoteTemplate.DownAngle | 耸肩 |
| NoteTemplate.Underline | 下划线 |
| NoteTemplate.Memo | 备注 |
| NoteTemplate.LayerName | 层名 |
| NoteTemplate.FontByname | 字体别名 |
| NoteTemplate.Filter | 是否过滤 |
| NoteTemplate.Byname | 别名 |
| NoteTemplate.HS | 对齐方式 |
- 注记属性字段见下表：
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 0-256 颜色号、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称，为 FeatureCodeTB 表中的 ObjectName |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_FontWidth | 字宽，以图上 0.01 毫米为单位 |
| SSObj_FontHeight | 字高，以图上 0.01 毫米为单位 |
| SSObj_FontName | 字体名 |
| SSObj_FontClass | 注记分类号 |
| SSObj_FontAlignment | 对齐方式，包括 0 （中心）、 1 （左上角）、 2 （左下角）、 3 （左中）、 4 （右上角）、 5 （右中）、 6 （右下角）、 7 （上中）、 8 （下中） |
| SSObj_FontWordAngle | 字角度 , 以度为单位 |
| SSObj_FontStringAngle | 串角度 , 以度为单位 |
| SSObj_FontWeight | 字重 |
| SSObj_FontIlaticAngle | 倾斜，包括 0 （不倾斜）、 -1 （左斜）、 1 （右斜） |
| SSObj_FontDownAngle | 耸肩，包括 0 （不耸肩）、 -1 （左耸肩）、 1 （右耸肩） |
| SSObj_FontUnderLine | 下划线，包括 0 （无下划线）、 1 （有下划线） |
| SSObj_FontInterval | 字隔，以图上 0.01 毫米为单位 |
| SSObj_FontString | 注记内容 |

**说明**：

执行该函数可实现选择集中满足matchConditions指定条件的对象的属性重置为NoteTemplateTB表的相应属性（如分类号、层名、颜色等）。

**示例**：

```vbs
下面示例把图中所有注记对象选入选择集，并实现层名、字高、字宽与NoteTemplateTB表相匹配的注记的分类号置成与NoteTemplateTB表的分类号。
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
SSProcess.UpdateNoteAttrByNoteTemplate "NoteTemplateTB_500", "NoteTemplate.LayerName= 'SSObj_ LayerName' and NoteTemplate.FontHeight = SSObj_FontHeight and NoteTemplate. FontWidth = SSObj_FontWidth", "SSObj_FontClass = NoteTemplate.FontClass"
End Sub
```


#### UpdateObjAttrByFeatureCode 根据编码表修改地物属性

**函数**：`UpdateObjAttrByFeatureCode`

**功能**：根据编码表修改选择集地物的属性。

**语法**：

```vbs
SSProcess.UpdateObjAttrByFeatureCode( featureTBName, matchConditions, updateSetAttrs)
```

**参数**：

- featureTBName 对照编码表FeatureCodeTB表名称
- matchConditions 匹配条件参数, 格式为标准SQL语法，“ FeatureCode属性 = 地物属性 and FeatureCode 属性=地物属性 and ...”,如果 Feature字段名的值是字符型, 则需要把地物属性值用''号括起来 如 Feature.Byname='SSObj_Byname'
- updateSetAttrs 重置地物属性参数, 格式为 "地物属性=FeatureCode属性,地物属性= FeatureCode属性,..."
- FeatureCode 属性字段见下表
| 参数 | 说明 |
|---|---|
| Feature.Code | EPS 编码 |
| Feature.Byname | 转换编码 |
| Feature.LayerName | 层名 |
| Feature.ObjectName | 对象名称 |
| Feature.Type | 几何类型 |
| Feature.LineType | 连接类型 |
| Feature.LineWidth | 线宽 |
| Feature.LineColor | 颜色 |
| Feature.Filter | 是否过滤 |
| Feature.Explode | 是否打散 |
| Feature.Reverse | 是否反向 |
| Feature.Thickness | 厚度 |
| Feature.Others | 其他 |
| Feature.AttrType | 属性类型 |
| Feature.Reference | 参考 |
| Feature.ExtraInfo | 附加信息 |
- 地物属性请参考ChangeSelectionObj函数

**说明**：

执行该函数可实现选择集中满足matchConditions指定条件的对象的属性重置为FeatureCodTB表的相应属性（如编码、层名、颜色等）。

**示例**：

```vbs
下面示例把图中所有点对象选入选择集，并实现扩展属性中[CODE]字段值等于Feature.Byname的对象层名和编码重置为FeatureCodeTB表中对应的编码和层名。
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
SSProcess.SelectFilter
SSProcess.UpdateObjAttrByFeatureCode "FeatureCodeTB_500", "Feature.Byname='[CODE]'", "SSObj_LayerName = Feature.LayerName,SSObj_Code = Feature.Code"
End Sub
```


#### UpdateObjectPoint 更新当前选择集指定地物的点列

**函数**：`UpdateObjectPoint`

**功能**：更新当前选择集指定地物的点列。

**语法**：

```vbs
SSProcess.UpdateObjectPoint(geoID,savetodatabase)
```

**参数**：

- geoID 地物ID
- savetodatabase 是否保存点列到数据库 , 0 只生成到内存 ，1 保存到数据库。

**说明**：

该函数为图形编辑操作函数，用于更新当前选择集指定地物的点列。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">" , "1"
SSProcess.SetSelectCondition "SSObj_Type", "==" , "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount -1
pointcount = SSProcess.GetSelGeoPointCount(i)
For j=0 To pointcount- 1
Dim x, y, z , pointtype, name
SSProcess.GetSelGeoPoint i , j, x , y, z , pointtype, name
x = x + 1000
y = y + 1000
SSProcess.SetSelGeoPoint i , j, x , y, z , pointtype, name
Next
SSProcess.UpdateObjectPoint i,1
Next
End Sub
```


#### UpdateSelGeoPoint 更新当前选择集地物的点列

**函数**：`UpdateObjectPoint`

**功能**：更新当前选择集指定地物的点列。

**语法**：

```vbs
SSProcess.UpdateObjectPoint(geoID,savetodatabase)
```

**参数**：

- geoID 地物ID
- savetodatabase 是否保存点列到数据库 , 0 只生成到内存 ，1 保存到数据库。

**说明**：

该函数为图形编辑操作函数，用于更新当前选择集指定地物的点列。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">" , "1"
SSProcess.SetSelectCondition "SSObj_Type", "==" , "POINT"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount -1
pointcount = SSProcess.GetSelGeoPointCount(i)
For j=0 To pointcount- 1
Dim x, y, z , pointtype, name
SSProcess.GetSelGeoPoint i , j, x , y, z , pointtype, name
x = x + 1000
y = y + 1000
SSProcess.SetSelGeoPoint i , j, x , y, z , pointtype, name
Next
SSProcess.UpdateObjectPoint i,1
Next
End Sub
```


#### UpdateSelNotePoint 更新当前选择集指定注记的点列

**函数**：`UpdateSelNotePoint`

**功能**：更新当前选择集指定注记的点列。

**语法**：

```vbs
SSProcess.UpdateSelNotePoint(index)
```

**参数**：

- index 选择集注记的索引号 ,从0开始

**说明**：

该函数为图形编辑操作函数，用于更新当前选择集指定注记的点列。

**示例**：

```vbs
将当前工程数据（2000坐标系）输出为一个80坐标系的EDB。
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_PointCount", ">", "1"
SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
SSProcess.SelectFilter
noteCount = SSProcess.GetSelNoteCount
For i=0 To noteCount-1
pointCount = SSProcess.GetSelNotePointCount(i)
For j = 0 To pointCount - 1
SSProcess.GetSelNotePoint i, j, x, y, z, ptype, name
Px = x+10:Py = y+10
SSProcess.SetSelNotePoint i, j, Px, Py, z, ptype, name
Next
SSProcess.UpdateSelNotePoint i
Next
End Sub
```


#### ZoomInoutPolygon 放缩多边形

**函数**：`ZoomInoutPolygon`

**功能**：放缩多边形。

**语法**：

```vbs
SSProcess.ZoomInoutPolygon(geoID,dist,copy,code)
```

**参数**：

- geoID 地物ID
- dist 缩放距离
- copy 是否复制（0不复制，1复制）
- code 地物编码

**说明**：

实现多边形放缩

**示例**：

```vbs
Sub OnClick()
SSProcess.ZoomInoutPolygon 290, 20, 1, "2"
End Sub
```


---

### 地模处理

#### CreateDem 创建格网

**函数**：`CreateDem`

**功能**：创建格网。

**语法**：

```vbs
SSProcess.CreateDem(triNetName, demName, ltX, ltY, cellSize, rows, cols)
```

**参数**：

- triNetName 创建格网的三角网名称, 可以是CreateTriNet创建的三角网或系统的三角网
- demName 创建格网名称
- ltX 格网左上X坐标
- ltY 格网左上Y坐标
- cellSize 格距
- rows 格网行数
- cols 格网列表

**说明**：

执行该函数, 将根据triNetName指定的三角网,自动生成以demName命名的格网。

**示例**：

```vbs
下面示例三角网和格网生成过程：
Sub OnClick()
'生成三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "高程点", "SSObj_Type", "=, "POINT"
SSProcess.SetSelectConditionGroup "等高线", "SSObj_Code", "=, "8310"
SSProcess.CreateTriNet "等高线", "三角网", "", 0, 200, 0.01, 200, 1, 0
'生成格网
SSProcess.CreateDem "三角网", "格网", 414745.000, 2517025.000, 5, 100, 100
'输出格网
SSProcess.ExportDem "格网", 2, 1, "c:\"
'删除格网
SSProcess.DeleteDem "格网"
'删除三角网
SSProcess.DeleteTriNet "三角网"
End Sub
```


#### CreateDgx 生成等高线

**函数**：`CreateDgx`

**功能**：生成等高线。

**语法**：

```vbs
SSProcess.CreateDgx(triNetName, dengGaoJu, jqxCode, sqxCode, minDist, maxDist, saveToDatabase)
```

**参数**：

- triNetName 三角网名称
- dengGaoJu 等高距
- jqxCode 计曲线编码
- sqxCode 首曲线编码
- minDist 等高线点间最小距离
- maxDist 等高线点间最大距离
- saveToDatabase 是否保存生成图形到数据库, 0 只生成到内存 1 保存到数据库

**说明**：

执行该函数, 将根据triNetName三角网生成等高线。

**示例**：

```vbs
下面示例生成等高线过程：
Sub OnClick()
SSProcess.PushUndoMark
'首先由现有等高线生成三角网,并生成三角网的外接多边形
'删除三角网外接多边形
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "LINE"
SSProcess.SetSelectCondition "SSObj_Code", "=", "1"
SSProcess.SetSelectCondition "SSObj_LayerName", "=", "三角网外接多边形"
SSProcess.SelectFilter
SSProcess.DeleteSelectionObj
'生成三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "等高线", "SSObj_Type", "=", "LINE"
SSProcess.SetSelectConditionGroup "等高线", "SSObj_LayerName", "=", "等高线"
SSProcess.CreateTriNet "等高线", "等高线三角网", "", 1, 50, 0.01, 40, 1, 0
'生成三角网的外接多边形
SSProcess.CreateTriNetOutPolygon "等高线三角网", "三角网外接多边形", 1, 1
'删除三角网
SSProcess.DeleteTriNet "等高线三角网"
'生成等高线范围内的三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "选择多边形", "SSObj_Type", "=", "LINE"
SSProcess.SetSelectConditionGroup "选择多边形", "SSObj_LayerName", "=", "三角网外接多边形"
SSProcess.SetSelectConditionGroup "选择多边形", "<Close>", "=", "1"
SSProcess.SetSelectConditionGroup "等高线内图形", "SSObj_Type", "=", "LINE"
SSProcess.SetSelectConditionGroup "等高线内图形", "SSObj_LayerName", "=", "DGX"
SSProcess.SetSelectConditionGroup "等高线内图形", "{SelectPolygon_GroupName}", "=", "选择多边形"
SSProcess.SetSelectConditionGroup "等高线内图形","{SelectPolygon_InOutMode}", "=", "5"
SSProcess.CreateTriNet "等高线内图形", "等高线三角网", "", 1, 50, 0.01, 40, 1, 0
'生成等高线
SSProcess.CreateDgx "等高线三角网", 1.0, 8120, 8110, 0.1, 800, 0
'删除三角网
SSProcess.DeleteTriNet "等高线三角网"
End Sub
```


#### CreateTriNetOutPolygon 生成三角网外接多边形

**函数**：`CreateTriNetOutPolygon`

**功能**：生成三角网外接多边形。

**语法**：

```vbs
SSProcess.CreateTriNetOutPolygon( triNetName, layerName, geoCode, saveToDatabase )
```

**参数**：

- triNetName 三角网名称
- layerName 放置新生成图形的图层名称
- geoCode 新生成图形编码
- saveToDatabase 是否保存生成图形到数据库, 0 只生成到内存 1 保存到数据库

**说明**：

执行该函数, 将生成triNetName三角网的外接多形到指定编码和图层。

**示例**：

```vbs
下面示例生成三角网外接多边形：
Sub OnClick()
SSProcess.PushUndoMark
'首先由现有等高线生成三角网,并生成三角网的外接多边形
'删除三角网外接多边形
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "LINE"
SSProcess.SetSelectCondition "SSObj_Code", "=", "1"
SSProcess.SetSelectCondition "SSObj_LayerName", "=", "三角网外接多边形"
SSProcess.SelectFilter
SSProcess.DeleteSelectionObj
'生成三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "等高线", "SSObj_Type", "=", "LINE"
SSProcess.SetSelectConditionGroup "等高线", "SSObj_LayerName", "="
SSProcess.CreateTriNet "等高线", "等高线三角网", "", 1, 50, 0.01, 40, 1, 0
'生成三角网的外接多边形
SSProcess.CreateTriNetOutPolygon "等高线三角网", "三角网外接多边形", 1, 1
End Sub
```


#### CreateTriNet 创建三角网

**函数**：`CreateTriNet`

**功能**：创建三角网。

**语法**：

```vbs
SSProcess.CreateTriNet( groupName, triNetName, txxCodes, minZ, maxZ, minSideLength, maxSideLength, filterZ0, insertStep)
```

**参数**：

- groupName 创建三角网数据源选择组名称
- triNetName 创建三角网名称
- txxCodes 特性线编码, 有多个特性线编码时,用逗号分隔
- minZ 最小有效高程值
- maxZ 最大有效高程值
- minSideLength 最小搜索边长
- maxSideLength 最大搜索边长
- filterZ0 是否处理0高程 1 不处理 0 处理
- insertStep 加密数据源弧段的步长值,为0时,不进行加密处理

**说明**：

执行该函数, 将根据groupName选定的数据源,自动生成以triNetName命名的三角网。

**示例**：

```vbs
Sub OnClick()
'生成三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "高程点", "SSObj_Type", "=, "POINT"
SSProcess.SetSelectConditionGroup "等高线", "SSObj_Code", "=, "8310"
SSProcess.CreateTriNet "等高线", "三角网", "", 0, 200, 0.01, 200, 1, 0
'生成格网
SSProcess.CreateDem "三角网", "格网", 414745.000, 2517025.000, 5, 100, 100
'输出格网
SSProcess.ExportDem "格网", 2, 1, "c:\"
'删除格网
SSProcess.DeleteDem "格网"
'删除三角网
SSProcess.DeleteTriNet "三角网"
End Sub
```


#### DeleteDem 删除格网

**函数**：`DeleteDem`

**功能**：删除格网。

**语法**：

```vbs
SSProcess.DeleteDem( demName )
```

**参数**：

- demName 要删除的格网名称

**说明**：

执行该函数, 将删除名称为demName的格网。

**示例**：

```vbs
Sub OnClick()
'生成三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "高程点", "SSObj_Type", "=, "POINT"
SSProcess.SetSelectConditionGroup "等高线", "SSObj_Code", "=, "8310"
SSProcess.CreateTriNet "等高线", "三角网", "", 0, 200, 0.01, 200, 1, 0
'生成格网
SSProcess.CreateDem "三角网", "格网", 414745.000, 2517025.000, 5, 100, 100
'输出格网
SSProcess.ExportDem "格网", 2, 1, "c:\"
'删除格网
SSProcess.DeleteDem "格网"
'删除三角网
SSProcess.DeleteTriNet "三角网"
End Sub
```


#### DeleteTriNet 删除三角网

**函数**：`DeleteTriNet`

**功能**：删除三角网。

**语法**：

```vbs
SSProcess.DeleteTriNet( triNetName )
```

**参数**：

- triNetName 要删除的三角网名称

**说明**：

执行该函数, 将删除名称为triNetName三角网。

**示例**：

```vbs
Sub OnClick()
'生成三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "高程点", "SSObj_Type", "=, "POINT"
SSProcess.SetSelectConditionGroup "等高线", "SSObj_Code", "=, "8310"
SSProcess.CreateTriNet "等高线", "三角网", "", 0, 200, 0.01, 200, 1, 0
'生成格网
SSProcess.CreateDem "三角网", "格网", 414745.000, 2517025.000, 5, 100, 100
'输出格网
SSProcess.ExportDem "格网", 2, 1, "c:\"
'删除格网
SSProcess.DeleteDem "格网"
'删除三角网
SSProcess.DeleteTriNet "三角网"
End Sub
```


#### ExportDem 输出格网

**函数**：`ExportDem`

**功能**：输出格网。

**语法**：

```vbs
SSProcess.ExportDem( demName, fileType, stretch, pathName)
```

**参数**：

- demName 要输出的格网名称
- fileType 格网文件格式 0(DEM) 1(BIL) 2(ASC) 3(GRD) 4(SGD)
- stretch 高程放大倍数
- pathname 输出格网文件的路径

**说明**：

执行该函数, 将输出名称为demName的格网数据到指定格式的文件。

**示例**：

```vbs
Sub OnClick()
'生成三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "高程点", "SSObj_Type", "=, "POINT"
SSProcess.SetSelectConditionGroup "等高线", "SSObj_Code", "=, "8310"
SSProcess.CreateTriNet "等高线", "三角网", "", 0, 200, 0.01, 200, 1, 0
'生成格网
SSProcess.CreateDem "三角网", "格网", 414745.000, 2517025.000, 5, 100, 100
'输出格网
SSProcess.ExportDem "格网", 2, 1, "c:\"
'删除格网
SSProcess.DeleteDem "格网"
'删除三角网
SSProcess.DeleteTriNet "三角网"
End Sub
```


#### TransTriNetToGeo 转换三角网为图形

**函数**：`TransTriNetToGeo`

**功能**：转换三角网为图形。

**语法**：

```vbs
SSProcess.TransTriNetToGeo( triNetName, layerName, geoCode, saveToDatabase )
```

**参数**：

- triNetName 三角网名称
- layerName 放置新生成图形的图层名称
- geoCode 新生成图形编码
- saveToDatabase 是否保存生成图形到数据库, 0 只生成到内存 1 保存到数据库

**说明**：

执行该函数, 将把名称为triNetName的三角网转换为指定编码的图形。

**示例**：

```vbs
下面示例三角网转换为图形的过程：
Sub OnClick()
SSProcess.PushUndoMark
'首先由现有等高线生成三角网,并生成三角网的外接多边形
'删除三角网外接多边形
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "LINE"
SSProcess.SetSelectCondition "SSObj_Code", "=", "1"
SSProcess.SetSelectCondition "SSObj_LayerName", "=", "三角网外接多边形"
SSProcess.SelectFilter
SSProcess.DeleteSelectionObj
'生成三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "等高线", "SSObj_Type", "=", "LINE"
SSProcess.SetSelectConditionGroup "等高线", "SSObj_LayerName", "=", "等高线"
SSProcess.CreateTriNet "等高线", "等高线三角网", "", 1, 50, 0.01, 40, 1, 0
'生成三角网的外接多边形
SSProcess.CreateTriNetOutPolygon "等高线三角网", "三角网外接多边形", 1, 1
'删除三角网
SSProcess.DeleteTriNet "等高线三角网"
'生成等高线范围内的三角网
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "选择多边形", "SSObj_Type", "=", "LINE"
SSProcess.SetSelectConditionGroup "选择多边形", "SSObj_LayerName", "=", "三角网外接多边形"
SSProcess.SetSelectConditionGroup "选择多边形", "<Close>", "=", "1"
SSProcess.SetSelectConditionGroup "等高线内图形", "SSObj_Type", "=", "LINE"
SSProcess.SetSelectConditionGroup "等高线内图形", "SSObj_LayerName", "=", "DGX"
SSProcess.SetSelectConditionGroup "等高线内图形", "{SelectPolygon_GroupName}", "=", "选择多边形"
SSProcess.SetSelectConditionGroup "等高线内图形","{SelectPolygon_InOutMode}", "=", "5"
SSProcess.CreateTriNet "等高线内图形", "等高线三角网", "", 1, 50, 0.01, 40, 1, 0
'三角网转换成编码为1的图形
SSProcess.TransTriNetToGeo "等高线三角网", "DEFAULT", 1, 0
'删除三角网
SSProcess.DeleteTriNet "等高线三角网"
End Sub
```


---

### 工程管理函数

#### CloseDatabase 关闭数据库

**函数**：`CloseDatabase`

**功能**：CloseDatabase 函数

**语法**：

```vbs
SSProcess.CloseDatabase()
```

**说明**：

执行该函数，将关闭当前打开的工程，并把软件启动时打开的工程做为当前工程。
函数执行成功，返回1，失败返回0。


#### CreateDatabase 创建数据库

**函数**：`CreateDatabase`

**功能**：CreateDatabase 函数

**语法**：

```vbs
SSProcess.CreateDatabase( templateName, fileName )
```

**参数**：

- templateName 为新建工程使用的模板文件全路径名
- fileName 为新建MDB数据的全路径文件名

**说明**：

templateName 应为EPS执行目录下的Templates目录下的模板文件，文件名应同时包含路径，如"C:\Program Files\EPS2012\Templates\GB_500.MDT"。
fileName，应包含路径和文件名，如"c:\epsedb\new.edb"，执行该函数，将自动关闭当前打开的工程，把新建的工程做为当前工程。
函数执行成功，返回1，失败返回0。

**示例**：

```vbs
下面示例使用GB_500.MDT模板创建名为new.edb的工程文件：
Sub OnClick()
SSProcess.CreateDatabase "C:\Program Files\EPS2012\Templates\GB_500.MDT", "c:\epsedb\ new.edb"
End Sub
```


#### GetProjectFileName 获取工程文件名

**函数**：`GetProjectFileName`

**功能**：获取工程文件名。

**语法**：

```vbs
SSProcess.GetProjectFileName()
```

**说明**：

获取工程文件名。

**示例**：

```vbs
Sub OnClick()
strProjectName=SSProcess.GetProjectFileName()
Msgbox strProjectName
End Sub
```


#### GetSysPathName 获取系统路径名

**函数**：`GetSysPathName`

**功能**：GetSysPathName 函数

**语法**：

```vbs
SSProcess.GetSysPathName( mode)
```

**参数**：

- mode 路径类别代号，代号不同获取的相关路径信息不同
- mode 参数取范围及返回信息说明如下：
- 0 ： 获取EPS.exe执行程序所在路径
- 1 ： 获取EPS平台安装目录下的“Templates”模板路径
- 2 ： 获取EPS平台安装目录下的“Comm”路径
- 3 ： 获取EPS平台安装目录下的“Script”路径，已经基本不用
- 4 ： 获取EPS平台安装目录下的“Temp”路径
- 5 ： 获取当前工程所在路径
- 6 ： 获取WINDOWS系统路径
- 7 ： 获取EPS平台安装目录下的“DeskTop”下的当前工作台面路径
- 返回完整路径。

**说明**：

执行该函数，获取EPS系统相关路径。

**示例**：

```vbs
Sub OnClick()
Pathname = SSProcess.GetSysPathName (7)
End Sub
```


#### GetTemplateFileName 获取模板文件名

**函数**：`GetTemplateFileName`

**功能**：获取模板文件名。

**语法**：

```vbs
SSProcess.GetTemplateFileName()
```

**说明**：

获取当前工程（EDB）使用的模板名，返回值为一个包含路径的字符串。

**示例**：

```vbs
Sub OnClick()
Msgbox SSProcess.GetTemplateFileName ()
End Sub
```


#### OpenDatabase 打开数据库

**函数**：`OpenDatabase`

**功能**：OpenDatabase 函数

**语法**：

```vbs
SSProcess.OpenDatabase( fileName )
```

**参数**：

- fileName 要打开的EPS MDB数据的全路径文件名。

**说明**：

fileName，应包含路径和文件名，如"c:\epsedb\new.edb"，执行该函数，将自动关闭当前打开的工程，把打开的工程做为当前工程。
函数执行成功，返回1，失败返回0。

**示例**：

```vbs
下面示例使打开名为new.edb的工程文件：
Sub OnClick()
SSProcess.OpenDatabase "c:\epsedb\new.edb"
End Sub
```


#### SelectFileName 选择文件名

**函数**：`SelectFileNam`

**功能**：SelectFileName函数

**语法**：

```vbs
SSProcess.SelectFileName( status, title, flags, filter )
```

**参数**：

- status 文件对话框选择状态，等于0时，为新建文件对话框，等于1时为选择文件对话框。
- title 为文件对话框的缺省选择文件名，如果不指定缺省文件名，允许为空字符串。
- flags 文件选择方式，等于0时，为单选文件，等于1时为多选文件。
- filter 为文件对话框的过滤条件，如"Chart Files (*.xlc)|*.xlc|Worksheet Files (*.xls)|*.xls|Data Files (*.xlc;*.xls)|*.xlc; *.xls|All Files (*.*)|*.*||"。

**说明**：

执行该函数，将弹出文件选择对话框，以选择文件名。
函数执行成功，返回选择的文件名，失败返回空字符串，如果文件选择对话框为多选择状态，则同时返回多个文件名，用“，”。分隔文件名。

**示例**：

```vbs
下面示例选择扩展名为mdb的单个文件名：
Sub OnClick()
Dim fileName
fileName = SSProcess.SelectFileName(1,"",0,"EPS MDB Files(*.mdb)|*.mdb|All Files (*.*)|*.*||")
End Sub
```


#### SelectPathName 选择路径

**函数**：`SelectPathName`

**功能**：SelectPathName 函数

**语法**：

```vbs
SSProcess.SelectPathName()
```

**说明**：

执行该函数，将弹出文件目录选择对话框，以选择目录名。
函数执行成功，返回选择的路径名，失败返回空字符串。

**示例**：

```vbs
Sub OnClick()
Dim pathName
pathName = SSProcess.SelectPathName()
End Sub
```


#### UpdateCurMap 更新当前地图

**函数**：`UpdateCurMap`

**功能**：更新当前地图。

**语法**：

```vbs
SSProcess.UpdateCurMap(mapHandle)
```

**参数**：

- projectHandle 工程句柄

**说明**：

执行该函数，更新当前地图

**示例**：

```vbs
下面示例切换当前工程为另外打开的工程：
Sub OnClick()
prjCount = SSProject.GetProjectCount '获取工程个数
curProjectHandle = SSProject.GetActiveProject '获取当前活动的工程句柄
For i=0 To prjCount- 1
projectHandle = SSProject.GetProject(i) '获取第i 个工程句柄
If projectHandle <> curProjectHandle Then
SSProject.SetActiveProject projectHandle '设置活动工程句柄
SSProcess.UpdateCurMap projectHandle '更新当前地图
End If
Next
End Sub
```


---

### 角度函数

#### AdjustAngle 调整弧度在 0-2PI 间

**函数**：`AdjustAngle`

**功能**：调整弧度在0-2PI间。

**语法**：

```vbs
SSProcess.AdjustAngle(radian)
```

**参数**：

- radian 需要被转换的弧度值

**说明**：

执行完函数后，弧度值被调整在0-2PI之间

**示例**：

```vbs
Sub OnClick()
radian= 4*3.1415926
'执行完后radian 值为2*3.1415926
radian= SSProcess.AdjustAngle(radian)
End Sub
```


#### AdjustDegAngle 调整度在 0-360 间

**函数**：`AdjustDegAngle`

**功能**：调整度在0-360间。

**语法**：

```vbs
SSProcess.AdjustDegAngle(deg)
```

**参数**：

- deg 度

**说明**：

该函数为角度函数，用于将单位为度的角度值调整到0 -360间。

**示例**：

```vbs
下面示例是求两条线的夹角，并将夹角调整到0 -360间：
Sub OnClick()
x0 = -101:y0 = 22:x1 = -101:y1= 22:flag = 0
SSProcess.XYSA x0,y0,x1 ,y1,dist ,angle, flag
Ang0= SSProcess.RadianToDeg (angle )
x2 = -101:y2= 66:flag = 0
SSProcess.XYSA x1,y1,x2 ,y2,dist ,angle, flag
Ang1= SSProcess.RadianToDeg (angle )
jiajiao= Ang1-Ang0
jiajiao= SSProcess.AdjustDegAngle(jiajiao )
Msgbox jiajiao
End Sub
```


#### AdjustDmsAngle 调整度分秒在 0-360 间

**函数**：`AdjustDmsAngle`

**功能**：调整度分秒在0-360间。

**语法**：

```vbs
SSProcess.AdjustDmsAngle(dms)
```

**参数**：

- dms 单位为度分秒的角度值

**说明**：

该函数为角度函数，用于将单位为度分秒的角度值调整到0 -360间。

**示例**：

```vbs
下面示例是求两条线的夹角，并将夹角调整到0 -360间：
Sub OnClick()
x0 = -10144.346656815158:y0 = 2285.589553773533
x1 = -10105.466515801174:y1= 2244.786768423693
flag = 0
SSProcess.XYSA x0,y0,x1 ,y1,dist ,angle, flag
Ang0= SSProcess.RadianToDms (angle )
Msgbox Ang0
x2 = -10050.564338655053:y2= 2282.598773695534
flag = 0
SSProcess.XYSA x1,y1,x2 ,y2,dist ,angle, flag
Ang1= SSProcess.RadianToDms (angle )
Msgbox Ang1
jiajiao= Ang1-Ang0
jiajiao= SSProcess.AdjustDmsAngle(jiajiao )
Msgbox jiajiao
End Sub
```


#### DegToDmm 度转度分分

**函数**：`DegToDmm`

**功能**：度转度分分。

**语法**：

```vbs
SSProcess.DegToDmm(deg)
```

**参数**：

- deg 度

**说明**：

执行该函数，将度转为度分分

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==" , "2203041"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i=0 To geocount -1
jd =SSProcess.GetSelGeoValue(i, "SSObj_Angle")
'弧度转角度
jd =SSProcess.RadianToDeg(jd)
JD =SSProcess.DegToDmm(jd)
Next
End Sub
```


#### DegToDms 度转度分秒

**函数**：`DegToDms`

**功能**：度转度分秒。

**语法**：

```vbs
SSProcess.DegToDms(deg)
```

**参数**：

- deg 度，可以用小数表示

**说明**：

执行该函数，将度转为度分秒

**示例**：

```vbs
Sub OnClick()
D= SSProcess.DegToDms(44.3)
Msgbox D
End Sub
```


#### DegToRadian 度转弧度

**函数**：`DegToDms`

**功能**：度转度分秒。

**语法**：

```vbs
SSProcess.DegToDms(deg)
```

**参数**：

- deg 度，可以用小数表示

**说明**：

执行该函数，将度转为度分秒

**示例**：

```vbs
Sub OnClick()
D= SSProcess.DegToDms(44.3)
Msgbox D
End Sub
```


#### DmmToDeg 度分分转度

**函数**：`DmmToDeg`

**功能**：度分分转度。

**语法**：

```vbs
SSProcess.DmmToDeg(dmm)
```

**参数**：

- dmm 度分分值

**说明**：

执行该函数，将度分分转为度

**示例**：

```vbs
Sub OnClick()
dmm=179.123423243
Angle=SSProcess.DmmToDeg(dmm)
Msgbox Angle
End Sub
```


#### DmsToDeg 度分秒转度

**函数**：`DmsToDeg`

**功能**：度分秒转度。

**语法**：

```vbs
SSProcess.DmsToDeg(dms)
```

**参数**：

- dms 度分秒值

**说明**：

执行该函数，将度分秒转为度

**示例**：

```vbs
下面示例将179度12分34秒转换成度：
Sub OnClick()
dms=179.1234
Angle=SSProcess.DmsToDeg(dms)
Msgbox Angle
End Sub
```


#### DmsToRadian 度分秒转弧度

**函数**：`DmsToRadian`

**功能**：度分秒转弧度。

**语法**：

```vbs
SSProcess.DmsToRadian(dms)
```

**参数**：

- dms 度分秒值

**说明**：

执行该函数，将度分秒转为弧度

**示例**：

```vbs
下面示例将179度59分60秒转换成弧度：
Sub OnClick()
dms=179.5960
angle=SSProcess.DmsToRadian( dms )
End Sub
```


#### DmsToSec 度分秒转秒

**函数**：`DmsToSec`

**功能**：度分秒转秒。

**语法**：

```vbs
SSProcess.DmsToSec(dms)
```

**参数**：

- dms 度分秒值

**说明**：

执行该函数，将度分秒转为秒。

**示例**：

```vbs
下面示例将179度59分60秒转换成秒：
Sub OnClick()
dms=179.5960
angle=SSProcess.DmsToSec(dms )
End Sub
```


#### RadianToDeg 弧度转度

**函数**：`RadianToDeg`

**功能**：弧度转度。

**语法**：

```vbs
SSProcess.RadianToDeg(radian)
```

**参数**：

- radian 要转换的角度的弧度值

**说明**：

执行该函数，将弧度转为度。

**示例**：

```vbs
Sub OnClick()
radian=10
SSProcess.RadianToDeg (radian)
End Sub
```


#### RadianToDms 弧度转度分秒

**函数**：`RadianToDms`

**功能**：弧度转度分秒。

**语法**：

```vbs
SSProcess.RadianToDms(radian)
```

**参数**：

- radian 要转换的角度的弧度值

**说明**：

执行该函数，将弧度转为度分秒。

**示例**：

```vbs
Sub OnClick()
Radian =10
aa =SSProcess.RadianToDms (radian)
Msgbox aa
End Sub
```


#### SecToDms 秒转度分秒

**函数**：`SecToDms`

**功能**：秒转度分秒。

**语法**：

```vbs
SSProcess.SecToDms(sec)
```

**参数**：

- radian 地物的秒值

**说明**：

使用此函数将地物角度从秒值转换为度分秒形式。

**示例**：

```vbs
下面示例将秒转换为度分秒值：
Sub OnClick()
sec = 251999.969142551
A = SSProcess. SecToDms (sec)
End Sub
```


#### XYSA 坐标正反算函数

**函数**：`XYSA`

**功能**：坐标正反算函数。

**语法**：

```vbs
SSProcess.XYSA(x0, y0, x1, y1, dist, angle, flag)
```

**参数**：

- x0 第1点X坐标
- y0 第1点Y坐标
- x1 第2点X坐标
- y1 第2点Y坐标
- dist 第1到第2点距离，米为单位
- angle 第1到第2点方位角,弧度为单位
- flag 正反算方式，0(反算) 1(正算)

**说明**：

执行该函数，可实现坐标的正反算。
正算时，flag为1，x0、y0、dist、angle为已知值，x1、y1为返回值。
反算时，flag为0，x0、y0、x1、y1为已知值，dist、angle为返回值。
点旋转角与方位角的关系如下图，换算公式：方位角 =2π-点旋转角（SSObj_angle)，此处的点旋转角度为数据存储值，单位是弧度。

**示例**：

```vbs
下面示例坐标正算：
Sub OnClick()
Dim x0, y0, x1, y1, dist, angle, flag
x0 = 100
y0 = 200
dist = 50
angle = 90/180*3.1415926
flag = 1
SSProcess.XYSA x0,y0,x1,y1,dist,angle,flag
End Sub
下面示例坐标反算：
Sub OnClick()
Dim x0, y0, x1, y1, dist, angle, flag
x0 = 100
y0 = 200
x1 = 500
y1 = 800
flag = 0
SSProcess.XYSA x0,y0,x1,y1,dist,angle,flag
End Sub
```


---

### 数据检查

#### AddCheckRecord 添加检查记录

**函数**：`AddCheckRecord`

**功能**：添加检查记录。

**语法**：

```vbs
SSProcess.AddCheckRecord( group, check, checkmodel, description, x, y, z, objtype, geoids, noteids )
```

**参数**：

- group 检查组名
- check 检查项名
- checkmodel 检查模型名称,如"脚本检查类->地物编码合法性检查"
- description 检查错误描述
- x 错误位置X坐标
- y 错误位置Y坐标
- z 错误位置Z坐标
- objtype 错误对象类型 0 点 1 线 2 面 3 注记
- geoids 错误相关地物ID列表,多个时用,号分隔
- noteids 错误相关注记ID列表,多个时用,号分隔

**说明**：

执行该函数，将往系统检查记录列表中增加检查记录。

**示例**：

```vbs
下面示例检查图中编码不等于6位的对象：
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT,LINE,AREA"
SSProcess.SetSelectCondition "SSObj_Code", "<", "1000000"
SSProcess.SetSelectCondition "SSObj_Code", ">", "99999"
SSProcess.SelectFilter
SSProcess.RemoveCheckRecord "基本属性检查", "编码位数合法性检查"
geoCount = SSProcess.GetSelGeoCount()
For i=0 To geoCount-1
geoID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
geoType = SSProcess.GetSelGeoValue(i, "SSObj_Type")
SSProcess.GetSelGeoPoint i, 0, x, y, z, pointtype, name )
SSProcess.AddCheckRecord "基本属性检查", "编码位数合法性检查", "自定义脚本检查类->编码位数合法性检查", "编码位数不等于6", x, y, z, geoType, geoID, ""
Next
SSProcess.ShowCheckOutput
End Sub
```


#### ClearCheckRecord 清空检查记录

**函数**：`ClearCheckRecord`

**功能**：ClearCheckRecord 函数

**语法**：

```vbs
SSProcess.ClearCheckRecord ()
```

**说明**：

使用此函数清空生成的检查记录

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearCheckRecord
End Sub
```


#### GetCheckRecordCount 获取检查记录总数

**函数**：`GetCheckRecordCount`

**功能**：获取检查记录总数。

**语法**：

```vbs
SSProcess.GetCheckRecordCount()
```

**说明**：

该函数为数据检查函数，用于获取检查总数。

**示例**：

```vbs
Dim strValues(100), geoID(100), geoType(100)
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", ">" , "999999"
SSProcess.SelectFilter
SSProcess.RemoveCheckRecord "基本属性检查", "编码位数合法性检查"
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount -1
geoID(i) = SSProcess.GetSelGeoValue(i, "SSObj_ID")
geoType(i) = SSProcess.GetSelGeoValue(i, "SSObj_Type")
SSProcess.GetSelGeoPoint i , 0, x, y, z, pointtype , name
SSProcess.AddCheckRecord "基本属性检查" , "编码位数合法性检查", "自定义脚本检查类->编码位数合法性检查" , "编码位数大于6", x, y , z, pointtype , geoID(i), ""
Next
nRecordUpper = SSProcess.GetCheckRecordCount- 1  '获取检查记录总数
strFields = "X,Y,GEOIDLIST,NOTEIDLIST,DESCRIPTION"
outstrValues =""
For d=0 To nRecordUpper
SSProcess.GetCheckRecordValue d, strFields, strValues (d) '获取检查记录
If outstrValues ="" Then
outstrValues =strValues( d)
Else
outstrValues =outstrValues&chr(13)&strValues(d)
End If
Next
Msgbox outstrValues
End Sub
```


#### GetCheckRecordValue 获取检查记录

**函数**：`GetCheckRecordValue`

**功能**：获取检查记录。

**语法**：

```vbs
SSProcess.GetCheckRecordValue(nIndex,strFields,strValues)
```

**参数**：

- nIndex 检查记录的索引号 ,从0开始
- strFields 要获取的字段名称
- strValues 获取的字段值

**说明**：

该函数为数据检查函数，用于获取检查记录。

**示例**：

```vbs
Dim strValues(100), geoID(100), geoType(100)
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", ">" , "999999"
SSProcess.SelectFilter
SSProcess.RemoveCheckRecord "基本属性检查", "编码位数合法性检查"
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount -1
geoID(i) = SSProcess.GetSelGeoValue(i, "SSObj_ID")
geoType(i) = SSProcess.GetSelGeoValue(i, "SSObj_Type")
SSProcess.GetSelGeoPoint i , 0, x, y, z, pointtype , name
SSProcess.AddCheckRecord "基本属性检查" , "编码位数合法性检查", "自定义脚本检查类->编码位数合法性检查" , "编码位数大于6", x, y , z, pointtype , geoID(i), ""
Next
nRecordUpper = SSProcess.GetCheckRecordCount- 1  '获取检查记录总数
strFields = "X,Y,GEOIDLIST,NOTEIDLIST,DESCRIPTION"
outstrValues =""
For d=0 To nRecordUpper
SSProcess.GetCheckRecordValue d, strFields, strValues (d) '获取检查记录
If outstrValues ="" Then
outstrValues =strValues( d)
Else
outstrValues =outstrValues&chr(13)&strValues( d)
End If
Next
Msgbox outstrValues
End Sub
```


#### LoadCheckRecord 读取检查记录

**函数**：`LoadCheckRecord`

**功能**：LoadCheckRecord 函数

**语法**：

```vbs
SSProcess.LoadCheckRecord()
```

**说明**：

执行该函数，将从数据库中读取所有检查记录到内存。

**示例**：

```vbs
Sub OnClick()
SSProcess.LoadCheckRecord
End Sub
```


#### RemoveCheckRecord 删除检查记录

**函数**：`RemoveCheckRecord`

**功能**：删除检查记录。

**语法**：

```vbs
SSProcess.RemoveCheckRecord( group, check )
```

**参数**：

- group 检查组名, 为空时,删除所有组检查记录
- check 检查项名, 为空时,删除组内所有检查项检查记录

**说明**：

执行该函数，将删除指定检查组名和检查项名的检查记录。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT,LINE,AREA"
SSProcess.SetSelectCondition "SSObj_Code", "<", "1000000"
SSProcess.SetSelectCondition "SSObj_Code", ">", "99999"
SSProcess.SelectFilter
SSProcess.RemoveCheckRecord "基本属性检查", "编码位数合法性检查"
geoCount = SSProcess.GetSelGeoCount()
For i=0 To geoCount-1
geoID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
geoType = SSProcess.GetSelGeoValue(i, "SSObj_Type")
SSProcess.GetSelGeoPoint i, 0, x, y, z, pointtype, name )
SSProcess.AddCheckRecord "基本属性检查", "编码位数合法性检查", "自定义脚本检查类->编码位数合法性检查", "编码位数不等于6", x, y, z, geoType, geoID, ""
Next
SSProcess.ShowCheckOutput
SSProcess.SaveCheckRecord
End Sub
```


#### SaveCheckRecord 保存检查记录

**函数**：`SaveCheckRecord`

**功能**：SaveCheckRecord 函数

**语法**：

```vbs
SSProcess.SaveCheckRecord()
```

**说明**：

执行该函数，将保存内存中的所有检查记录到数据库。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT,LINE,AREA"
SSProcess.SetSelectCondition "SSObj_Code", "<", "1000000"
SSProcess.SetSelectCondition "SSObj_Code", ">", "99999"
SSProcess.SelectFilter
SSProcess.RemoveCheckRecord "基本属性检查", "编码位数合法性检查"
geoCount = SSProcess.GetSelGeoCount()
For i=0 To geoCount-1
geoID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
geoType = SSProcess.GetSelGeoValue(i, "SSObj_Type")
SSProcess.GetSelGeoPoint i, 0, x, y, z, pointtype, name )
SSProcess.AddCheckRecord "基本属性检查", "编码位数合法性检查", "自定义脚本检查类->编码位数合法性检查", "编码位数不等于6", x, y, z, geoType, geoID, ""
Next
SSProcess.ShowCheckOutput
SSProcess.SaveCheckRecord
End Sub
```


#### ShowCheckOutput 显示检查结果

**函数**：`ShowCheckOutput`

**功能**：ShowCheckOutput 函数

**语法**：

```vbs
SSProcess.ShowCheckOutput()
```

**说明**：

执行该函数，将弹出检查记录输出对话框

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT,LINE,AREA"
SSProcess.SetSelectCondition "SSObj_Code", "<", "1000000"
SSProcess.SetSelectCondition "SSObj_Code", ">", "99999"
SSProcess.SelectFilter
SSProcess.RemoveCheckRecord "基本属性检查", "编码位数合法性检查"
geoCount = SSProcess.GetSelGeoCount()
For i=0 To geoCount-1
geoID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
geoType = SSProcess.GetSelGeoValue(i, "SSObj_Type")
SSProcess.GetSelGeoPoint i, 0, x, y, z, pointtype, name )
SSProcess.AddCheckRecord "基本属性检查", "编码位数合法性检查", "自定义脚本检查类->编码位数合法性检查", "编码位数不等于6", x, y, z, geoType, geoID, ""
Next
SSProcess.ShowCheckOutput
SSProcess.SaveCheckRecord
End Sub
```


#### WriteCheckRecord 保存检查记录到数据库

**函数**：`WriteCheckRecord`

**功能**：保存检查记录到数据库。

**语法**：

```vbs
SSProcess.WriteCheckRecord()
```

**说明**：

使用此函数将数据监理检查的错误记录保存到数据库中，相当于EPS中数据监理功能中的保存检查记录。

**示例**：

```vbs
Sub OnClick()
SSProcess.WriteCheckRecord
End Sub
```


---

### 数据库操作

#### AccessIsEOF 是否是记录集的尾部

**函数**：`AccessIsEOF`

**功能**：是否是记录集的尾部。

**语法**：

```vbs
SSProcess.AccessIsEOF( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句

**说明**：

执行该函数，获取是否是记录集的尾部,是返回1, 否返回0。

**示例**：

```vbs
Sub OnClick()
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
sql ="SELECT GeoLineTB.ID FROM GeoLineTB WHERE [GeoLineTB].[Mark] Mod 2<>0 ;"
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount>0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql) = False)
SSProcess.GetAccessRecord projectName, sql, fields, values
SSProcess.AccessMoveNext projectName, sql
Wend
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Sub
```


#### AccessMove 移到记录集指定位置

**函数**：`AccessMove`

**功能**：移到记录集指定位置。

**语法**：

```vbs
SSProcess.AccessMove(mdbName, sql, lRows)
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句
- lRows 记录集指针要移动到的行号

**说明**：

执行该函数，将把记录集指针移到记录集指定位置。


#### AccessMoveFirst 移到记录集头部

**函数**：`AccessMoveFirst`

**功能**：移到记录集头部。

**语法**：

```vbs
SSProcess.AccessMoveFirst( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句

**说明**：

执行该函数，将把记录集指针移到记录头部。

**示例**：

```vbs
Sub OnClick()
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
sql ="SELECT GeoLineTB.ID FROM GeoLineTB WHERE [GeoLineTB].[Mark] Mod 2<>0 ;"
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount>0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.GetAccessRecord projectName, sql, fields, values
SSProcess.AccessMoveNext projectName, sql
Wend
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Sub
```


#### AccessMoveLast 移到记录集尾部

**函数**：`AccessMoveLast`

**功能**：移到记录集尾部。

**语法**：

```vbs
SSProcess.AccessMoveLast( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句

**说明**：

执行该函数，将把记录集指针移到记录集尾部。

**示例**：

```vbs
Sub OnClick()
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
sql ="SELECT GeoLineTB.ID FROM GeoLineTB WHERE [GeoLineTB].[Mark] Mod 2<>0 ;"
SSProcess.OpenAccessRecordset projectName, sql
SSProcess. AccessMoveLast projectName, sql
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Sub
```


#### AccessMoveNext 移到记录集下一个

**函数**：`AccessMoveNext`

**功能**：移到记录集下一个。

**语法**：

```vbs
SSProcess.AccessMoveNext( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句

**说明**：

执行该函数，将把记录集指针移到记录集下一个。

**示例**：

```vbs
Sub OnClick()
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
sql ="SELECT GeoLineTB.ID FROM GeoLineTB WHERE [GeoLineTB].[Mark] Mod 2<>0 ;"
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount>0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.GetAccessRecord projectName, sql, fields, values
SSProcess.AccessMoveNext  projectName, sql
Wend
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Sub
```


#### AccessMovePrev 移到记录集上一个

**函数**：`AccessMovePrev`

**功能**：移到记录集上一个。

**语法**：

```vbs
SSProcess.AccessMovePrev( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句

**说明**：

执行该函数，将把记录集指针 移到记录集上一个。


#### AddAccessRecord 添加表记录

**函数**：`AddAccessRecord`

**功能**：添加表记录。

**语法**：

```vbs
SSProcess.AddAccessRecord( mdbName,sql,fields,values )
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句
- fields 添加数据的字段名称列表,字段间用逗号分隔
- values 添加数据的字段值列表,值间用逗号分隔

**说明**：

执行该函数，将往记录集中添加数据记录。


#### CloseAccessMdb 关闭 Access MDB 文件

**函数**：`CloseAccessMdb`

**功能**：关闭Access Mdb文件。

**语法**：

```vbs
SSProcess.CloseAccessMdb( mdbName )
```

**参数**：

- mdbName 要关闭的数据库文件名称

**说明**：

执行该函数，将关闭指定数据库文件。

**示例**：

```vbs
Sub OnClick()
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
sql ="SELECT GeoLineTB.ID FROM GeoLineTB WHERE [GeoLineTB].[Mark] Mod 2<>0 ;"
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount>0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.GetAccessRecord projectName, sql, fields, values
SSProcess.AccessMoveNext projectName, sql
Wend
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Sub
```


#### CloseAccessRecordset 关闭数据表记录集

**函数**：`CloseAccessRecordset`

**功能**：关闭数据表记录集。

**语法**：

```vbs
SSProcess.CloseAccessRecordset( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql 数据表名

**说明**：

执行该函数，将关闭指定sql语句的记录集。

**示例**：

```vbs
Sub OnClick()
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
sql ="SELECT GeoLineTB.ID FROM GeoLineTB WHERE [GeoLineTB].[Mark] Mod 2<>0 ;"
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount>0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.GetAccessRecord projectName, sql, fields, values
SSProcess.AccessMoveNext projectName, sql
Wend
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Sub
```


#### CreateAccessMdb 创建 Access MDB 文件

**函数**：`CreateAccessMdb`

**功能**：创建Access Mdb数据库文件。

**语法**：

```vbs
SSProcess.CreateAccessMdb( templateName, mdbName )
```

**参数**：

- templateName 模板数据库文件名称,允许为空
- mdbName 要创建的数据库文件名称

**说明**：

执行该函数，将根据 指定的模板,创建新数据库文件,如果模板文件为空,则创建空数据库。

**示例**：

```vbs
Sub OnClick()
SSProcess.CreateAccessMdb "GB_500.mdt","基础地理"
End Sub
```


#### CreateAccessTable 根据 SQL 语句创建数据表

**函数**：`CreateAccessTable`

**功能**：根据SQL语句创建数据表。

**语法**：

```vbs
SSProcess.CreateAccessTable( mdbName,createTabSql)
```

**参数**：

- mdbName 数据库文件名称
- createTabSql 创建表Sql语句

**说明**：

执行该函数，将根据指定Sql语句创建数据表。

**示例**：

```vbs
'创建管线点表字段描述
Function CreatePointFieldStruct()
Dim strFields
strFields = ""
strFields = strFields + "[ID] autoincrement,"
strFields = strFields + "[Exp_No] varchar(9) NULL," '物探点号
CreatePointFieldStruct = strFields
End Function
'创建管线点表
Function CreateExportPointTable( ByVal mdbName, ByVal tableName )
fieldStructs = CreatePointFieldStruct
createTabSql = "CREATE TABLE [" & tableName & "] (" & fieldStructs & ");"
SSProcess.CreateAccessTable mdbName, createTabSql
End Function
Sub OnClick()
mdbName = "c:\c.mdb"
SSProcess.CreateAccessMdb "", mdbName
SSProcess.OpenAccessMdb mdbName
CreateExportPointTable mdbName, "管线点属性表"
SSProcess.CloseAccessMdb mdbName
End Sub
```


#### CreateAccessTable1 创建指定图层和类型的属性表

**函数**：`CreateAccessTable1`

**功能**：创建指定图层和类型的属性表。

**语法**：

```vbs
SSProcess.CreateAccessTable1( mdbName,layerName, objecType)
```

**参数**：

- mdbName 数据库文件名称
- layerName 图层名
- objecType 图层对象类型 0 点 1 线 2 面 3 注记

**说明**：

执行该函数，将根据数据转换脚本中设定的属性表对照关系,创建指定图层和类型所对应的属性表。


#### DelAccessRecord 删除表记录

**函数**：`DelAccessRecord`

**功能**：删除表记录。

**语法**：

```vbs
SSProcess.DelAccessRecord( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句

**说明**：

执行该函数，将删除记录集的当前记录。


#### ExecuteAccessSql 执行 Sql 语句

**函数**：`ExecuteAccessSql`

**功能**：执行Sql语句。

**语法**：

```vbs
SSProcess.ExecuteAccessSql( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql sql查询语句

**说明**：

执行该函数，将在指定数据库中执行指定的sql语句。
执行前需要打开数据库进行操作。

**示例**：

```vbs
sub onclick()
mdbName= SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb mdbName
sql = "UPDATE ZD_权属调查 SET ZD_权属调查.QSDCJS = '123' WHERE (([ZD_权属调查].[QSDCJS]='1'));"
SSProcess.ExecuteAccessSql mdbName,sql
SSProcess.CloseAccessMdb mdbName
End Sub
```


#### ExportSelectionToAccessTable 输出选择集中指定图层和类型的属性数据

**函数**：`ExportSelectionToAccessTable`

**功能**：输出选择集中指定图层和类型的属性数据。

**语法**：

```vbs
SSProcess.ExportSelectionToAccessTable( mdbName,layerName, objecType)
```

**参数**：

- mdbName 数据库文件名称
- layerName 图层名
- objecType 图层对象类型 0 点 1 线 2 面 3 注记

**说明**：

执行该函数，将输出指定图层和类型的属性数据到相对应的数据表,数据表需先用函数CreateAccessTable1创建。需有一个空的结构完整的AccessMDB作为模板。


#### GetAccessFieldInfo 获取数据库表字段结构

**函数**：`GetAccessFieldInfo`

**功能**：获取数据库表字段结构。

**语法**：

```vbs
SSProcess.GetAccessFieldInfo(mdbName, tableName, fieldInfos)
```

**参数**：

- mdbName 数据库文件名称
- tableName 数据表名
- fieldInfos 返回的字段信息列表,字段信息间用分号分隔,字段内信息用逗号分隔
- 一个字段的信息内容包括: 字段名称,数据类型,字段大小,字段属性,字段序号,是否必须字段,是否允许为空,排序比较方式,字段别名,源字段名,源表名,字段规则,字段规则内容,缺省值。

**说明**：

执行该函数，将返回指定表的所有字段结构信息。

**示例**：

```vbs
Sub OnClick()
EDBFilename = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb EDBFilename
tablename = SSProcess.GetCodeAttrTableName (3103003, 0)
Msgbox SSProcess.GetAccessFieldInfo(EDBFilename, tableName, fieldInfos)
End Sub
```


#### GetAccessFieldInfo1 获取数据库表指定字段结构

**函数**：`GetAccessFieldInfo1`

**功能**：获取数据库表指定字段结构。

**语法**：

```vbs
SSProcess.GetAccessFieldInfo1(mdbName,tableName,fieldName,fieldInfos)
```

**参数**：

- mdbName 数据库文件名称
- tableName 数据表名
- fieldName 字段名称
- fieldInfos 返回的字段信息列表,字段内信息用逗号分隔
- 字段的信息内容包括: 字段名称,数据类型,字段大小,字段属性,字段序号,是否必须字段,是否允许为空,排序比较方式,字段别名,源字段名,源表名,字段规则,字段规则内容,缺省值

**说明**：

执行该函数，将返回指定表指定字段的结构信息。

**示例**：

```vbs
Sub OnClick()
EDBFilename = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb EDBFilename
tablename = SSProcess.GetCodeAttrTableName (3103003, 0 )
fileName = "房屋层数"
SSProcess.GetAccessFieldInfo1 EDBFilename, tableName,fileName, fieldInfos
Msgbox fieldInfos
End Sub
```


#### GetAccessRecord 获取表记录

**函数**：`GetAccessRecord`

**功能**：获取表记录

**语法**：

```vbs
SSProcess.GetAccessRecord( mdbName,sql,fields,values)
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句
- fields 获取数据的字段名称列表,字段间用逗号分隔
- values 获取数据的字段值列表,值间用逗号分隔

**说明**：

执行该函数，将获取当前记录集的当前记录字段值列表。

**示例**：

```vbs
Sub OnClick()
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
sql ="SELECT GeoLineTB.ID FROM GeoLineTB WHERE [GeoLineTB].[Mark] Mod 2<>0 ;"
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount>0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.GetAccessRecord projectName, sql, fields, values
SSProcess.AccessMoveNext projectName, sql
Wend
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Sub
```


#### GetAccessRecordCount 获取表记录集记录个数

**函数**：`GetAccessRecordCount`

**功能**：获取表记录集记录个数

**语法**：

```vbs
SSProcess.GetAccessRecordCount( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句

**说明**：

执行该函数，将返回记录集内的记录个数。

**示例**：

```vbs
Sub OnClick()
edbname = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb edbname
sql= "Select 姓名,口令,检查状态 from 质检验收人员表 Where 姓名='" &user & "'"
SSProcess.OpenAccessRecordset edbname, sql
recordc=SSProcess.GetAccessRecordCount(edbname, sql)
If recordc=1 Then
SSProcess.ModifyAccessRecord edbname, sql, "姓名,口令,检查状态", values1 & "," & func
Else
SSProcess.AddAccessRecord edbName, sql, "名称,口令,检查状态", user & ",1," & func
End If
SSProcess.CloseAccessRecordset edbname, sql
SSProcess.CloseAccessMdb edbname
End Sub
```


#### GetAccessTableNames 获取数据库表名列表

**函数**：`GetAccessTableNames`

**功能**：获取数据库表名列表。

**语法**：

```vbs
SSProcess.GetAccessTableNames( mdbName,tableNames)
```

**参数**：

- mdbName 数据库文件名称
- tableNames 返回的表名列表,有多个时,用逗号分隔

**说明**：

执行该函数，将返回指定数据库文件中的所有数据表名, 返回值为表的个数。

**示例**：

```vbs
Sub OnClick()
SSProcess.GetAccessTableNames "基础地理", tableNames
MsgBox tableNames
End Sub
```


#### IsExistentTable 判断指定表是否存在

**函数**：`IsExistentTable`

**功能**：判断指定表是否存在。

**语法**：

```vbs
SSProcess.IsExistentTable( mdbName,tableName )
```

**参数**：

- mdbName 要表所在的数据库文件名称
- tableName 要判断的表名称

**说明**：

执行该函数，将判断指定的数据表是否存在,如果存在返回 1, 不存在则返回 0。

**示例**：

```vbs
'**************************************************************
'@Function : NotExistTables
'@abstract : 该函数从给出的表名中筛选出当前 EDB工程中不存在的表
'@Parameters : tableNames ：需判断的表名（多个表名中间用英文逗号隔开，
'如tableNames="标注属性表,测量控制点属性表"）
'**************************************************************
Function NotExistTables(tableNames)
Dim arrName(1000)
NotExistTables=""
mdbName =SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb mdbName
SSFunc.ScanString tableNames, "," , arrName, nCount
For i = 0 To nCount - 1
If SSProcess.IsExistentTable (mdbName, arrName(i))=0 Then
If NotExistTables = "" Then
NotExistTables = arrName(i)
Else
NotExistTables =NotExistTables& ","& arrName(i)
End If
End If
Next
SSProcess.CloseAccessMdb mdbName
End Function
Sub OnClick()
Msgbox NotExistTables("EDB,FBI,GeoLineTB")
End Sub
```


#### ModifyAccessRecord 修改表记录

**函数**：`ModifyAccessRecord`

**功能**：修改表记录。

**语法**：

```vbs
SSProcess.ModifyAccessRecord( mdbName,sql,fields,values)
```

**参数**：

- mdbName 数据库文件名称
- sql 记录集查询语句
- fields 修改数据的字段名称列表,字段间用逗号分隔
- values 修改数据的字段值列表,值间用逗号分隔

**说明**：

执行该函数，将修改当前记录集的当前记录为指定的字段值。

**示例**：

```vbs
Sub OnClick()
mdbname=SSProcess.GetProjectFileName()
result=SSProcess.OpenAccessMdb( mdbname)
sql ="SELECT房屋面属性表.层数 FROM GeoAreaTB INNER JOIN房屋面属性表 ON GeoAreaTB.ID =房屋面属性表.ID WHERE (([GeoAreaTB].[Mark] Mod 2<>0) and (房屋面属性表.栋名='A栋'));"
result=SSProcess.OpenAccessRecordset( mdbName, sql)
If result>0 Then
SSProcess.AccessMoveFirst mdbname, sql
while (SSProcess.AccessIsEOF (mdbname, sql ) = False)
SSProcess.ModifyAccessRecord mdbName, sql, "层数" , "5"'输出到mdb表中
SSProcess.AccessMoveNext mdbname, sql
Wend
End If
SSProcess.CloseAccessRecordset mdbname, sql
SSProcess.CloseAccessMdb mdbName
End Sub
```


#### OpenAccessMdb 打开 Access MDB 文件

**函数**：`OpenAccessMdb`

**功能**：打开Access Mdb文件。

**语法**：

```vbs
SSProcess.OpenAccessMdb( mdbName )
```

**参数**：

- mdbName 要打开的数据库文件名称

**说明**：

执行该函数，将打开指定数据库文件,使用完成后需执行CloseAccessMdb关闭数据。

**示例**：

```vbs
Sub OnClick()
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
sql ="SELECT GeoLineTB.ID FROM GeoLineTB WHERE [GeoLineTB].[Mark] Mod 2<>0 ;"
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount>0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.GetAccessRecord projectName, sql, fields, values
SSProcess.AccessMoveNext projectName, sql
Wend
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Sub
```


#### OpenAccessRecordset 打开数据表记录集

**函数**：`OpenAccessRecordset`

**功能**：打开数据表记录集

**语法**：

```vbs
SSProcess.OpenAccessRecordset( mdbName,sql )
```

**参数**：

- mdbName 数据库文件名称
- sql 数据表名

**说明**：

执行该函数，将打开指定sql语句的记录集。

**示例**：

```vbs
Sub OnClick()
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
sql ="SELECT GeoLineTB.ID FROM GeoLineTB WHERE [GeoLineTB].[Mark] Mod 2<>0 ;"
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount>0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.GetAccessRecord projectName, sql, fields, values
SSProcess.AccessMoveNext projectName, sql
Wend
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Sub
```


---

### 数据整理

#### AdjustArcGISMdbBLCoord 调整 ArcGIS MDB 经纬度精度

**函数**：`AdjustArcGISMdbBLCoord`

**功能**：调整ArcGIS MDB经纬度精度。

**语法**：

```vbs
SSProcess.AdjustArcGISMdbBLCoord(fileName, ke, centralMeridian, stripWidth)
```

**参数**：

- fileName ArcGISMdb文件名
- ke 带宽（3°带或6°带）
- centralMeridian 中央子午线
- stripWidth 投影带宽

**说明**：

该函数为数据整理函数，执行后将调整ArcGIS_MDB经纬度精度。


#### AreaAttrToLabelAttr 属性复制（面 -> 点）

**函数**：`AreaAttrToLabelAttr`

**功能**：属性复制。

**语法**：

```vbs
SSProcess.AreaAttrToLabelAttr(areaCode, labelCode, filterFields, relFields)
```

**参数**：

- areaCode 面对象编码
- labelCode 属性点编码
- filterFields 过滤字段，也就是不需要匹配赋值的字段
- relFields 需要复制的字段；如果为空，则表示除了filterFields的字段全部赋值

**说明**：

把面属性复制到点属性。

**示例**：

```vbs
Sub OnClick()
'宗地面属性赋值给属性点
labelCode = "6803111"
areaCode = "6803153"
filterFields = "ID,FeatureGUID,BSM,YSDM,QLEMC,ZDH,SJYT"
relFields = ""
SSProcess.AreaAttrToLabelAttr areaCode,labelCode,filterFields,relFields
End Sub
```


#### ByteStackRestore 从二进制包获取数据

**函数**：`ByteStackRestore`

**功能**：从二进制包获取数据。

**语法**：

```vbs
SSProcess.ByteStackRestore(dataValue, pos, dataType, value)
```

**参数**：

- dataValue 二进制包存储值
- pos 整数，表示位置
- datatype 数据类型：2、short类型；3、long类型；4、float类型；5、double类型； 7、DateTime类型；8、CString类型
- value 从二进制包解析后的值

**说明**：

使用此函数将从二进制包中获取相应数值。


#### ByteStackStore 写数据到二进制包

**函数**：`ByteStackStore`

**功能**：写数据到二进制包。

**语法**：

```vbs
SSProcess.ByteStackStore(dataValue, value)
```

**参数**：

- dataValue 二进制存储位置
- value 数据值

**说明**：

执行该函数，写数据到二进制包。


#### ChangeCodeByDataTransMap 根据数据转换对照方案换码

**函数**：`ChangeCodeByDataTransMap`

**功能**：根据数据转换对照方案换码。

**语法**：

```vbs
SSProcess.ChangeCodeByDataTransMap(schemeName)
```

**参数**：

- schemeName 模板目录下的文件名称，可通过平台的“调入分类器”功能设置换码方案并创建此文件，此文件扩展名为 *.par，参数中不用加扩展名称。

**说明**：

使用此函数将根据对照方案对当前数据进行换码操作。

**示例**：

```vbs
下面示例将根据DataClassIfyDlg中的换码方案对当前数据进行换码。
Sub OnClick()
SSProcess.ChangeCodeByDataTransMap("DataClassIfyDlg")
End Sub
```


#### ChangeCodeCopy 指定 ID 换码复制

**函数**：`ChangeCodeCopy`

**功能**：指定ID换码复制。

**语法**：

```vbs
SSProcess.ChangeCodeCopy(geoID,newCode)
```

**参数**：

- geoID 需要复制的地物ID
- newCode 新地物的编码

**说明**：

使用此函数将对某地物进行复制，并将复制后的地物进行换码。

**示例**：

```vbs
Sub OnClick()
geoID=11
newCode= 211000
SSProcess.ChangeCodeCopy geoID,newCode
End Sub
```


#### ChangeCodeCopySelectionObj 换码复制选择集对象

**函数**：`ChangeCodeCopySelectionObj`

**功能**：换码复制选择集对象。

**语法**：

```vbs
SSProcess.ChangeCodeCopySelectionObj(code)
```

**参数**：

- code 转换新编码

**说明**：

执行该函数，将换码复制选择集对象。
如果选择集对象为注记,指定的编码为点对象时，生成新对象的点位取注记的中心点。
如果选择集对象为注记,指定的编码为线或面对象时，生成新对象的点位取注记的外接矩形。
如果选择集对象为线或面,指定的编码为点对象时，生成新对象的点位取线或面中心点。

**示例**：

```vbs
下面示例把编码为8521的地物换码复制为2110：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "8521"
SSProcess.SelectFilter
SSProcess.ChangeCodeCopySelectionObj 2110
End Sub
```


#### ClearSelectionObjNodeSymbol 自动节点取舍处理

**函数**：`ClearSelectionObjNodeSymbol`

**功能**：自动节点取舍处理。

**语法**：

```vbs
SSProcess.ClearSelectionObjNodeSymbol( dealLinkToFrame, dealInvalidAngle)
```

**参数**：

- dealLinkToFrame 与图廓相接点自动节点取舍处理, 0不处理, 1处理
- dealInvalidAngle 不合理顶点夹角自动节点取舍处理, 0不处理, 1处理, 取舍夹角范围为<30度,>120度

**说明**：

执行该函数，将对符合条件的地物自动进行节点取舍处理。

**示例**：

```vbs
下面示例对棚房进行节点取舍处理：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "2150"
SSProcess.SelectFilter
SSProcess.ClearSelectionObjNodeSymbol 1, 1
End Sub
```


#### CloseSelectionObj 选择集对象闭合处理

**函数**：`CloseSelectionObj`

**功能**：选择集对象闭合处理。

**语法**：

```vbs
SSProcess.CloseSelectionObj( closeMode, hangDist)
```

**参数**：

- closeMode 闭合方式，0（尾部加入首点）、1（满足悬距要求时，粘合首尾点） 2（满足悬距要求时，粘合首尾点，否则尾部加入首点）
- hangDist 首尾点允许悬挂距离，以米为单位

**说明**：

执行该函数，将对选择集对象作封闭处理。

**示例**：

```vbs
下面示例把编码为1的地物作封闭处理：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
SSProcess.CloseSelectionObj 1,0.01
End Sub
```


#### CopyObjectAttr 复制对象属性到指定对象

**函数**：`CopyObjectAttr`

**功能**：复制对象属性到指定对象。

**语法**：

```vbs
SSProcess.CopyObjectAttr(srcID, tagID, copyPoints, copyExAttr)
```

**参数**：

- srcID 原地物ID
- tagID 目标地物ID
- copyPoints 0目标地物坐标不变 1目标地物坐标复制原地物坐标
- copyExAttr 0 复制扩展属性 1不复制扩展属性


#### CreateSelectionAreaLabel 创建选择集拓扑面属性点

**函数**：`CreateSelectionAreaLabel`

**功能**：创建选择集拓扑面属性点。

**语法**：

```vbs
SSProcess.CreateSelectionAreaLabel(area2labelCodes)
```

**参数**：

- area2labelCodes 面转点的编码，面与点的编码用:号分隔，多条记录用;号分隔。

**说明**：

执行该函数，可创建选择集拓扑面属性点。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "2"
SSProcess.SelectFilter
SSProcess.CreateSelectionAreaLabel "2:2203041"
End Sub
```


#### CreateSelectionObjCenterLine 创建选择集对象的中心线

**函数**：`CreateSelectionObjCenterLine`

**功能**：创建选择集对象的中心线。

**语法**：

```vbs
SSProcess.CreateSelectionObjCenterLine(lineCode, areaRefWidth)
```

**参数**：

- lineCode 生成中心线的编码
- areaRefWidth 参考宽度

**说明**：

执行此函数可创建选择集对象的中心线。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "=", "3103003"
SSProcess.SelectFilter
SSProcess.CreateSelectionObjCenterLine 1, 7
End Sub
```


#### CreateUpdateCell 生成更新单元格

**函数**：`CreateUpdateCell`

**功能**：生成更新单元格。

**语法**：

```vbs
SSProcess.CreateUpdateCell(cellWidth, cellHeight, cellCode)
```

**参数**：

- cellWidth 单元格宽度
- cellHeight 单元格高度
- cellCode 单元格编码

**说明**：

执行此函数可生成更新单元格。

**示例**：

```vbs
Sub OnClick()
SSProcess.CreateUpdateCell 10, 10, 2
End Sub
```


#### DangleCleanLineToFixedLine 有固定边悬挂处理

**函数**：`DangleCleanLineToFixedLine`

**功能**：有固定边悬挂处理。

**语法**：

```vbs
SSProcess.DangleCleanLineToFixedLine(fixLineCodes, lineCodes, limit, fixedLineAllowInsertPt)
```

**参数**：

- fixLineCodes 固定边编码，多个时用“ ,”分隔
- lineCodes 悬挂线编码，多个时用“ ,”分隔
- limit 限距（米)
- fixedLineAllowInsertPt 是否允许在固定边上插点，0表示不允许，1表示允许

**示例**：

```vbs
Sub OnClick()
SSProcess.DangleCleanLineToFixedLine 3109004,2203052,3,1
SSProcess.RefreshView
End Sub
```


#### DangleCleanLineToLine 线线悬挂处理

**函数**：`DangleCleanLineToLine`

**功能**：线线悬挂处理。

**语法**：

```vbs
SSProcess.DangleCleanLineToLine(lineCodes, limit, bDealRepeatLine, bCanInertPtTo_P_Y)
```

**参数**：

- lineCodes 悬挂线编码，多个时用“ ,”分隔
- limit 限距（米）
- bDealRepeatLine 是否删除与它线重叠的线头，0表示不删除，1表示删除
- bCanInertPtTo_P_Y 是否在P、Y类地物上插点，0表示不插点，1表示插点

**示例**：

```vbs
Sub OnClick()
SSProcess.DangleCleanLineToLine "2203052", "3" , "1", "1"
End Sub
```


#### ExecuteSql 执行 Sql 更新查询语句

**函数**：`ExecuteSql`

**功能**：执行Sql更新查询语句。

**参数**：

- sqlString SQL语句

**说明**：

执行该函数，可实现对EPS MDB数据库中数据表的更新查询。


#### ExecuteToolboxCommand 执行工具箱命令

**函数**：`ExecuteToolboxCommand`

**功能**：执行工具箱命令。

**语法**：

```vbs
SSProcess.ExecuteToolboxCommand(pageName, itemName)
```

**参数**：

- pageName 工具箱页面名
- itemName 工具箱工具名

**示例**：

```vbs
Sub OnClick()
SSProcess.ExecuteToolboxCommand "500到2000缩编","缩编处理"
End Sub
```


#### GetObjectBinaryAttr 获取对象二进制属性值

**函数**：`GetObjectBinaryAttr`

**功能**：获取二进制对象属性值。

**语法**：

```vbs
SSProcess.GetObjectBinaryAttr(index, attrField)
```

**参数**：

- index 对象的索引号,从0开始
- attrField 对象属性字段名

**说明**：

使用此函数获取对象二进制属性值。
返回值即为指定字段所对应的属性值。


#### GetSelGeoBinaryAttr 获取选择对象二进制属性值

**函数**：`GetSelGeoBinaryAttr`

**功能**：获取选择对象二进制属性值。

**语法**：

```vbs
SSProcess.GetSelGeoBinaryAttr(index, attrField)
```

**参数**：

- index 对象的索引号 ,从0开始
- attrField 对象属性字段名

**说明**：

使用此函数获取选择对象二进制属性值。
返回值即为指定字段所对应的属性值。


#### IdentIfySelectionObj 选择集对象符号识别

**函数**：`IdentIfySelectionObj`

**功能**：选择集对象符号识别。

**语法**：

```vbs
SSProcess.IdentIfySelectionObj( symbolIndex, objCode, deleteSrcObj, parameter)
```

**参数**：

- symbolIndex 识别符号索引，见下表：
| 列1 | 列2 | 列3 | 列4 |
|---|---|---|---|
| symbolIndex 索引号 | 符号名称 | parameter 参数列表 | 参数说明 |
| 1 | 虚线 | 数值误差 , 落笔距离 , 抬笔距离 |  |
| 2 | 点线 | 数值误差 , 小圆半径 ( 为 0 认为是点 ), 抬笔距离 |  |
| 3 | 点划线 | 数值误差 ,0, 抬笔距离 1, 落笔距离 1,... |  |
| 4 | 未加固坎 | 数值误差 , 坎毛短线长度 , 短线间隔 , 符号方向 (0 正向 1 反向 ) |  |
| 5 | 加固坎 | 数值误差 , 坎毛短线长度 , 短线间隔 , 坎点与坎顶距离 , 符号方向 (0 正向 1 反向 ) |  |
| 6 | 规则过街楼 | 数值误差 |  |
| 7 | 不规则过街楼 | 数值误差 , 边线编码 , 交叉线编码 |  |
| 8 | 围墙 | 数值误差 , 围墙宽度 , 符号方向 (0 正向 1 反向 ) |  |
| 9 | 栅栏 | 数值误差 , 小圆半径 , 短线长度 , 长线长度 , 符号方向 (0 正向 1 反向 ) |  |
| 10 | 棚房 | 数值误差 , 棚房短线长度 , 符号方向 (0 正向 1 反向 ) |  |
| 11 | 雨水污水检修井 | 数值误差 , 外圆半径 |  |
| 12 | 圆形下水篦子 | 数值误差 , 外圆半径 , 横线间隔 |  |
| 13 | 方形下水篦子 | 数值误差 , 外框长度 , 外框宽度 |  |
| 14 | 简单楼梯 | 数值误差 , 平台最小宽度 |  |
| 15 | 高压电力线 | 数值误差 |  |
| 16 | 低压电力线 | 数值误差 |  |
| 17 | 消火栓 | 数值误差 , 圆半径 ,T 横线长 , T 竖线长 |  |
| 18 | 圆形块 | 数值误差 , 圆半径 |  |
| 19 | 花圃符号 | 数值误差 , 底横线长 , 竖线长 , 左斜线长 , 右斜线长 |  |
| 20 | 草地符号 | 数值误差 , 竖线长 , 竖线间隔 |  |
| 21 | 旱地符号 | 数值误差 , 底横线长 , 竖线长 , 竖线间隔 |  |
| 22 | 双臂路灯 | 数值误差 , 圆半径 , 底圆与上圆圆心横距 , 底圆与上圆圆心纵距 , 长竖线长 , 短竖线长 |  |
| 23 | 单臂路灯左 | 数值误差 , 圆半径 , 上横线长 , 下横线长 , 竖线长 |  |
| 24 | 单臂路灯右 | 数值误差 , 圆半径 , 上横线长 , 下横线长 , 竖线长 |  |
- objCode 识别后转换编码
- deleteSrcObj 是否删除原地物 0(不删除) 1 删除
- parameter 符号参数列表，逗号分隔,详细参数见上表说明

**说明**：

执行该函数，将对选择集中对象进行符号识别。

**示例**：

```vbs
下面示例识别不规则过街楼地物：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "1,218002"
SSProcess.SelectFilter
SSProcess.IdentIfySelectionObj 7, 2180, 1, "0.001,218002,1"
End Sub
```


#### ImportMdbTable 导入 MDB 数据库表

**函数**：`ImportMdbTable`

**功能**：导入MDB数据库表。

**语法**：

```vbs
SSProcess.ImportMdbTable( mdbName, srcTable, destTable, mode)
```

**参数**：

- mdbName MDB数据库文件名
- srcTable 要导入的MDB数据库表名
- destTable 导入到工程文件的新表名
- mode 数据表导入方式, 0(只导入表结构) 1(同时导入表结构和表数据)

**说明**：

执行该函数，可实现导入指定MDB数据库的指定表到当前工程文件中。

**示例**：

```vbs
下面示例导入数据表：
Sub OnClick()
Dim mdbFile
mdbFile = SSProcess.SelectFileName (1, "选择MDB文件", 0 ,"MDB Files(*.mdb)|*.mdb|All Files (*.*)|*.*||")
SSProcess.ImportMdbTable mdbFile, "building", "DGNbuilding", 0
SSProcess.ExecuteSql "delete * from DGNbuilding"
SSProcess.ImportMdbTable mdbFile, "building", "DGNbuilding", 1
SSProcess.ImportMdbTable mdbFile, "watersystem", "DGNwatersystem", 0
SSProcess.ExecuteSql "delete * from DGNwatersystem"
SSProcess.ImportMdbTable mdbFile, "watersystem", "DGNwatersystem", 1
End Sub
```


#### InsertWaterLevelPoint 自动内插水面高程点

**函数**：`InsertWaterLevelPoint`

**功能**：自动内插水面高程点。

**语法**：

```vbs
SSProcess.InsertWaterLevelPoint(waterAreaSelectGroup,referenceHeightSelectGroup,searchRange, height, heightPointCode)
```

**参数**：

- waterAreaSelectGroup 水面边界面选择组名称
- referenceHeightSelectGroup 参考高程选择组名称
- searchRange 搜索参考高程点范围
- height 水面到地面的高度
- heightPointCode 水面高程点编码

**说明**：

执行该函数，将根据自动内插水面高程点。

**示例**：

```vbs
下面示例构面过程：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "水面线", "SSObj_Type", "=", "AREA"
SSProcess.SetSelectConditionGroup "水面线", "SSObj_Code", "=", "6240"
SSProcess.SetSelectConditionGroup "参考高程点", "SSObj_Type", "=", "POINT"
SSProcess.SetSelectConditionGroup "参考高程点", "SSObj_Code", "=", "8310"
SSProcess.InsertWaterLevelPoint "水面线", "参考高程点", 10, 0.5, 8310
End Sub
```


#### LabelAttrToAreaAttr 属性复制（点 -> 面）

**函数**：`LabelAttrToAreaAttr`

**功能**：属性复制（点->面）。

**语法**：

```vbs
SSProcess.LabelAttrToAreaAttr(labelCode, areaCode, filterFields, relFields)
```

**参数**：

- labelCode 属性点编码
- areaCode 面编码
- filterFields 过滤字段，也就是不需要匹配赋值的字段
- relFields 需要复制的字段；如果为空，则表示除了filterFields的字段全部赋值

**说明**：

执行该函数可以把属性点的属性值复制给对应的属性面。

**示例**：

```vbs
Sub OnClick()
labelCode = "6803111"
areaCode = "6803153"
filterFields = "ID,BSM,YSDM,DLDM"
relFields = ""
SSProcess.LabelAttrToAreaAttr labelCode, areaCode, filterFields, relFields
End Sub
```


#### LinkBreakLine 断线连接

**函数**：`LinkBreakLine`

**功能**：断线连接。

**语法**：

```vbs
SSProcess.LinkBreakLine(noDirCodes, haveDirCodes, limit)
```

**参数**：

- noDirCodes 无方向线编码
- haveDirCodes 有方向线编码
- limit 连接限距

**说明**：

执行该函数可以把连接限距内的同层、同码的断线进行连接。

**示例**：

```vbs
Sub OnClick()
noDirCodes = SSProcess.ReadEpsIni ("断线连接", "NoDirCodes", "3804073")
haveDirCodes = SSProcess.ReadEpsIni ("断线连接", "HaveDirCodes", "7602022")
limit = SSProcess.ReadEpsIni ("断线连接", "limit", "0.01" )
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "无方向线编码", noDirCodes, 0, "", "支持用->指定区间,多组时用逗号分隔, 如: 1->1000,2000->3000, 由系统自动排除指定的'有方向线编码'"
SSProcess.AddInputParameter "有方向线编码", haveDirCodes, 0, "", "支持用->区间指定,多组时用逗号分隔, 如: 1->1000,2000->3000"
SSProcess.AddInputParameter "连接限距", limit,0, "", "米为单位"
res = SSProcess.ShowInputParameterDlg ("断线连接")
If res = 0  Then
Exit Sub
End If
noDirCodes = SSProcess.GetInputParameter ("无方向线编码" )
haveDirCodes = SSProcess.GetInputParameter ("有方向线编码" )
limit = SSProcess.GetInputParameter ("连接限距")
SSProcess.WriteEpsIni "断线连接", "NoDirCodes", noDirCodes
SSProcess.WriteEpsIni "断线连接", "HaveDirCodes" , haveDirCodes
SSProcess.WriteEpsIni "断线连接", "limit" , limit
SSProcess.PushUndoMark
SSProcess.LinkBreakLine noDirCodes, haveDirCodes, CDbl(limit)
End Sub
```


#### LinkNearNoteObj 合并相近的两个注记

**函数**：`LinkNearNoteObj`

**功能**：合并相近的两个注记。

**语法**：

```vbs
SSProcess.LinkNearNoteObj( mainCondition, subCondition, searchRange, fieldName, fieldValue)
```

**参数**：

- mainCondition 注记选择条件分组名，由SetSelectConditionGroup函数设定
- subCondition 合并注记选择条件分组名，由SetSelectConditionGroup函数设定
- searchRange 以注记为中心搜索合并注记的范围，以米为单位
- fieldName 要修改注记的字段名称
- fieldValue 要修改注记的值方式,支持回调函数方式,由脚本定义处理函数来控制获取内容的方式，格式为： CallBackFunc_函数名称，（如："CallBackFunc_LinkText(SSObj_FontString, SSSubObj_FontString)")

**说明**：

执行该函数，将合并两个相近的注记。

**示例**：

```vbs
下面示例结构注记"砖"与层数合并：
Function CallBackFunc_LinkText(string1, string2)
CallBackFunc_LinkText = string1&string2
End Function
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "结构注记", "SSObj_FontClass", "==", "JMXZ"
SSProcess.SetSelectConditionGroup "结构注记", "SSObj_FontString", "==", "砖,砼"
SSProcess.SetSelectConditionGroup "结构注记", "SSObj_Type", "==", "NOTE"
SSProcess.SetSelectConditionGroup "层数注记", "SSObj_FontClass", "==", "JMXZ"
SSProcess.SetSelectConditionGroup "层数注记", "SSObj_FontString", "<>", "砖,砼"
SSProcess.SetSelectConditionGroup "层数注记", "SSObj_Type", "==", "NOTE"
SSProcess.LinkNearNoteObj "结构注记","层数注记",3,"SSObj_FontString","CallBackFunc_ LinkText(SSObj_FontString,SSSubObj_FontString)"
SSProcess.ClearSelectConditionGroups
End Sub
```


#### MapMethod 调用 Map 方法

**函数**：`MapMethod`

**功能**：调用Map方法。

**语法**：

```vbs
SSProcess.MapMethod( methodName, parameters )
```

**参数**：

- methodName 方法名称
- parameters 参数列表
- 常用方法列表
| 列1 | 列2 | 列3 |
|---|---|---|
| 方法名称 | 方法参数 | 说明 |
| LoadData | 图层名称列表（多个时用逗号分隔，为空时代表所有图层） | 装层指定图层数据到内存 |
| UnloadData | 图层名称列表（多个时用逗号分隔，为空时代表所有图层） | 从内存释放指定图层数据 |
| MoveScreen | X 坐标 ,Y 坐标 | 移动图形到指定窗口位置 |
| ZoomWindow | 左下角 X 坐标 , 左下角 Y 坐标 , 右上角 X 坐标 , 右上角 Y 坐标 , | 缩放图形到指定窗口范围 |
| ZoomIn | 无参数 | 放大窗口 |
| ZoomOut | 无参数 | 缩小窗口 |
| FreeGraphicInfoFile | 无参数 | 释放图形特征缓存 |
| FreeMemoDataFile | 无参数 | 释放外部信息缓存 |

**示例**：

```vbs
Sub OnClick()
'清空相关属性表的缓存
SSProcess.MapMethod "clearattrbuffer", "专业管点属性表"
SSProcess.MapMethod "clearattrbuffer", "专业管线属性表"
End Sub
```


#### MergeIslandAreaObj 合并带岛面

**函数**：`MergeIslandAreaObj`

**功能**：合并带岛面。

**语法**：

```vbs
SSProcess.MergeIslandAreaObj()
```

**说明**：

执行该函数，将合并选择集中包含岛的面。

**示例**：

```vbs
下面示例合并房屋层岛面：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "AREA"
SSProcess.SetSelectCondition "SSObj_LayerName", "==", "房屋面"
SSProcess.SelectFilter
SSProcess.MergeIslandAreaObj
End Sub
```


#### MergeNearGeoObj 合并相近的两地物

**函数**：`MergeNearGeoObj`

**功能**：合并相近的两地物。

**语法**：

```vbs
SSProcess.MergeNearGeoObj( mainCondition, mergeCondition, searchRange, insertPointCount， pointDirection， deleteMainObj)
```

**参数**：

- mainCondition 主地物选择条件分组名，由SetSelectConditionGroup函数设定
- mergeCondition 合并地物选择条件分组名，由SetSelectConditionGroup函数设定
- searchRange 以主地物为中心搜索合并地物的范围，以米为单位
- insertPointCount 从主地物中插入点数，0是所有点，其它为实际指定值
- pointDirection 边接方向类型，0与主地物头连接，1与主地物尾连接 2 自动判断最近点连接
- deleteMainObj 是否同时删除主地物,0（不删除）、1（删除）

**说明**：

执行该函数，将合并相近的两地物的空间点列。

**示例**：

```vbs
下面示例把编码为8360地物的空间点合并到与其相近的编码为8310的地物,并删除8360地物：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "8360地物", "SSObj_Code", "==", "8360"
SSProcess.SetSelectConditionGroup "3521地物", "SSObj_Code", "==", "3521"
SSProcess.MergeNearGeoObj "8360地物","3521地物",1.0,0,2,1
SSProcess.ClearSelectConditionGroups
End Sub
```


#### MergeNearLineObj 合并与线相近的点地物

**函数**：`MergeNearLineObj`

**功能**：合并与线相近的点地物。

**语法**：

```vbs
SSProcess.MergeNearLineObj(mainCondition,mergeCondition,searchRange,insertPointRange, deleteMainObj)
```

**参数**：

- mainCondition 主地物选择条件分组名，由SetSelectConditionGroup函数设定
- mergeCondition 合并地物选择条件分组名，由SetSelectConditionGroup函数设定
- searchRange 以主地物为中心搜索合并地物的范围，以米为单位
- insertPointRange 插入点与主点的最大距离，当点地物与线上点的距离小于该值时，只替换最近点，否则新插入一点
- deleteMainObj 是否同时删除主地物,0（不删除）、1（删除）

**说明**：

执行该函数，将合并与线地物相近的点地物的空间点点列。

**示例**：

```vbs
下面示例把编码为8360点地物的空间点合并到与其相近的编码为2110的线地物,并删除8360地物。
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "2110地物", "SSObj_Code", "==", "2110"
SSProcess.SetSelectConditionGroup "8360地物", "SSObj_Code", "==", "8360"
SSProcess.MergeNearLineObj "8360地物","2110地物",1.0, 0.5, 1
SSProcess.ClearSelectConditionGroups
End Sub
```


#### MergeNearNoteObj 合并与地物相近的注记

**函数**：`MergeNearNoteObj`

**功能**：合并与地物相近的注记。

**语法**：

```vbs
SSProcess.MergeNearNoteObj( mainCondition, mergeCondition, searchRange, insertPoint， objAttrField， getTextParameter， deleteMainObj)
```

**参数**：

- mainCondition 注记选择条件分组名，由SetSelectConditionGroup函数设定
- mergeCondition 合并地物选择条件分组名，由SetSelectConditionGroup函数设定
- searchRange 以注记为中心搜索合并地物的范围，以米为单位
- insertPoint 是否把注记的左中点加入合并地物，0不加，1加入
- objAttrField 更新合并地物的属性名称，见ChangeSelectionObjAttr函数中的attrField参数说明
- getTextParameter 取注记内容格式，支持RIGHT(字符数)，LEFT(字符数),MID(起始字符索引，字符数）三个取位函数的一个或多个的组合,如LEFT(2)+MID(4,2)+RIGHT(1),如果为空，则直接取注记内容。同时支持回调函数方式,由脚本定义处理函数来控制获取内容的方式，格式为： CallBackFunc_函数名称，（如："CallBackFunc_TrimString(SSObj_FontString)"）
- deleteMainObj 是否同时删除注记,0（不删除）、1（删除）

**说明**：

执行该函数，将合并与地物相近的注记内容到地物的指定属性字段中。

**示例**：

```vbs
下面示例把分类号为a的注记合并到与其相近的编码为8310的地物的点名属性,把注记的左右点位加入合并地物，并删除注记。
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "控制点注记", "SSObj_FontClass", "==", "a"
SSProcess.SetSelectConditionGroup "3521地物", "SSObj_Code", "==", "3521"
SSProcess.MergeNearNoteObj "控制点注记","3521地物",1.0,1,"SSObj_PointName(0)","",1
SSProcess.ClearSelectConditionGroups
End Sub
```


#### MergeNearPointObj 合并相近的两个点地物

**函数**：`MergeNearPointObj`

**功能**：合并相近的两个点地物。

**语法**：

```vbs
SSProcess.MergeNearPointObj(mainCondition, mergeCondition,searchRange,deleteMainObj)
```

**参数**：

- mainCondition 主地物选择条件分组名，由SetSelectConditionGroup函数设定
- mergeCondition 合并地物选择条件分组名，由SetSelectConditionGroup函数设定
- searchRange 以主地物为中心搜索合并地物的范围，以米为单位
- deleteMainObj 是否同时删除主地物，0（不删除）、1（删除）

**说明**：

执行该函数，将合并相近的两个点物的空间点列。

**示例**：

```vbs
下面示例把编码为8360点地物的空间点合并到与其相近的编码为8310的点地物,并删除8360地物。
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "8310地物", "SSObj_Code", "==", "8310"
SSProcess.SetSelectConditionGroup "3521地物", "SSObj_Code", "==", "3521"
SSProcess.MergeNearPointObj "8310地物","3521地物",1.0,1
SSProcess.ClearSelectConditionGroups
End Sub
```


#### MergePolygon 面对象合并

**函数**：`MergePolygon`

**功能**：面对象合并。

**语法**：

```vbs
SSProcess.MergePolygon(objIDs, limit, inheritAttrID, delSrcObj)
```

**参数**：

- objIDs 合并的面对象id集合，id之间用英文逗号隔开
- limit 合并限距
- inheritAttrID 多个面合并时，合并后的面扩展属性继承的id地物，为0时不继承任何属性
- delSrcObj 是否删除原面对象，0不删除；1删除

**说明**：

使用此函数将合并指定的面id对象。

**示例**：

```vbs
Sub OnClick()
objIDs="194,195"
limit=0.01
inheritAttrID="194"
delSrcObj=0
SSProcess.MergePolygon objIDs ,limit, inheritAttrID,delSrcObj
End Sub
```


#### ObjectDeal 对象处理器

**函数**：`ObjectDeal`

**功能**：对象处理器。

**语法**：

```vbs
SSProcess.ObjectDeal(objID, methodName, parameters, result)
```

**参数**：

- objID 处理对象ID
- methodName 方法名称
- parameters 参数列表
- result 返回值
- 常用方法列表
| 列1 | 列2 | 列3 | 列4 |
|---|---|---|---|
| 方法名称 | 方法参数 | 返回值 | 说明 |
| GotoPoints |  |  | 折线化指定对象点列 |
| Reverse |  |  | 反向指定对象点列 |
| ClearNodeSymbol | 要取舍地物的编码列表，多个时用逗号分隔 |  | 节点自动取舍 |
| SetSelectionObjectToNavigation |  |  | 选择集对象加入导航器 |
| AddToSelection |  |  | 指定 ID 对象加入到选择集 |
| FreeDisplayList |  |  | 释放指定 ID 的显示列表 |
| FreeSelectionObjectDisplayList |  |  | 释放选择集对象显示列表 |
| GetHouseFunctionDef |  |  | 获取房产功能区定义信息 |
| GetHouseApportionTree |  |  | 获取房产分摊树信息 |
| RepairAllHatchSymbol |  |  | 按旧算法重置面填充符号 |
| OrderAreaLoops | 0 ArcGIS 方式，外环顺 , 内环逆； 1 Oracle 方式，外环逆 , 内环。 |  | 重排面环顺逆时针方向 |

**示例**：

```vbs
Sub OnClick()
折线化
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "2120000,2150000"
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount()
For i=0 To geoCount-1
innerObjID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
SSProcess.ObjectDeal innerObjID, "GotoPoints", "", result
SSProcess.SetObjectAttr innerObjID, "SSObj_DataMark", "处理"
Next
'清空选择集
SSProcess.ClearSelection
'清空选择条件
SSProcess.ClearSelectCondition
'设置选择条件
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT"
SSProcess.SetSelectCondition "SSObj_Code", "=", GeoCode
SSProcess.SetSelectCondition "SSObj_DataMark", "=", "高匹配错误"
'执行选择过滤
SSProcess.SelectFilter
'将未匹配成功的添加到导航器
SSProcess.ObjectDeal 0, "SetSelectionObjectToNavigation", "", result
SSProcess.ClearSysSelection '清空系统选择集
SSProcess.ObjectDeal ID_JZW, "AddToSelection", "", result
SSProcess.UpdateSysSelection 1   '脚本选择集更新到系统选择集
SSProcess.ObjectDeal geoID, "FreeDisplayList", parameters, result
SSProcess.RefreshView
'重新符号化
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type","==","POINT"
SSProcess.SetSelectCondition "SSObj_Code","==","733300"
SSProcess.SelectFilter
SSProcess.ObjectDeal 0, "FreeSelectionObjectDisplayList", "", result
End Sub
```


#### OutsideUpdateRegionObjToSelection 选择更新区域外图形对象到选择集

**函数**：`OutsideUpdateRegionObjToSelection`

**功能**：选择更新区域外图形对象到选择集。

**语法**：

```vbs
SSProcess.OutsideUpdateRegionObjToSelection()
```

**说明**：

使用此函数将选择更新区域外图形对象到选择集。

**示例**：

```vbs
Sub OnClick()
SSProcess.OutsideUpdateRegionObjToSelection ()
SSProcess.UpdateSysSelection 1
End Sub
```


#### QuickFill 在指定坐标处构面

**函数**：`QuickFill`

**功能**：在指定坐标处构面。

**语法**：

```vbs
SSProcess.QuickFill(x, y, limitDist, areaCode, explodeCodes, notExplodeCodes)
```

**参数**：

- x 指定点x坐标
- y 指定点y坐标
- limitDist 距离指定坐标点的距离阀值
- areaCode 构面所用编码
- explodeCodes 进行构面的编码集合
- notExplodeCodes 不进行构面的编码集合

**说明**：

使用此函数在指定坐标处构面。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
geocount = SSProcess.GetSelGeoCount
For i = 0 To geocount -1
SSProcess.GetSelPolygonFocus i , x, y , z
SSProcess.QuickFill x ,y, 1,"2" ,explodeCodes,notExplodeCodes
Next
End Sub
```


#### RebuildSelectionTopRelation 重建选择集内拓扑面与弧段关系

**函数**：`RebuildSelectionTopRelation`

**功能**：重建选择集内拓扑面与弧段关系。

**语法**：

```vbs
SSProcess.RebuildSelectionTopRelation(areaLayer, arcLayer, limit)
```

**参数**：

- areaLayer 拓扑面所在层名
- arcLayer 弧段所在层名
- limit 限距大小

**说明**：

此函数功能为重建选择集内拓扑面与弧段关系。

**示例**：

```vbs
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
title="请选择拓扑面"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Function RebuildTopRelation( polygonID )
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_ID", "==", polygonID
SSProcess.SelectFilter
SSProcess.RebuildSelectionTopRelation "地类图斑" , "地类界线", 0.1
End Function
Sub OnOK()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.UpdateSysSelection 0
SSProcess.RebuildSelectionTopRelation "地类图斑" , "地类界线", 0.001
End Sub
```


#### RebuildTopRelation 重建拓扑面与弧段关系

**函数**：`RebuildTopRelation`

**功能**：重建拓扑面与弧段关系。

**语法**：

```vbs
SSProcess.RebuildTopRelation(areaCondition, lineCondition, limit)
```

**参数**：

- areaCondition 拓扑面条件
- lineCondition 弧段条件
- limit 限距大小。

**说明**：

此函数功能为重建拓扑面与弧段关系。

**示例**：

```vbs
Function RebuildOneType(areaLayer,lineLayer)
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "面", "SSObj_LayerName", "==" , areaLayer
SSProcess.SetSelectConditionGroup "面", "SSObj_Type", "==" , "AREA"
SSProcess.SetSelectConditionGroup "界线", "SSObj_LayerName", "==", lineLayer
SSProcess.SetSelectConditionGroup "界线", "SSObj_Type", "==", "LINE"
'先单元处理polygonID
SSProcess.RebuildTopRelation "面", "界线", 0.001
End Function
Sub OnClick()
RebuildOneType "地类图斑" , "地类界线"
RebuildOneType "宗地", "界址线"
RebuildOneType  "行政区划村界面" , "行政区划县界,行政区划乡界,行政区划村界"
RebuildOneType "基本农田保护块" , "基本农田保护块"
RebuildOneType "开发园区" , "开发园区"
RebuildOneType  "坡度图", "坡度图"
End Sub
```


#### RepairDGXHeight 修复选择集对象的非法高程值

**函数**：`RepairDGXHeight`

**功能**：修复选择集对象的非法高程值。

**语法**：

```vbs
SSProcess.RepairDGXHeight()
```

**说明**：

执行该函数，将自动修改选择对象中各点高程值与主高程值不一致的情况,认为0或大于等于9000的高程值为非法高程。

**示例**：

```vbs
下面示例修复8111编码对象的非法高程值：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "8111"
SSProcess.SelectFilter
SSProcess.RepairDGXHeight
End Sub
```


#### RepairExtentAttr 修复对象扩展属性

**函数**：`RepairExtentAttr`

**功能**：修复对象扩展属性。

**语法**：

```vbs
SSProcess.RepairExtentAttr()
```

**说明**：

执行此函数可修复对象扩展属性。

**示例**：

```vbs
Sub OnClick()
SSProcess.RepairExtentAttr()
End Sub
```


#### RepairDGXHeight 修复选择集对象的非法高程值

**函数**：`RepairGeoHeight`

**功能**：修复选择集对象的非法高程值。

**语法**：

```vbs
SSProcess.RepairGeoHeight(repairZero, repair9999, repairInvalid, maxHeightDIf)
```

**参数**：

- repairZero 是否修复0高程
- repair9999 是否修复 >=9000的无效高程
- repairInvalid 是否修复高差大于maxHeightDIf值的高程
- maxHeightDIf 最大高差

**说明**：

执行该函数，将自动修改选择对象中各点高程值与主高程值不一致的情况,认为0或大于等于9000的高程值为非法高程。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectCondition "SSObj_Type", "<>", "POINT,NOTE"
SSProcess.SelectFilter
repairZero = 1 '是否修复0高程
repair9999=1  '是否修复 >=9000的无效高程
repairInvalid=1  '是否修复高差大于maxHeightDIf值的高程
maxHeightDIf = 4 '最大高差
SSProcess.RepairGeoHeight repairZero, repair9999, repairInvalid, maxHeightDIf
End Sub
```


#### RepairSelectionObjPointList 修复选择集对象的非法空间点

**函数**：`RepairSelectionObjPointList`

**功能**：修复选择集对象的非法空间点。

**语法**：

```vbs
SSProcess.RepairSelectionObjPointList( miniDist)
```

**参数**：

- miniDist 相邻点允许最小距离,以米为单位

**说明**：

执行该函数，将自动修改选择对象中的非法空间点，包括回头线，相邻重复点。

**示例**：

```vbs
下面示例修复选择集空间点：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
SSProcess.RepairSelectionObjPointList 0.001
End Sub
```


#### RepairUpdateObject 修复对象更新特性

**函数**：`RepairUpdateObject`

**功能**：RepairUpdateObject 函数

**语法**：

```vbs
SSProcess.RepairUpdateObject()
```

**说明**：

执行此函数可修复对象扩展属性。

**示例**：

```vbs
Sub OnClick()
SSProcess.RepairExtentAttr
End Sub
```


#### SaveSourceData 保存原始数据

**函数**：`SaveSourceData`

**功能**：保存原始数据。

**语法**：

```vbs
SSProcess.SaveSourceData(beginID)
```

**参数**：

- beginID 数据源句柄

**说明**：

执行该函数，将保存原始数据。


#### SearchCityPartAttr 查找城市部件关联属性

**函数**：`SearchCityPartAttr`

**功能**：查找城市部件关联属性。

**语法**：

```vbs
SSProcess.SearchCityPartAttr(partCondition,partFieldName,roadCondition,roadFieldName,
```

**参数**：

- partCondition 部件选择条件组名称
- partFieldName 部件属性字段名
- roadCondition 道路选择条件组名称
- roadFieldName 道路属性字段名
- courtyardCondition 院落选择条件组名称
- courtyardFieldName 院落属性字段名
- interespointCondition 兴趣点选择条件组名称
- interespointFieldName 兴趣点属性字段名
- interespointCount 搜索兴趣点个数

**说明**：

执行该函数，将搜索与部件最近的道路、院落、兴趣点之间的关系，并把关系描述内容填充到相应的属性字段中。

**示例**：

```vbs
下面示例更新部件的[WZMS]属性字段
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "部件", "SSObj_Type", "==", "POINT"
SSProcess.SetSelectConditionGroup "部件", "SSObj_Code", ">", "100"
SSProcess.SetSelectConditionGroup "部件", "SSObj_Code", "<", "609"
SSProcess.SetSelectConditionGroup "道路", "SSObj_Code", "==", "940"
SSProcess.SetSelectConditionGroup "道路", "{OuterObjSearchRange}", "==", "200"
SSProcess.SetSelectConditionGroup "院落", "SSObj_Code", "==", "2280"
SSProcess.SetSelectConditionGroup "院落", "{OuterObjSearchRange}", "==", "200"
SSProcess.SetSelectConditionGroup "兴趣点", "SSObj_Code", "==", "943"
SSProcess.SetSelectConditionGroup "兴趣点", "{NearObjSearchRange}", "==", "200"
SSProcess.SearchCityPartAttr "部件","[WZMS]","道路","[Name]","院落","[街巷名称]","兴趣点","[WZMS]", 2
SSProcess.ClearSelectConditionGroups
End Sub
```


#### SearchInnerObjIDs 查找被指定地物包含的对象

**函数**：`SearchInnerObjIDs`

**功能**：查找被指定地物包含的对象。

**语法**：

```vbs
SSProcess.SearchInnerObjIDs( polygonID, objecType, codes, innerObjGetPointMode )
```

**参数**：

- polygonID 多边形对象ID
- objecType 搜索对象类型(0 点 1 线 2 面 3注记 10 所有)
- codes 搜索对象编码,有多个时,用,号分隔,为空时,则搜索所有编码
- innerObjGetPointMode 被包含对象取点方式, 0(判断所有点) 1(只判断焦点)

**说明**：

执行该函数，将搜索给定多边形内,指定搜索对象类型和编码的对象。
返回符合搜索条件的对象ID值,有多个时,用逗号分隔。

**示例**：

```vbs
下面示例搜索落在指定多边形内的地物,并修改其高程值
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SetSelectCondition "<Close>", "=, "1"
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount()
Dim arID(1000), idCount
For i=0 To geoCount-1
polygonID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
ids = SSProcess.SearchInnerObjIDs(polygonID, 0, "8310,3521", 0)
If ids <> "" Then
SSFunc.ScanString ids, ",", arID, idCount
For k=0 To idCount-1
SSProcess.SetObjectAttr CInt(arID(k)), "SSObj_Z", "123.44"
Next
End If
Next
End Sub
```


#### SearchInPolyObjIDs 查找与指定多边形相关的对象

**函数**：`SearchInPolyObjIDs`

**功能**：查找与指定多边形相关的对象。

**语法**：

```vbs
SSProcess.SearchInPolyObjIDs(polygonID, objecType, codes, innerObjGetPointMode, includeInner, includeCross)
```

**参数**：

- polygonID 指定的多边形ID
- objecType 搜索目标对象几何类型(0 点 1 线 2 面 3注记 10 所有)
- codes 搜索目标对象编码,有多个时,用,号分隔,为空时,则搜索所有编码
- innerObjGetPointMode 被包含对象取点方式, 0(判断所有点) 1(只判断焦点)
- includeInner 是否搜索完全落入的被包含对象, 0(不搜索) 1(搜索)
- includeCross 是否搜索相交对象, 0(不搜索) 1(搜索)

**说明**：

执行该函数，将搜索与指定多边形相关的目标编码的对象，包括落入多边形内部以及与指定多边形相交。
返回符合搜索条件的对象ID值,有多个时,用,号分隔,返回结果按面积从小到大排序


#### SearchNearObjIDs 查找与指定坐标点相近的地物

**函数**：`SearchNearObjIDs`

**功能**：查找与指定坐标点相近的地物。

**语法**：

```vbs
SSProcess.SearchNearObjIDs( x, y, searchRange, objecType, codes, filterID )
```

**参数**：

- x 查找坐标点X坐标
- y 查找坐标点Y坐标
- searchRange 搜索范围(米为单位)
- objecType 搜索对象类型(0 点 1 线 2 面 3注记 10 所有)
- codes 搜索对象编码,有多个时,用,号分隔,为空时,则搜索所有编码
- filterID 不包含在返回结果中的对象ID

**说明**：

执行该函数，将按给定的坐标点,搜索范围,搜索对象类型和编码搜索与坐标点相近的对象。
返回符合搜索条件的对象ID值,有多个时,用逗号分隔。

**示例**：

```vbs
下面示例搜索与指定地物节点相近的高程点
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type, "==","LIKE"
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
Dim arID(1000), idCount
geoCount = SSProcess.GetSelGeoCount()
For i = 0 To geoCount - 1
pointCount = SSProcess.GetSelGeoPointCount(i)
For j = 0 To pointCount - 1
SSProcess.GetSelGeoPoint i, j, x, y, z, pointType, name
ids = SSProcess.SearchNearObjIDs(x, y, 8.0, 0, "8310", 0 )
MsgBox ids
Next
Next
End Sub
```


#### SearchNearObjIDs1 查找与指定线相近的地物

**函数**：`SearchNearObjIDs2`

**功能**：查找与指定线相近的地物。

**语法**：

```vbs
SSProcess.SearchNearObjIDs2(lineID, searchRange, codes, filterID)
```

**参数**：

- lineID 线ID
- searchRange 搜索范围(米为单位)
- codes 搜索对象编码,有多个时,用,号分隔,为空时,则搜索所有编码
- filterID 不包含在返回结果中的对象ID

**说明**：

执行该函数，将查找与指定线相近的地物。

**示例**：

```vbs
下面示例查找与指定线相近的地物
Sub OnClick()
ids = SSProcess.SearchNearObjIDs2 (70, 1, "", 0 )
Msgbox ids
End Sub
```


#### SearchNearObjIDs2 查找与指定线相近的地物 000000

**函数**：`SearchNearObjIDs2`

**功能**：查找与指定线相近的地物。

**语法**：

```vbs
SSProcess.SearchNearObjIDs2(lineID, searchRange, codes, filterID)
```

**参数**：

- lineID 线ID
- searchRange 搜索范围(米为单位)
- codes 搜索对象编码,有多个时,用,号分隔,为空时,则搜索所有编码
- filterID 不包含在返回结果中的对象ID

**说明**：

执行该函数，将查找与指定线相近的地物。

**示例**：

```vbs
下面示例查找与指定线相近的地物
Sub OnClick()
ids = SSProcess.SearchNearObjIDs2 (70, 1, "", 0 )
Msgbox ids
End Sub
```


#### SearchObjIDs 对象查找

**函数**：`SearchObjIDs`

**功能**：查找与指定对象相交的对象。

**语法**：

```vbs
SSProcess.SearchObjIDs(xyCoords, objecType, codes, innerObjGetPointMode)
```

**参数**：

- xyCoords 线坐标串，"x1,y1,x2,y2"
- objecType 搜索对象类型(0 点 1 线 2 面 3注记 10 所有)
- codes 搜索对象编码,有多个时,用,号分隔,为空时,则搜索所有编码
- innerObjGetPointModecodes 被包含对象取点方式, 0(判断所有点) 1(只判断焦点)

**说明**：

执行该函数，将搜索与指定坐标串相交的符合指定被搜索对象类型和编码的对象。
返回值为符合搜索条件的对象ID值,多个时,用,号分隔

**示例**：

```vbs
Sub OnClick()
xyLineCoords = "433008.652,2887620.877,433088.625,2887598.821"
nCode="2"
strIDs = SSProcess.SearchObjIDs(xyLineCoords, 2, nCode, 1)
Msgbox strIDs
End Sub
```


#### SearchOnLineObjIDs 查找在线上的点对象

**函数**：`SearchOnLineObjIDs`

**功能**：查找在线上的点对象。

**语法**：

```vbs
SSProcess.SearchOnLineObjIDs(x1, y1, x2, y2, strWhere, OnlineLimit)
```

**参数**：

- x1 线第一点x 坐标
- y1 线第一点y 坐标
- x2 线第二点x 坐标
- y2 线第二点y 坐标
- strWhere 点对象集合
- OnlineLimit 线距

**说明**：

执行该函数，将查找在线上的点对象。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==" ,"1"
SSProcess.SelectFilter
geoCount= SSProcess.GetSelGeoCount()
For i=0 To geoCount - 1
objID= SSProcess.GetSelGeoValue( i ,"SSObj_ID" )
pcount= SSProcess.GetObjectAttr(objID ,"SSObj_PointCount" )
For j=0 To pcount -2
SSProcess.GetObjectPoint objID, j, x0 , y0, z,ptype , name
SSProcess.GetObjectPoint objID, j+1, x1, y1 , z1, ptype, name
strPids= SSProcess.SearchOnLineObjIDs(x0 ,y0,x1 ,y1, "1101021",0.001)
Msgbox strPids
Next
Next
End Sub
```


#### SearchOuterObjIDs 查找包含指定地物的对象

**函数**：`SearchOuterObjIDs`

**功能**：查找包含指定地物的对象。

**语法**：

```vbs
SSProcess.SearchOuterObjIDs( innerObjID, objecType, codes, innerObjGetPointMode )
```

**参数**：

- innerObjID 被包含对象ID
- objecType 搜索对象类型(0 点 1 线 2 面 3注记 10 所有)
- codes 搜索对象编码,有多个时,用,号分隔,为空时,则搜索所有编码
- innerObjGetPointMode 被包含对象取点方式, 0(判断所有点) 1(只判断焦点)

**说明**：

执行该函数，将搜索包含给定地物的多边形,指定搜索对象类型和编码的对象。
返回符合搜索条件的对象ID值,有多个时,用,号分隔,返回结果按面积从小到大排序。

**示例**：

```vbs
下面示例搜索包含指定地物的多边形内列表,并修改其高程值。
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount()
Dim arID(1000), idCount
For i=0 To geoCount-1
innerObjID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
ids = SSProcess.SearchOuterObjIDs(innerObjID , 2 "2", 0)
If ids <> "" Then
SSFunc.ScanString ids, ",", arID, idCount
For k=0 To idCount-1
SSProcess.SetObjectAttr CInt(arID(k)), "SSObj_Z", "123.44"
Next
End If
Next
End Sub
```


#### SearchRelatePolygonIDs 查找与指定面对象关联的面对象

**函数**：`SearchRelatePolygonIDs`

**功能**：查找与指定面对象关联的面对象。

**语法**：

```vbs
SSProcess.SearchRelatePolygonIDs(polygonID, codes)
```

**参数**：

- polygonID 指定面对象的ID
- codes 与指定面对象有关联的面地物特征编码

**说明**：

使用此函数查找与指定面对象关联的面对象。
返回值为关联面对象的ID。

**示例**：

```vbs
Sub OnClick()
A = SSProcess.SearchRelatePolygonIDs ( 144,2610004) '设定选择条件
End Sub
```


#### SelectionObjInnerInsertPoint 选择集对象等距内插点

**函数**：`SelectionObjInnerInsertPoint`

**功能**：选择集对象等距内插点。

**语法**：

```vbs
SSProcess.SelectionObjInnerInsertPoint( step )
```

**参数**：

- step 内插步距(米为单位)

**说明**：

执行该函数，将对当前选择集中点数大于1的线或面地物按指定的步距进行内插点。

**示例**：

```vbs
下面示例内插过程
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
'根据设定的选择条件,选择对象到选择集中
SSProcess.SelectFilter
'按3米步距进行内插
SSProcess.SelectionObjInnerInsertPoint 3.0
End Sub
```


#### SelectionObjMerge 选择集对象相邻面合并处理

**函数**：`SelectionObjMerge`

**功能**：选择集对象相邻面合并处理。

**语法**：

```vbs
SSProcess.SelectionObjMerge( anewAreaCode, newAreaLayer, miniDist, delOldObj )
```

**参数**：

- anewAreaCode 合并后新面编码,等于0时,使用合并前的面编码
- newAreaLayer 合并后新面放置图层,等于空时, 使用合并前的面图层
- miniDist 最小悬距,米为单位
- delOldObj 是否删除原地物, 0 不删除 1 删除

**说明**：

执行该函数，将对当前选择集中有共边的面进行合并处理。

**示例**：

```vbs
下面示例面合并过程
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "AREA"
SSProcess.SetSelectCondition "SSObj_LayerName", "==", "工矿面"
SSProcess.SelectFilter
SSProcess.SelectionObjMerge 2,"DEFAULT",0.001,0
End Sub
```


#### SelectionObjOrderby 选择集对象按指定规则排序

**函数**：`SelectionObjOrderby`

**功能**：选择集对象按指定规则排序。

**语法**：

```vbs
SSProcess.SelectionObjOrderby(orderbyMode, refpointx, refpointy)
```

**参数**：

- orderbyMode 排序方式
- refpointx
- refpointy


#### SelectionObjPartZ 选择集对象内插节点高程

**函数**：`SelectionObjPartZ`

**功能**：选择集对象内插节点高程。

**语法**：

```vbs
SSProcess.SelectionObjPartZ(step)
```

**参数**：

- step 内插步距(米为单位)

**说明**：

执行该函数，将对当前选择集中地物的无效高程点进行插值处理,要求地物至少保证起终点有有效高程值。

**示例**：

```vbs
下面示例内插过程
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
'根据设定的选择条件,选择对象到选择集中
SSProcess.SelectFilter
SSProcess.SelectionObjPartZ 2
End Sub
```


#### SelectionObjTopProcess 选择集对象拓扑处理

**函数**：`SelectionObjTopProcess`

**功能**：选择集对象拓扑处理。

**语法**：

```vbs
SSProcess.SelectionObjTopProcess( areaCode, areaLayer, miniDist, delArc )
```

**参数**：

- areaCode 生成面编码
- areaLayer 生成面放置图层
- miniDist 最小悬距,米为单位
- delArc 是否删除拓扑弧段 0 删除 1 保留

**说明**：

执行该函数，将根据当前选择集中的弧段拓扑构面。

**示例**：

```vbs
下面示例构面过程
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
SSProcess.SetSelectCondition "SSObj_LayerName", "==", "LAF"
SSProcess.SelectFilter
SSProcess.SelectionObjTopProcess 2, "Default", 0.001, 1
End Sub
```


#### SetObjectBinaryAttr 设置对象二进制属性值

**函数**：`SetObjectBinaryAttr`

**功能**：设置对象二进制属性值。

**语法**：

```vbs
SSProcess.SetObjectBinaryAttr(geoID, attrField, attrValue)
```

**参数**：

- geoID 对象ID
- attrField 属性字段名
- attrValue 属性值

**说明**：

该函数设置对象二进制属性值。


#### SetSelGeoBinaryAttr 设置选择对象二进制属性值

**函数**：`SetSelGeoBinaryAttr`

**功能**：设置选择对象二进制属性值。

**语法**：

```vbs
SSProcess.SetSelGeoBinaryAttr(index, attrField, attrValue)
```

**参数**：

- geoID 对象ID
- attrField 属性字段名
- attrValue 属性值


#### SplitPolygon 面对象分割

**函数**：`SplitPolygon`

**功能**：面对象分割。

**语法**：

```vbs
SSProcess.SplitPolygon(polygonID, splitLineIDs, delSrcObj)
```

**参数**：

- polygonID 面地物ID
- splitLineIDs 分割地物ID
- delSrcObj 原始地物是否删除：0 保留； 1 删除

**说明**：

执行该函数，对面地物用分割地物进行分割，返回分割后地物ID集合。

**示例**：

```vbs
Sub OnClick()
areaID= 15  '面 ID
lineID= 16  '分割线 ID
'面分割
sids = SSProcess.SplitPolygon (areaID , lineID, 0 )
End Sub
```


#### TopProcess 拓扑构面

**函数**：`TopProcess`

**功能**：拓扑构面。

**语法**：

```vbs
SSProcess.TopProcess(progressMsg)
```

**参数**：

- progressMsg 数据源句柄

**说明**：

执行该函数，进行拓扑构面。

**示例**：

```vbs
Sub OnClick()
'参数: 拓扑线编码，拓扑面条件。拓扑线编码 --多个间用逗号分割；拓扑面条件 -- 属性点编码 1,面编码1/ 属性点编码 2,面编码2
'Respy居民地
TPGM "3102005,3100002,3107005,3110025,3104005,3105005,3106005,3103005" ,"3102009,3102003/3107009,3107003/3110029,3110023/3103009,3103003/3104009,3104003/3105009,3105003/3106009,3106003"
End Sub
'拓扑构面
'参数: 拓扑线编码，拓扑面条件。拓扑线编码 --多个间用逗号分割；拓扑面条件 -- 属性点编码 1,面编码1/ 属性点编码 2,面编码2
Function TPGM(ArcCodes,AreaCondition )
SSProcess.PushUndoMark
SSProcess.ClearFunctionParameter
'悬挂点处理限距
SSProcess.AddFunctionParameter "limitdist=0.0001"
'拓扑弧段编码
SSProcess.AddFunctionParameter "SrcArcCodes="&ArcCodes
'删除源弧段
SSProcess.AddFunctionParameter "DelSrcArc=0"
'删除上次生成的重叠弧段
SSProcess.AddFunctionParameter "DelNewArc=0"
'删除上次生成的原拓扑面
SSProcess.AddFunctionParameter "DelOldTopArea=1"
'数据处理后是否存盘
SSProcess.AddFunctionParameter "SaveDB=1"
'是否生成拓扑面
SSProcess.AddFunctionParameter "CreateTopArea=1"
'拓扑面编码设置属性点编码 1,面编码1, 图层名称1/属性点编码 2,面编码2, 图层名称2
SSProcess.AddFunctionParameter "NewObject="& AreaCondition
'判断属性点重复的关键字
SSProcess.AddFunctionParameter "LabelKeyFields="
'生成拓扑弧段选择：
'0 不生成弧段
'1 生成统一编码弧段，编码由 UniqueArcCode指定
'2 生成弧段, 当有多种线状地物重叠时，按 ReserveArcOrder设置的编码顺序优先从前选取
'3 自动生成与其他弧段重叠的新弧段 , 按CreateOverlayArc 设置
SSProcess.AddFunctionParameter "CreateTopArc=0"
SSProcess.TopProcess "拓扑构面"
End Function
```


---

### 数据转换

#### ClearDataXParameter 清除数据转换参数列表

**函数**：`ClearDataXParameter`

**功能**：清除数据转换参数列表。

**语法**：

```vbs
SSProcess.ClearDataXParameter()
```

**说明**：

执行该函数，将清除由SetDataXParameter函数设置的转换参数列表。


#### ClearMapFrameSourceData 清除指定图幅更新前原始数据

**函数**：`ClearMapFrameSourceData`

**功能**：清除指定图幅更新前原始数据。

**语法**：

```vbs
SSProcess.ClearMapFrameSourceData(mapNumber)
```

**参数**：

- mapNumber 图幅号

**说明**：

该函数用于清除指定图幅更新前原始数据。

**示例**：

```vbs
Sub OnClick()
SSProcess. ClearMapFrameSourceData ( 432750.00 - 2887500.00 )
End Sub
```


#### CreateMapClipHandle 创建图幅切割句柄

**函数**：`CreateMapClipHandle`

**功能**：创建图幅切割句柄。

**语法**：

```vbs
SSProcess.CreateMapClipHandle ()
```

**说明**：

执行该函数，将创建图幅切割句柄。


#### CutRegionToEdb 裁剪指定区域数据到 EDB 文件

**函数**：`CutRegionToEdb`

**功能**：裁剪指定区域数据到EDB文件。

**语法**：

```vbs
SSProcess.CutRegionToEdb(regionID, exportEdbFileOrPath, cutParameters, exportNewEdbFiles)
```

**参数**：

- regionID 剪裁区域ID
- exportEdbFileOrPath 打开EDB文件或路径
- cutParameters 剪裁参数
- exportNewEdbFiles 输出的新EDB文件

**说明**：

此函数功能为剪裁指定区域数据到EDB文件。

**示例**：

```vbs
Sub OnClick()
'裁剪输出到本EDB
notCnotECodes="Codes_Const ="&"211000"&";NotClipHasTurnPointObj=1;Codes_Const_Symbol_P =1;" '不裁剪, 不打散的地物编码
notCnotECodes_Y = "Codes_Const_Symbol_Y=1;" '不裁剪, 不打散的地物包括所有Y类
notCnotECodes_E= "Codes_Const_Symbol_E =1;"'不裁剪, 不打散的地物包括所有E类
notCnotEObjNoRepeat = "NoRepeatObjInMulitEDB=1;" '不重复输出标志
notCdoECodes="Codes_ExplodeOnly =;"  '不裁剪, 只打散的地物编码
SelSourceData = "SourceData=1;" '2全部数据集或 1活动数据集
doEdoCCodes="Codes_ExplodeAndClip =;"   '先打散, 后裁剪的地物编码
doCdoECodes="Codes_ClipAndExplod =;" '先裁剪, 后打散的地物编码
outputEDBFlag = "OutputEDBFlag =0;"  '结果输出标志: 0 整体输出到本edb, 1 整体输出到一个新edb, 2 分幅(或分区)逐个输出新edb
ClipTypeInRange = "ClipTypeInRange =2;" '区域内再分割标志: 0 不分割, 1 按标准分幅分割, 2 按指定编码区划分割
RangeCode = "RangeCode =211000"
Parameters = notCnotECodes & notCnotECodes_Y & notCnotECodes_E & notCnotEObjNoRepeat & notCdoECodes & doEdoCCodes & doCdoECodes & SelSourceData & outputEDBFlag & ClipTypeInRange & RangeCode
result = SSProcess.CutRegionToEdb (0, "", parameters, paths)
End Sub
```


#### ExportData 执行数据输出

**函数**：`ExportData`

**功能**：执行数据输出。

**语法**：

```vbs
SSProcess.ExportData()
```

**说明**：

执行该函数，将根据SetDataXParameter设定的转换参数进行数据输出。
在数据导出到指定格式前,将会把转换数据加入到脚本的选择集中,并自动调用脚本回调函数BeforeExportData(),通过BeforeExportData函数可实现导出前的数据整理(如:属性修改,换码,换层,坐标换算等)。
如果转换参数设置TransCoordStatus为1,则在导出前,会执行回调函数TransCoordXYZ(ByVal planeSysName, ByVal heightSysName, ByVal reverse, ByVal transxy, ByVal transz, ByRef x, ByRef y, ByRef z)实施坐标转换处理。

**示例**：

```vbs
下面示例SHP格式输出：
Sub OnClick()
SSProcess.ClearDataXParameter
SSProcess.SetDataXParameter "DataType", "5"
SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_SHP"
SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_500"
SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_500"
SSProcess.SetDataXParameter "ExportPathName", ""
SSProcess.SetDataXParameter "DataBoundMode", "0"
SSProcess.SetDataXParameter "ExportLayerCount", "0"
SSProcess.SetDataXParameter "SymbolExplodeMode"，"1"
'图层对照关系
'图层对照关系个数
SSProcess.SetDataXParameter "LayerRelationCount", "2"
'EPS层名:客户方点层名:客户方线层名:客户方面层名:客户方点注记层名:客户方线注记层名
SSProcess.SetDataXParameter "LayerRelation1", "房屋面:房屋面点:房屋面线:房屋面:房屋面点注记:房屋面线注记"
SSProcess.SetDataXParameter "LayerRelation2", "工矿点:工矿点:工矿点线:工矿点面:工矿点点注记:工矿点线注记"
'字段输出定制
'字段定制个数
SSProcess.SetDataXParameter "TableFieldDefCount", "25"
'字段定义 层名,类型(0点,1线,2面,3注记),EPS字段名,客户字段名,系统字段名,缺省值,字段类型,字段长度,小数位
SSProcess.SetDataXParameter "TableFieldDef1", "房屋面,2,ID,ID,ID,,dbLong,0,0"
SSProcess.SetDataXParameter "TableFieldDef2", "房屋面,2,Code,要素代码,Code,,dbLong,0,0"
SSProcess.SetDataXParameter "TableFieldDef3", "房屋面,2,图形特征,图形特征,图形特征,,dbText,127,0"
SSProcess.SetDataXParameter "TableFieldDef10", "工矿点,0,ID,ID,ID,,dbLong,0,0"
SSProcess.SetDataXParameter "TableFieldDef11", "工矿点,0,Code,要素代码,Code,,dbLong,0,0"
SSProcess.SetDataXParameter "TableFieldDef12", "工矿点,0,图形特征,图形特征,图形特征,,dbText,127,0"
SSProcess.SetDataXParameter "TableFieldDef13", "工矿点,0,点号,点号,PointName,,dbText,127,0"
SSProcess.SetDataXParameter "TableFieldDef14", "工矿点,0,高程,高程,Z,,dbDouble,19,3"
SSProcess.SetDataXParameter "TableFieldDef20", "居民地注记,3,ID,ID,ID,,dbLong,0,0"
SSProcess.SetDataXParameter "TableFieldDef21", "居民地注记,3,分类号,要素代码,分类号,,dbText,10,0"
SSProcess.SetDataXParameter "TableFieldDef22", "居民地注记,3,图形特征,图形特征,图形特征,,dbText,127,0"
SSProcess.SetDataXParameter "TableFieldDef23", "居民地注记,3,注记内容,注记内容,注记内容,,dbText,127,0"
SSProcess.SetDataXParameter "TableFieldDef24", "居民地注记,3,字高,字高,字高,,dbInteger,7,0"
SSProcess.SetDataXParameter "TableFieldDef25", "居民地注记,3,字宽,字宽,字宽,,dbInteger,7,0"
SSProcess.ExportData
End Sub
下面示例高斯直角坐标输出为经纬度的ArcGIS的MDB格式
Function BeforeExportData()
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount-1
pointCount = SSProcess.GetSelGeoPointCount (i)
SSProcess.LockSelGeoPoint i, 1
For j=0 To pointCount-1
SSProcess.GetSelGeoPoint i, j, x, y, z, ptype, pname
SSProcess.xy80ToLongiLati 114, x, y, y0, x0
SSProcess.SetSelGeoPoint i, j, x0, y0, z, ptype, pname
Next
SSProcess.UpdateSelGeoPoint i
SSProcess.LockSelGeoPoint i, 0
Next
End Function
Sub OnClick()
'清空转换参数
SSProcess.ClearDataXParameter
'设置导入文件格式为ArcGIS PDB
SSProcess.SetDataXParameter "DataType", "22"
SSProcess.SetDataXParameter "AddSystemFieldMode", "0"
'开始导入数据
SSProcess.ExportData
End Sub
```


#### FreeExportObjBufferData 释放转换对象缓存

**函数**：`FreeExportObjBufferData`

**功能**：FreeExportObjBufferData 函数

**语法**：

```vbs
SSProcess.FreeExportObjBufferData()
```

**说明**：

执行该函数，将释放转换对象缓存。


#### FreeMapClipHandle 释放图幅切割句柄

**函数**：`FreeMapClipHandle`

**功能**：释放图幅切割句柄。

**语法**：

```vbs
SSProcess.FreeMapClipHandle()
```

**说明**：

执行该函数，将释放图幅切割句柄。

**示例**：

```vbs
Sub OnClick()
SSProcess.CreateMapFrameByRegion 999700
frameCount = SSProcess.GetMapFrameCount()
'创建图幅切割句柄
SSProcess.CreateMapClipHandle
For i=0 To frameCount- 1
SSProcess.GetMapFrameCenterPoint i , x, y
SSProcess.SetCurMapFrame x , y, 0, ""
frameID = SSProcess.GetCurMapFrame()
mapNumber = SSProcess.GetObjectAttr( CLng (frameID), "[MapNumber]" )
'判断图幅是否有更新前原始数据
If mapNumber <> ""Then
'裁剪输出到EDB
edbFileName = pathName & mapNumber & ".edb"
notCnotECodes ="Codes_Const=" & "211000,211300,211500,212000,213000,213300,214000,215000,216000,217000,218000,221100,221200, " & ";" '不裁剪, 不打散的地物编码
notCnotECodes_Y = "Codes_Const_Symbol_Y=1;" '不裁剪, 不打散的地物包括所有Y类notCnotECodes_E = "Codes_Const_Symbol_E=1;"'不裁剪, 不打散的地物包括所有E类
notCnotEObjNoRepeat = "NoRepeatObjInMulitEDB=1;" '不重复输出标志
notCdoECodes ="Codes_ExplodeOnly=;"   '不裁剪, 只打散的地物编码
SelSourceData = "SourceData=1;" '2全部数据集或 1活动数据集
doEdoCCodes ="Codes_ExplodeAndClip=;"  '先打散, 后裁剪的地物编码
parameters=notCnotECodes&notCnotECodes_Y&notCnotECodes_E&notCnotEObjNoRepeat&notCdoECodes&doEdoCCodes&doCdoECodes&SelSourceData
result = SSProcess.CutCurMapFrameToEdb(edbFileName,parameters)
End If
Next
'释放图幅切割句柄
SSProcess.FreeMapClipHandle
If fileName <> "" Then
Msgbox "输出完成！"
End If
End Sub
```


#### GetImportFileName 获取当前转换的文件名

**函数**：`GetImportFileName`

**功能**：获取当前转换的文件名。

**语法**：

```vbs
SSProcess.GetImportFileName()
```

**说明**：

执行该函数，将返回当前正在转换的数据文件名。
返回当前正在转换的文件名,包含全路径。

**示例**：

```vbs
下面示例调入DWG文件时,把数据源文件名设置为图形对象的数据标识：
'数据转换过程中,在保存图形对象到数据库前调入的函数
Function BeforeSaveImportData()
Dim filename, datamark
'获取转换文件名
filename = SSProcess.GetImportFileName()
index = InStrRev(filename, "", -1, 1 )
datamark = Right(filename, len(filename)-index )
count = SSProcess.GetSelGeoCount()
For i=0 To count-1
SSProcess.SetSelGeoValue i, "SSObj_DataMark", datamark
Next
count = SSProcess.GetSelNoteCount()
For i=0 To count-1
SSProcess.SetSelNoteValue i, "SSObj_DataMark", datamark
Next
End Function
Sub OnClick()
'清空转换参数
SSProcess.ClearDataXParameter
'设置导入文件格式为DWG
SSProcess.SetDataXParameter "DataType", "1"
'设置扩展属性存储在MemoData中
SSProcess.SetDataXParameter "SaveAttrToMemoData", "1"
'设置加载的图层
SSProcess.SetDataXParameter "ImportLayerList", "DGX,GCD"
'开始导入数据
SSProcess.ImportData
End Sub
```


#### ImportData 执行数据导入

**函数**：`ImportData`

**功能**：执行数据导入。

**语法**：

```vbs
SSProcess.ImportData()
```

**说明**：

执行该函数，将根据SetDataXParameter设定的转换参数进行数据导入。
在数据存入数据库前,将会把转换数据加入到脚本的选择集中,并自动调用脚本回调函数BeforeS aveImportData(),通过BeforeSaveImportData函数可实现数据存库前的数据整理(如:属性修改,换码,换层,坐标换算等)。
如果转换参数设置TransCoordStatus为1,则在存库前,会执行回调函数TransCoordXYZ( ByVal planeSysName, ByVal heightSysName, ByVal reverse, ByVal transxy, ByVal transz, ByRef x, ByRef y, ByRef z)实施坐标转换处理。

**示例**：

```vbs
下面示例DWG格式导入：
Sub OnClick() '清空转换参数
SSProcess.ClearDataXParameter '设置导入文件格式为DWG
SSProcess.SetDataXParameter "DataType", "1"
'设置扩展属性存储在MemoData中
SSProcess.SetDataXParameter "SaveAttrToMemoData", "1"
'开始导入数据
SSProcess.ImportData
' 设置Undo标志
SSProcess.PushUndoMark
'清空选择集
SSProcess.ClearSelection
'清空选择条件
SSProcess.ClearSelectCondition
'设置选择条件为只选择点对象
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT"
'执行选择过滤
SSProcess.SelectFilter
'选择集对象进行编码匹配
SSProcess.UpdateObjAttrByFeatureCode "FeatureCodeTB_Dxf", "('#'+Feature.Byname)='SSObj_Name' and Feature.Type=0 ", "SSObj_LayerName ="
Feature.LayerName,SSObj_Code = Feature.Code,SSObj_Color="
'清空选择集
SSProcess.ClearSelection
'清空选择条件
SSProcess.ClearSelectCondition
'设置选择条件为只选择点对象
SSProcess.SetSelectCondition "SSObj_Type", "=", "LINE"
'执行选择过滤
SSProcess.SelectFilter
'选择集对象进行编码匹配
SSProcess.UpdateObjAttrByFeatureCode "FeatureCodeTB_Dxf", "Feature.Byname='[AcadThickness]' and Feature.Type<>0", "SSObj_LayerName = Feature.LayerName,SSObj_Code = Feature.Code,SSObj_Color=Feature.LineColor,SSObj_Reverse=Feature.Reverse,SSObj_Name=,SSObj_MemoData="
'清空选择集
SSProcess.ClearSelection
'清空选择条件
SSProcess.ClearSelectCondition
End Sub
下面示例经纬度ArcGIS的MDB格式数据导入，并把经纬度换算为高斯直角坐标：
Function BeforeSaveImportData()
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount-1
pointCount = SSProcess.GetSelGeoPointCount (i)
SSProcess.LockSelGeoPoint i, 1
For j=0 To pointCount-1
SSProcess.GetSelGeoPoint i, j, x, y, z, ptype, pname
SSProcess.LongiLatiToxy80 114, y, x, x0, y0
SSProcess.SetSelGeoPoint i, j, x0, y0, z, ptype, pname
Next
SSProcess.UpdateSelGeoPoint i
SSProcess.LockSelGeoPoint i, 0
Next
End Function
Sub OnClick()
'清空转换参数
SSProcess.ClearDataXParameter
'设置导入文件格式为ArcGIS PDB
SSProcess.SetDataXParameter "DataType", "22"
SSProcess.SetDataXParameter "SaveAttrToMemoData", "1"
'开始导入数据
SSProcess.ImportData
End Sub
```


#### IsMapFrameExistentSourceData 指定图幅是否存在更新前原始数据

**函数**：`IsMapFrameExistentSourceData`

**功能**：指定图幅是否存在更新前原始数据。

**语法**：

```vbs
SSProcess.IsMapFrameExistentSourceData(mapNumber)
```

**参数**：

- mapNumber 图幅号
- 返回值：0 不存在，1 存在

**说明**：

执行该函数，判断指定图幅是否存在更新前原始数据。

**示例**：

```vbs
Sub OnClick()
BZ = SSProcess.IsMapFrameExistentSourceData ("432750.00 - 2887500.00" )
Msgbox BZ
End Sub
```


#### SetDataXParameter 设定数据转换参数

**函数**：`SetDataXParameter`

**功能**：设定数据转换参数。

**语法**：

```vbs
SSProcess.SetDataXParameter( parameterName, parameterValue )
```

**参数**：

- parameterName 转换参数名称
- 转换参数定义见下表：
| 参数 | 说明 |
|---|---|
| DataType | 数据格式格式。 0(ArcGIS SDE) 、 1(DWG) 、 2(DXF) 、 3(E00) 、 4(Coverage) 、 5(Shp) 、 6(Tab) 、 7(MIf) 、 8(MGE Txt) 、 9(Dgn) 、 10(WtWlWp) 、 11(WatWalWap) 、 12(Xyz) 、 13(Vtr) 、 14(Ebf) 、 15(Svf) 、 16(Cas) 、 17(Wex) 、 18(Eps Exf) 、 19(Eps Exx) 、 20(Eps   Mdb) 、 21(Eps Edb) 、 22(ArcGIS Mdb) 、 23( 国标 Vct) 、 24(Eps Cor/Not) 、 25 (Eps Cef) 、 26( 国土农村 VCT) 、 27( 国土城镇 VCT) 、 28( 上海 MDO) 、 29( 上航天远景 XML) 。 |
| Version | 版本号 : AutoCad 数据版本号。 12,14,2000,2004,2006 Exf 1(1.0) 2(2.0) |
| FeatureCodeTBName | 输出使用编码对照表名称。 |
| SymbolScriptTBName | 输出使用符号描述表名称。 |
| NoteTemplateTBName | 输出使用注记分类表名称。 |
| ExportPathName | 输出文件名 ( 或者路径名 ), 如果为空时 , 则自动弹出对话框选择 , 如果为 {GetFromNewEpsIni}, 则从 NewEps.INI 文件的 [DataX] 节的 ExportPathName 变量获取。 |
| ImportPathName | 导入文件名 ( 或者路径名 ), 如果为空时 , 则自动弹出对话框选择。 |
| FileNameFormat | 分层输出时文件名格式。如果为空用层名为文件名 , 用 * 代表层名 , 用 & 代表当前图幅号 , 用 @ 代表当前工程名 , 可加前后缀 , 如 A901* 或 &* 或 @* 。 |
| SaveObjectToDatabase | 导入数据时，图形数据存储方式。 0 （不保存到数据库）， 1 （保存到数据库）。 |
| SaveAttrToMemoData | 导入数据时，扩展属性存储方式。 0 （扩展属性存贮在属性表中）， 1 （扩展属性存贮在 MemoData 中）。 |
| ImportNotMatchAttrToMemoData | 导入数据时，无匹配扩展属性存储方式。 0 （丢弃无匹配属性）， 1 （无匹配属性存储到 MemoData 中）。 |
| CreateAttrTable | 导入数据时，是否自动创建属性表结构。 0 （不创建） 1 （创建）。 |
| DataBoundMode | 数据输出范围方式， 0( 所有数据 ) ， 1( 选择集数据 ) ， 2( 当前图幅 ) ， 3( 缓冲区 ) ， 4( 指定编码闭合地物 ) ， 5( 指定 ID 闭合地物 ) ， 6( 所有图幅 ) 。 |
| DataBoundBuffer | 数据输出范围缓冲区名称。只有在 DataBoundMode 为 3 时起作用。 |
| DataBoundCode | 数据输出范围地物编码。只有在 DataBoundMode 为 4 时起作用。 |
| DataBoundID | 数据输出范围地物 ID 。只有在 DataBoundMode 为 5 时起作用。 |
| ZoomInOutDataBound | 数据输出范围缩放量，米为单位，缺省值为 -0.0001 。 |
| UseUpdateDataStatus | 数据转换是否支持数据更新模式 , 0 不支持 1 支持。 如果设为支持状态 , 则在数据调入时 , 自动存储更新前数据 , 数据输出时 , 自动按新增 , 修改 , 删除和未改四个文件输出。 |
| UseUpdateDataMode | 数据转换按更新模式处理时 , 原数据存储方式 , 0 上海分幅方式 1 非分幅方式。 |
| ExportClipObjToOneFile | 数据转换按更新模式处理时 , 跨图幅对象只输出一次 , 0 重复输出 1 只输出到一个图幅。 |
| ExportLayerMode | 输出图层的方式 ,0( 一个图层分点线面四个文件输出 ),1( 一个图层输出一个文件 ) 。 |
| ExportLayerCount | 输出图层个数。如果等于 0 ，则只输出当前打开的图层。 |
| ExportLayer[Index] | 输出图层名称。 [Index] 为图层索引序号，从 1 开始。 |
| LayerRelationCount | 输出层名对照关系个数。 |
| LayerRelation[Index] | 输出层名对照关系。 格式： "EPS 层名 : 客户方点层名 [, 别名 ]: 客户方线层名 [, 别名 ]: 客户方面层名 [, 别名 ]: 客户方点注记层名 [, 别名 ]: 客户方线注记层名 [, 别名 ]" ，如果一个图层输出一个文件 , 则对照关系可以是 "EPS 层名 : 客户方层名 " ， [Index] 为对照关系索引序号，从 1 开始。 如果客户方层名需指定数据集名称时，可在客户方层名前加数据集名称，用点分隔 如：水系 . 水系面。 格式中的 [] 表示可选项 , 实际写参数时不写 [] 符号。 |
| TableRelationCount | 输出属性表名与层名类型的对照关系个数。 |
| TableRelation[Index] | 输出属性表名与层名类型的对照关系。 格式 1 ： " 属性表名 : 层名 : 类型 (0 点 ,1 线 ,2 面 ,3 注记 ,10 点线面共用 )" ， [Index] 为对照关系索引序号，从 1 开始 格式 2 ： " 客户属性表名 1, 客户属性表名 2->EPS 属性表名 : 层名 : 类型 (0 点 ,1 线 ,2 面 ,3 注记 ,10 点线面共用 )" ， [Index] 为对照关系索引序号，从 1 开始 |
| UserLayerInfoCount | EPS 临时用户层定义个数。 |
| UserLayerInfo[Index] | EPS 临时用户层定义。该用户层定义只供数据输出时获取挂接属性表名使用。 格式：用户层名 , 点属性表名 , 线属性表名 , 面属性表名 , 注记属性表名 |
| TableFieldDefCount | 属性字段输出定义个数。 |
| TableFieldDef[Index] | 属性字段输出定义。 格式 1 ： " 层名 , 类型 (0 点 ,1 线 ,2 面 ,3 注记 ,10 点线面共用 ),EPS 字段名 , 客户字段名 ,[ 客户字段别名 ,] 系统字段名 , 缺省值 , 字段类型 , 字段长度 , 小数位 " 。 格式 2 ： " 表名 ,EPS 字段名 , 客户字段名 ,[ 客户字段别名 ,] 系统字段名 , 缺省值 , 字段类型 , 字段长度 , 小数位 " 。 其中“字段类型 ” 包括以下类型： (dbBoolean,dbByte,dbInteger,dbLong,dbSingle,dbDouble,dbDate,dbText,dbMemo,dbLongBinary,dbGUID) “缺省值 ” 除了设置缺省值外 , 同时支持设置值对照 , 优先于值对照设置 , 值对照格式如下： ValRelate=EPS 值 1: 客户值 1;EPS 值 2: 客户值 2;... 如果要转出 ACAD 扩展属性时 , 需在客户字段别名位置按如下格式定义： AcadAppName:AcadGroupCode[:AcadFieldName:AcadFieldNameGroupCode] 如果指定了 AcadFieldName, 则会在输出字段值的同时输出字段名称。 如果 AcadAppName 值等于 DefineSystemField 时 , 该对照字段不输出 , 但该对照生成的值可用于替换系统字段名为 Sys_AllField 所关联的字段。 格式中的 [] 表示可选项 , 实际写参数时 [] 符号不要。 " 系统字段名 " 包括以下字段： ID 对象 ID LayerName 对象层名 DateTime 修改时间 CreateTime 创建时间 Byname 对象别名 Code 对象编码 ObjectName 对象名称 Thickness 对象厚度 Angle 对象角度 ( 度为单位 , 正北 0 度 逆时针 ) Angle_Radian 对象角度 ( 弧度为单位 , 正北 0 度 逆时针 ) GeoAngle 对象方位角 ( 度为单位 , 正北 0 度 顺时针 ) GeoAngle_Radian 对象方位角 ( 弧度为单位 , 正北 0 度 顺时针 ) ArithAngle 对象角度 ( 度为单位 , 正东 0 度 逆时针 ) ArithAngle_Radian 对象角度 ( 弧度为单位 , 正东 0 度 逆时针 ) Length 对象长度 Length3D 对象 3D 长度 Perimeter 对象周长 Area 对象面积 EllipsoidArea 椭球面积 FrameEllipsoidArea 图幅椭球面积 DataMark 数据标识 Name 名称 LayerName 层名 MapNumber 图幅号 X 首点 X 坐标 Y 首点 Y 坐标 Z 首点 Z 坐标 PointType 首点点类型信息 PointName 首点点名 X(i) 指定点 X 坐标 (i 为 End 时 , 认为是尾点 ) Y(i) 指定点 Y 坐标 Z(i) 指定点 Z 坐标 PointType(i) 指定点点类型 PointName(i) 指定点点名 LineColor 对象颜色 (RGB 值 ) LineColor_RGB 对象颜色 (R,G,B) LineColor_Index 对象颜色 ( 颜色号 ) FillColor 对象填充颜色 (RGB 值 ) FillColor_RGB 对象填充颜色 (R,G,B) FillColor_Index 对象填充颜色 ( 颜色号 ) LineWidth 对象线宽 ( 图上 0.01mm) LineWidth_mm 对象线宽 ( 图上 1mm) 图形特征 图形符号化信息描述 Esri_2DShape          ArcGIS 2DShape Esri_3DShape          ArcGIS 3DShape Esri_ObjectID         ArcGIS ObjectID GruphicGUID 图形全球唯一 ID StringText 注记文本内容 FontName 字体名 FontColor 注记颜色 (RGB 值 ) FontColor_RGB 注记颜色 (R,G,B) FontColor_Index 注记颜色 ( 颜色号 ) FontWeight 字体磅数 Alignment 对齐方式 Underline 是否下划线 IlaticAngle 斜体角度 DownAngle 耸肩角度 WordAngle 字角度 ( 度为单位 ) WordAngle_Radian 字角度 ( 弧度为单位 ) StringAngle 字串角度 ( 度为单位 ) StringAngle_Radian 字串角度 ( 弧度为单位 ) FontHeight 字高 ( 图上 0.01mm) FontHeight_mm 字高 ( 图上 1mm) FontHeight_m 字高 ( 实地 m) FontWidth 字宽 ( 图上 0.01mm) FontWidth_mm 字宽 ( 图上 1mm) FontWidth_m 字宽 ( 实地 m) FontInterval 字间隔 ( 图上 0.01mm) FontInterval_m 字间隔 ( 图上 1mm) FontInterval_m 字间隔 ( 实地 m) Direction 字头朝向 FontClass 注记分类号 FontByname 注记字体别名 FontStyle 注记形状 (0\|1\|2\|3\|4\|5 -- 0 代表正体 ,1 代表左斜 ,2 代表右斜 ,3 代表左耸 ,4 代表右耸 ) FontSize 字号 (1p=0.353mm) AttrType 地物或注记属性类型 AutomaticID 自动 ID AcadHandle            AcadDWG 文件中对象的 Handle Sys_AllField 所有字段信息 , EPS 字段名可以同时填写多个 , 用 [] 括住每个字段 , 如 [ 点名 ]-[ 高程 ] 由系统替换字段值 , 如果字段间的要使用西文逗号时 , 需要用 $ ， 代替西文逗号 ToNullFilter 为空时不输出该字段项 |
| NotCheckUpdateTableFieldCount | 不检测更新属性字段定义个数 |
| NotCheckUpdateTableField[Index] | 不检测更新属性字段定义。 格式 : 层名 , 类型 (0 点 ,1 线 ,2 面 ,3 注记 ,10 点线面共用 ), 字段 1, 字段 2,... |
| AddSystemFieldMode | 输出数据时 , 是否强制输出系统字段 ,   0( 不输出 ) 1( 输出 ) 。 |
| ExportAttrMode | 输出数据时 , 扩展属性输出方式 ,   0( 优先输出外部信息 , 无外部信息时则输出属性表属性 ) 1 ( 只输出属性表属性 ) 2 ( 只输出外部信息 ) 。 |
| TransCheckInfo | 转换数据时 , 是否转换检查排除和核查信息， 0( 不转换 ) 1( 转换 ) 。 |
| ExportPZValue | 是否输出点高程值 , 0( 不输出 ) 1( 输出 ) 。 |
| ExportLZValue | 是否输出线高程值 , 0( 不输出 ) 1( 输出 ) 。 |
| ExportAZValue | 是否输出面高程值 , 0( 不输出 ) 1( 输出 ) 。 |
| ExportNZValue | 是否输出注记高程值 , 0( 不输出 ) 1( 输出 ) 。 |
| ZeroLineWidth | 输出 AutoCAD 数据时， 0 线宽分界值，小于或等于该值的线宽，输出时均设为 0 。 |
| FontWidthScale | 输出注记字宽缩放比， FontClass_ 分类号 = 缩放比 , 如果直接填写缩放比 , 则默认为全局缩放比 , 有多个分类号时 , 用逗号分隔 ( 如 FontClass_0=0.6,FontClass_1=0.7), 缩放比取值范围 0-1 。 |
| FontHeightScale | 输出注记字高缩放比， FontClass_ 分类号 = 缩放比 , 如果直接填写缩放比 , 则默认为全局缩放比 , 有多个分类号时 , 用逗号分隔 ( 如 FontClass_0=0.6,FontClass_1=0.7), 缩放比取值范围 0-1 。 |
| FontPositionOffsetX | 输出注记 X 方向移位 ,FontClass_ 分类号 = 中心对齐移位量 : 左上对齐移位量 : 左下对齐移位量 : 左中对齐移位量 : 右上对齐移位量 : 右中对齐移位量 : 右下对齐移位量 : 上中对齐移位量 : 下中对齐移位量 如果直接填写移位量 , 则默认为全局移位量 , 有多个分类号时 , 用逗号分隔 ( 如 FontClass_0=0.1:0.1:0.1:0.1:0.1:0.1:0.1:0.1:0.1,FontClass_1=0.1:0.1:0.1:0.1:0.1:0.1:0.1:0.1:0.1), 移位量以字宽为参考基准 , 取值范围 (0-1) |
| FontPositionOffsetY | 输出注记 X 方向移位 ,FontClass_ 分类号 = 中心对齐移位量 : 左上对齐移位量 : 左下对齐移位量 : 左中对齐移位量 : 右上对齐移位量 : 右中对齐移位量 : 右下对齐移位量 : 上中对齐移位量 : 下中对齐移位量 如果直接填写移位量 , 则默认为全局移位量 , 有多个分类号时 , 用逗号分隔 ( 如 FontClass_0=0.1:0.1:0.1:0.1:0.1:0.1:0.1:0.1:0.1,FontClass_1=0.1:0.1:0.1:0.1:0.1:0.1:0.1:0.1:0.1), 移位量以字宽为参考基准 , 取值范围 (0-1) 。 |
| FontSizeUseStatus | 字体大小使用状态 0 （按注记分类表设置字高宽输出）、 1 （按注记设置字高宽输出）。 |
| UsePersonalFontSizeClasss | 使用原字体宽高注记分类列表 , 多个时用逗号分隔。 |
| ChangeArcNoteToHashAlignment | 转换弧段注记为散列注记 0 （不转换）、 1 （转换）。 |
| ChangeNoteAlignment | 转换注记的对齐方式 FontClass_ 分类号 = 如果直接填写对齐方式，则为默认的对齐方式，未指定的分类号均使用该对齐方式 对齐方式，包括 0 （中心）、 1 （左上角）、 2 （左下角）、 3 （左中）、 4 （右上角）、 5 （右中）、 6 （右下角）、 7 （上中）、 8 （下中） |
| GotoPointsMode | 输出图形折线化方式。 0 （不折线化）、 1 （只折线化曲线）、 2 （所有图形折线化）。 |
| TransPAreaToSingleRing | 带岛面转成单环面状态。 0 （不转）、 1 （转换）。 |
| ExplodeObjLayerStatus | 内嵌符号图层输出方式。 0 （按符号描述设定输出）、 1 （与主地物同层输出）。 |
| ExplodeObjColorStatus | 内嵌符号颜色输出方式。 0 （按符号描述设定输出）、 1 （与主地物同色输出）。 |
| ExplodeNoteStatus | 注记打散状态。 0 （不打散）、 1 （打散到单点注记） 、 2 （打散到图形）。 |
| TransBeelineNoteToPointNote | 直线注记转为带字隔的单点注记状态。 0 （不转）、 1 （转换）。 |
| ExplodeNoteToGraphicUseCode | 注记打散到图形使用编码。 格式 1: 注记分类号 1= 打散图形编码 1, 注记分类号 2= 打散图形编码 2 格式 2: 打散图形编码 |
| ExplodeNoteToPointClass | 打散到点注记的注记分类号列表，逗号分隔，如果为空时则打散所有注记，该参数只有 ExplodeNoteStatus= |
| ExplodeObjMakeGroup | 输出 AutoCAD 数据时，打散对象编组输出方式。 0 （不编组）、 1 （ 编组，同时要求 FeatureCodeTB 表中的 ExtraInfo=1 ） |
| AcadLinetypeGenerateMode | 输出 AutoCAD 数据时，线型生成是否启用。 0 （禁用） 1 （启用） |
| AcadLinetypeGenerateException | 输出 AutoCAD 数据时，线型生成是否启用例外线型名，多个时用逗号分隔 |
| Acad 图案输出设置 | 在 FeatureCodeTB 的 Reference 字段设置相应参数 Reference 填充格式 : 图案名称 , 图案类型 , 图案风格 , 填充比例 , 填充角度 , 填充间隔 , 是否双向填充 其中 : 1. Reference 为空时 , 填充样式为 SOLID, 每项值均为非必填项 2. 图案类型可选值 UserDefined PreDefined CustomDefined 3. 图案风格可选值 Normal Outer Ignore 4. 填充角度以度为单位 5. 是否双向填充填 0 或 1 |
| AcadAttributeDefCount | 输出 AutoCAD 数据时，图块属性定义个数 |
| AcadAttributeDef[Index] | 输出 AutoCAD 数据时，图块属性定义 格式： 块名 ,EPS 字段名 , 系统字段名 , 小数位 ,ACAD 标记名 ,ACAD 标记提示 , 默认值 , 不可见 , 固定 , 验证 , 预置 , 文字样式 , 对齐方式 , 高度 , 宽度比例 , 旋转 , 倾斜角度 , 图层 , 颜色 ,X 偏移 ,Y 偏移 小数位 =-1 不处理小数位 >-1 为处理小数位位数 对齐方式 包括 0 （中心）、 1 （左上角）、 2 （左下角）、 3 （左中）、 4 （右上角）、 5 （右中）、 6 （右下角）、 7 （上中）、 8 （下中）。 |
| ColorUseStatus | 颜色使用状态。 0 （按编码表设定颜色输出）、 1 （按地物设定颜色输出）。 |
| UsePersonalColorCodes | 使用原色编码列表 , 多个时用逗号分隔。 |
| UsePersonalColorClasss | 使用原色注记分类列表 , 多个时用逗号分隔。 |
| LayerUseStatus | 数据输出层名使用状态。 0 （按编码表设定层名输出）、 1 （按地物设定层名输出）。 |
| ImportLayerUseStatus | 数据导入层名使用状态。 0 （由系统根据数据格式自动设定的图层调入）、 1 （按地物设定层名 调入）。 |
| UsePersonalLayerCodes | 使用原层名编码列表 , 多个时用逗号分隔。 |
| UsePersonalLayerClasss | 使用原层名注记分类列表 , 多个时用逗号分隔。 |
| LineWidthUseStatus | 线宽使用状态。 0 （按编码表设定线宽输出）、 1 （按地物设定线宽输出） |
| LineExportMode | 输出 AutoCAD 数据时，多义线输出方式， 0 （缺省方式，带不同高程时按 3DPolyline 输出，其余按 2DPolyline 输出）、 1 （强制按 2DPolyline 输出）、 2 （ 强制按 3DPolyline 输出） 3 （ 强制按 Polyline 输出）。 |
| ExportTo2DPolylineCodes | 强制按 2DPolyline 输出的编码列表 , 用逗号分隔。 |
| ExportTo3DPolylineCodes | 强制按 3DPolyline 输出的编码列表 , 用逗号分隔。 |
| ExportToPolylineCodes | 强制按 Polyline 输出的编码列表 , 用逗号分隔。 |
| TransArcToPolyline | 是否把圆弧转换成多义线输出 , 0 （不转换）、 1 （转换） |
| TransCircleToPolyline | 是否把圆转换成多义线输出 , 0 （不转换）、 1 （转换） |
| ThicknessExportMode | 输出 AutoCAD 数据时，厚度输出方式。 0 （地物编码）、 1 （编码表中的厚度）、 2 （编码表中的别名）、 3 （置成 0 ） |
| ThicknessExportToZFactor | 输出 AutoCAD 数据时，厚度输出到块 Z 比例方式。 0 （不输出）、 1 （输出） |
| SymbolExplodeMode | 符号打散方式。 0 （自动打散）、 1 （根据编码表设定打散）、 2 （全部不打散） |
| AcadLinFileName | 自动加载 Acad 线型全路径文件名。 |
| AcadLineScaleMode | Acad 线型比例输出方式。 0 与比例尺成正比输出 1 总是按 1 输出。 |
| AcadUsePersonalLineScaleCodes | 指定使用特殊线形比例的编码。 格式 1 ： 比例 1= 编码 1, 编码 2; 比例 2= 编码 1, 编码 2 格式 2 ： 比例 ( 该方式指定所有编码均使用指定的线型比例 ) |
| AcadBlockScaleMode | Acad 块比例输出方式。 0 与比例尺成正比输出 1 总是按 1 输出 |
| AcadUsePersonalBlockScaleCodes | 指定使用特殊块比例的编码。 格式 1 ： 比例 1= 编码 1, 编码 2; 比例 2= 编码 1, 编码 2 格式 2 ： 比例 ( 该方式指定所有编码均使用指定的块比例 ) |
| AcadDwtFileName | 输出 DWG 模板文件名。 |
| AcadColorMethod | 输出 DWG 颜色使用方式 0 （颜色号）、 1 （ RGB 颜色值） |
| AcadSetRandomAngleBlockNames | 输出 DWG, 设置随机旋转角度的块名列表 , 用逗号分隔。 |
| ExportAcadAttrToMdb | 输出 ACAD 数据时 , 属性数据是否输出到 MDB 文件 , 0 不输出 1 新建文件输出 2 使用模板文件输出 ( 通过 AcadAttrDBTemplateName 指定路径名 ) 。 |
| AcadAttrDBTemplateName | ACAD 属性输出模板文件名。 |
| AcadBlockUseColorMode | Acad 图块输出颜色使用方式。 0 随层 1 随块 2 随块内实体。 |
| AcadLineWeightMode | Acad 线重输出方式。 0 地物线宽 1 随层 2 随块 3 随线定义 |
| AcadLineWidthMode | Acad 线宽输出方式。 0 不输出 1 输出 |
| AcadExportObjMode | 输出 ACAD 数据时 , 对象输出方式 0 Acad 标准对象 1 Eps 自定义对象 |
| AcadExportObjRotateByFrame | 输出 ACAD 数据时 , 对象是否随图廓旋转 0 不随图廓旋转 1 随图廓旋转 |
| AcadExportNotRotateObjCodes | 输出 ACAD 数据时 , 不随图廓旋转地物编码列表 , 用逗号分隔 , 该参数只有在 AcadExportObjRotateByFrame 设为 1 时才有效。 |
| AcadExportNotRotateNoteCodes | 输出 ACAD 数据时 , 不随图廓旋转注记编码列表 , 用逗号分隔 , 该参数只有在 AcadExportObjRotateByFrame 设为 1 时才有效。 |
| AcadMlineCodes | 输出 ACAD 数据时 , 按多线输出编码列表 格式 : EPS 编码 1= 对正方式 , 缺省比例 ;EPS 编码 2= 对正方式 , 缺省比例 ;... 说明 : 多线样式填写在编码表的 Byname 字段 , 如果是变宽线 , 则把宽度值作为多线比例 , 否则使用指定缺省值 , 对正方式 0 顶端 1 零 2 底端 |
| HasInnerPointAreaCodes | 面内含带方向内嵌点的面编码列表 , 逗号分隔 , 当该类面被图幅切割时 , 自动往每个切割面内复制一个内嵌点。 |
| TransCoordStatus | 在导入或输出数据时，是否进行坐标换算处理。 0 （不换算）、 1 （换算） |
| PlaneSysName | 平面坐标系名称。 |
| HeightSysName | 高程基准名称。 |
| TransCoordReverse | 正反向换算方式。 0 （正向）、 1 （反向） |
| TransCoordXY | 是否进行平面换算。 0 （不进行）、 1 （进行） |
| TransCoordZ | 是否进行高程换算。 0 （不进行）、 1 （进行） |
| TransCoordScriptMode | 坐标转换脚本获取方式。 0 （从当前脚本中获取） 1 （从加密锁中获取） |
| XYCoordScale | 数据转换时， XY 坐标缩放比例。 |
| ZCoordScale | 数据转换时， Z 坐标缩放比例。 |
| XCoordOffset | 数据转换时， X( 横坐标 ) 偏移量。 |
| YCoordOffset | 数据转换时， Y( 纵坐标 ) 偏移量。 |
| ZCoordOffset | 数据转换时， Z 偏移量。 |
| ExportMapInfoPointSymbolSize | MapInfo 点符号大小。 |
| ExportMapInfoCharset | MapInfo 使用字符集 , 如 WindowsSimpChinese 。 |
| ExportMapInfoDelimiter | MapInfo 使用分隔符。 |
| ExportMapInfoProjectiondef | MapInfo 投影系名称 , 如 (CoordSys   NonEarth 或 CoordSys Earth   Projection 1) 。 |
| ExportMapInfoUnits | MapInfo 使用单位 ,(cm,ch,ft,in,km,li,m,mi,mm,nmi,rd,survey   ft,yd) 。 |
| ExportCodeMode | 编码输出方式 ,0 取 CODE 值 ,1 取 BYNAME 值 |
| ImportLayerList | 加载图层名列表，为空时，加载所有层，有多个图层名时，用逗号分隔，如（ DGX,DCD,KZD ） |
| UnImportLayerList | 不加载图层名列表，有多个图层名时，用逗号分隔，如（ DGX,DCD,KZD ） , 如果设置了 ImportLayerList 参数 , 该参数无效 |
| ImportUnknownBlock | DWG 调入时，匿名块调入方式。 0 （丢弃）、 1 （保留）、 1 （打散） |
| ExplodeBlockNames | DWG 调入时，需要打散的块名列表，多个时用逗号分隔。 |
| ExplodeInsertAtZeroBlock | DWG 调入时，是否打散插入坐标为 0,0 的块。 0 （不打散）、 1 （打散） |
| ExportArcInfoNoteToPoint | ArcInfo 的 E00 和 Coverage 的注记是否转成点输出 (0 不转成点 1 转成点 ) |
| ExportArcInfoPrecision | ArcInfo 的 E00 和 Coverage 的输出精度 (1 单精度 2 双精度 ) 。 |
| ExportArcInfoAnnoMode | ArcInfo 的 E00 和 Coverage 的注记输出方式 (1 TXT 2 TX6 3 TX7) 在输出注记时 , 将从注记的扩展属性数据中的 Anno_SubClass 字段获取 SubClass 信息 , 从 Anno_Level 字段获取 Level 信息 , 从 Anno_Symbol 字段获取 Symbol 信息 , 因此需定义注记属性的字段结构 . 如 : SSProcess.SetDataXParameter   "TableFieldDef1",   "anno,3,Anno_SubClass,Anno_SubClass,Byname,,dbText,10,0" SSProcess.SetDataXParameter   "TableFieldDef2",   "anno,3,Anno_Level,Anno_Level,AttrType,,dbText,10,0" SSProcess.SetDataXParameter   "TableFieldDef3", "anno,3,Anno_Symbol,Anno_Symbol,FontByname,,dbText,10,0" |
| ExportArcInfoTolInfo | ArcInfo 的 E00 和 Coverage 的允限值设置 格式： index1,flag1,value1;index2,flag2,value2;.... |
| EXCHANGE_CAS_CIRCLETO2PCODE | CAS 数据转换中的 2 点圆编码列表 ( 用 , 号分隔 ) 。 |
| EXCHANGE_CAS_HEIGHTNOTEPOINTCODE | CAS 数据转换中的高程点 , 控制点 , 水深点编码 ( 用 , 号分隔 ) 。 |
| EXCHANGE_CAS_JZXCODE | CAS 数据转换中的界址线编码 ( 用 , 号分隔 ) 。 |
| EXCHANGE_CAS_SPLINECODE | CAS 数据转换中的样条曲线编码 ( 用 , 号分隔 ) 。 |
| EXCHANGE_CAS_TAIJIECODE | CAS 数据转换中的台阶编码 ( 用 , 号分隔 ) 。 |
| EXCHANGE_CAS_XIEPOCODE | CAS 数据转换中的斜坡编码 ( 用 , 号分隔 ) 。 |
| EXCHANGE_WEX_ESYMBOLCODE | WEX 数据转换中的斜坡楼梯台阶编码 , 填写 WEX 中 SymbolCode( 用 , 号分隔 ) 。 |
| EXCHANGE_SHP_PrjFile | SHP 数据转换中的图层坐标系定义文件 , 如 c:\ArcGIS\Bin\Beijing 1954.prj 。 |
| EXCHANGE_PDB_PrjFile | PDB 数据转换中的图层坐标系定义文件 , 如 c:\ArcGIS\Bin\Beijing 1954.prj 。 |
| EXCHANGE_PDB_SpatialRF_MinX | PDB 数据转换中的图层空间范围左下角 X 坐标。 |
| EXCHANGE_PDB_SpatialRF_MinY | PDB 数据转换中的图层空间范围左下角 Y 坐标。 |
| EXCHANGE_PDB_SpatialRF_MaxX | PDB 数据转换中的图层空间范围右上角 X 坐标。 |
| EXCHANGE_PDB_SpatialRF_MaxY | PDB 数据转换中的图层空间范围右上角 Y 坐标。 |
| EXCHANGE_PDB_SpatialRF_XYResolution | PDB 数据转换中的图层空间 XY 坐标精度值 , 如 0.0001 。 |
| EXCHANGE_PDB_SpatialRF_XYTolerance | PDB 数据转换中的图层空间 XY 坐标容差值 , 如 0.0002 。 |
| EXCHANGE_PDB_SpatialRF_MinZ | PDB 数据转换中的图层最小 Z 坐标值。 |
| EXCHANGE_PDB_SpatialRF_MaxZ | PDB 数据转换中的图层最大 Z 坐标值。 |
| EXCHANGE_PDB_SpatialRF_ZResolution | PDB 数据转换中的图层空间 Z 坐标精度值 , 如 0.0001 。 |
| EXCHANGE_PDB_SpatialRF_ZTolerance | PDB 数据转换中的图层空间 Z 坐标容差值 , 如 0.0002 。 |
| EXCHANGE_PDB_SpatialRF_FalseX | PDB 数据转换中的图层空间 FalseX 值。 |
| EXCHANGE_PDB_SpatialRF_FalseY | PDB 数据转换中的图层空间 FalseY 值。 |
| EXCHANGE_PDB_SpatialRF_FalseZ | PDB 数据转换中的图层空间 FalseZ 值。 |
| EXCHANGE_PDB_SpatialRF_xyUnits | PDB 数据转换中的图层空间 xyUnits 值 , 如 1000 代表 Resolution 为 0.001 。 |
| EXCHANGE_PDB_SpatialRF_zUnits | PDB 数据转换中的图层空间 zUnits 值 , 如 1000 代表 Resolution 为 0.001 。 |
| EXCHANGE_PDB_SpatialRF_mUnits | PDB 数据转换中的图层空间 zUnits 值 , 如 1000 代表 Resolution 为 0.001 。 |
| EXCHANGE_PDB_SpatialRF_GridSize | PDB 数据转换中的图层空间索引格值。 |
| EXCHANGE_PDB_SpatialRF_MinM | PDB 数据转换中的图层最小 M 值。 |
| EXCHANGE_PDB_SpatialRF_MaxM | PDB 数据转换中的图层最大 M 值。 |
| EXCHANGE_PDB_SpatialRF_MResolution | PDB 数据转换中的图层空间 M 精度值 , 如 0.0001 。 |
| EXCHANGE_PDB_SpatialRF_MTolerance | PDB 数据转换中的图层空间 M 容差值 , 如 0.0002 。 |
| EXCHANGE_PDB_Unit | PDB 数据转换中的图层单位 : esriUnknownUnits = 0, esriInches = 1, esriPoints = 2, esrIfeet = 3, esriYards = 4, esriMiles = 5, esriNauticalMiles = 6, esriMillimeters = 7, esriCentimeters = 8, esriMeters = 9, esriKilometers = 10, esriDecimalDegrees = 11, esriDecimeters = 12, esriUnitsLast = 13 |
| EXCHANGE_PDB_ExportNoteMode | PDB 数据转换中的注记输出模式 ,0( 直接转换为 ArcGIS 的 Annotation) 1( 所有注记转换为线方式 )   2( 所有注记转换为点方式 ) 3( 注记分别按点和线方式转换 ) 。 |
| EXCHANGE_PDB_TemplateFileName | 输出 PDB 数据模板文件名。 |
| EXCHANGE_PDB_ExportEmptyLayer | 输出 PDB 数据时是否输出空层 ,   0 不输出 1 输出 |
| EXCHANGE_PDB_ExportShapeMode | 输出 PDB 数据时 Shape 输出方式 ( 脚本未指定 Shape 字段时有效 ), 0 2D 1 3D |
| EXCHANGE_VCT_DataMark | VCT 数据转换中的文件头 DataMark 值。 |
| EXCHANGE_VCT_Unit | VCT 数据转换中的文件头 Unit 值 : 坐标单位 ,K 表示公里 ,M 表示米 ,D 表示以度为单位的经纬度 ,S 表示以度分秒表示的经纬度 ( 此时坐标格式为 DDDMMSS.SSSS, DDD 为度 , MM 为分 ,   SS.SSSS 为秒 ) 。 |
| EXCHANGE_VCT_Dim | VCT 数据转换中的文件头 Dim 值 : 坐标维数。 2 表示仅有二维坐标 ,3 表示有三维坐标。三维时 , 无论 Unit 如何定义 , 高程坐标单位均用米。 |
| EXCHANGE_VCT_Coordinate | VCT 数据转换中的文件头 Coordinate 值 : 坐标单位。 |
| EXCHANGE_VCT_Dim | VCT 数据转换中的文件头 Dim 值 : 坐标维数。 2 表示仅有二维坐标 ,3 表示有三维坐标。三维时 , 无论 Unit 如何定义 , 高程坐标单位均用米。 |
| EXCHANGE_VCT_Dim | VCT 数据转换中的文件头 Dim 值 : 坐标维数。 2 表示仅有二维坐标 ,3 表示有三维坐标。三维时 , 无论 Unit 如何定义 , 高程坐标单位均用米。 |
| EXCHANGE_VCT_Topo | VCT 数据转换中的文件头的 Topo 值 : 是否带结点与线段的拓扑关系。 2 表示有结点关联线目标的标识以及线目标有起结点、终结点左多边形、右多边形的拓扑信息 ,   1 则表示没有这些信息但有多边形关联的线目标标识码。 0 表示没有拓扑，多边形直接带坐标。 |
| EXCHANGE_VCT_Projection | VCT 数据转换中的文件头的 Projection 值 : 坐标投影名称。 |
| EXCHANGE_VCT_Spheroid | VCT 数据转换中的文件头的 Spheroid 值 : 参考椭球体。 |
| EXCHANGE_VCT_Separator | VCT 数据转换中的文件头的 Separator 值 : 任意单字节非空白字符 , 用做属性字段分隔符。基本部分，缺省为半角字符逗号“ ,” 。 |
| EXCHANGE_VCT_Date | VCT 数据转换中的文件头的 Date 值：外业调查完成的日期。 |
| EXCHANGE_VCT_FontSizeMode | VCT 数据转换中的字体大小输出方式 : 0 ( 字高 , 字宽 )   1 ( 字号 ) 。 |
| EXCHANGE_MapGIS_CoordScale | MapGIS 数据转换坐标比例系数。 |
| EXCHANGE_MapGIS_GraphicScale | MapGIS 数据转换图形参数比例系数。 |
| MapGIS 转换参数 | 在 FeatureCode 表的 Byname 和 ExtraInfo 字段设置地物转换参数 点对象设置 : Byname( 子图号 ) ExtraInfo( 高度 , 宽度 , 辅助色 , 图层 , 覆盖方式 [, 主颜色 ]) 线对象设置 : Byname( 线型号 [- 辅助线型号 ])   ExtraInfo(X 系数 ,Y 系数 , 辅助色 , 图层 , 覆盖方式 [, 主颜色 ]) 面对象设置 : Byname( 填充图案号 [- 边界颜色 - 边界线型号 - 边界辅助线型号 ]) ExtraInfo( 图案高 , 图案宽 , 图案颜色 , 图案所在图层 , 覆盖方式 , 边界 X 系数 , 边界 Y 系数 , 边界辅助色 , 边界图层 , 边界覆盖方式 [, 主颜色 ]) 在 Notetemplate 表的 FontName 和 FontByname 字段设置注记转换参数 注记对象设置 : FontName( 中文字体名 西文字体名 ) FontByname( 字形 , 图层 , 覆盖方式 [, 主颜色 ]) |
| EXCHANGE_EDB_EncryptCoordMode | EDB 数据转换时坐标加解密方式 : 0 系统默认加解密方式 1 强制使用通用版加解密方式 |
| EXCHANGE_EDB_LoadAttrMode | EDB 数据转换时属性缓存方式 : 0 缓存到属性表 1 缓存到外部信息 |
| EXCHANGE_MDB_LoadAttrMode | MDB 数据转换时属性缓存方式 : 0 缓存到属性表 1 缓存到外部信息 |
| EXCHANGE_MGE_FileMode | MGE 数据转换文件方式 : 0 TXT 和 DXF 1 TXT 和 DGN |
| EXCHANGE_EPSMDB_TemplateName | EPSMDB 转换时 , 指定 MDB 使用的模板名称 , 不含路径时 , 从当前执行目录下的 Templates 目录下查找。 |
| EXCHANGE_EPSEDB_TemplateName | EPSMDB 转换时 , 指定 EDB 使用的模板名称 , 不含路径时 , 从当前执行目录下的 Templates 目录下查找。 |
| EXCHANGE_DGN_SeedFile | DGN 转换使用种子文件名。 |
| EXCHANGE_DGN_ExportAttrToMdb | DGN 转换时 , 同时输出属性到 Access   MDB 文件。 |
| EXCHANGE_DGN_AttrDBTemplateName | DGN 转换时 , 属性输出 Access   MDB 模板文件名。 |
| EXCHANGE_DGN_LineStyleFile | DGN 转换时 , 线型对照表 , 格式 : 线型 ID, 线型名。 |
| EXCHANGE_DGN_ExportAreaMode | DGN 转换时 , 面地物输出方式 0 输出成 Shape  1 输出成 LineString |
| EXCHANGE_DGN_MaxLineStringPointCount | DGN 转换时 , 线串最多点数 , 缺省值 101 。 |
| EXCHANGE_DGN_MSLinkEntityNum | DGN 转换时 ,MSLink 的 EntityNum 值。 |
| EXCHANGE_DGN_ExportLineScale | DGN 转换时 , 是否输出线型比例 0 不输出 1 输出 |
| EXCHANGE_DGN_LineScale | DGN 转换时 , 线型比例值。 |
| EXCHANGE_DGN_ExportCellScale | DGN 转换时 , 是否输出单元比例 0 不输出 1 输出 |
| EXCHANGE_DGN_CellScaleX | DGN 转换时 , 单元 X 比例值。 |
| EXCHANGE_DGN_CellScaleY | DGN 转换时 , 单元 Y 比例值。 |
| EXCHANGE_DGN_ExportCellHeadMode | DGN 转换时 , 单元头信息输出方式 0 不输出 1 按编码表设置输出 |
| EXCHANGE_DGN_ExplodeHasIntervalNote | DGN 转换时 , 打散有字隔的单点注记方式 0 不打散 1 打散字隔大于 20 的单点注记 |
| EXCHANGE_DGN_ImportCellMode | DGN 转换时 , 有名称单元调入方式 0 打散方式调入 1 按块方式调入 |
| EXCHANGE_DGN_ImportAnonymousCellMode | DGN 转换时 , 匿名单元调入方式 0 打散方式调入 1 按块方式调入 |
| EXCHANGE_DGN_ImportCellUseLayerMode | DGN 转换时 , 单元调入使用图层方式 0 使用单元头层名 1 使用内部图形层名 |
| EXCHANGE_DGN_UseReverseImportModeCellName | DGN 转换时 , 单元调入使用反方式调入的单元名列表 , 用逗号分隔。 |
| EXCHANGE_DGN_CloneOnePartCellObj | DGN 转换时 , 对单部件图元自动复制成双部件图元 ( 适应 ARCMAP 显示 DGN 数据存在的 BUG)   0 不复制 1 复制 |
| EXCHANGE_DGN_ExplodeNoteRelation | DGN 转换时 , 打散注记是否输出关联关系 0 不输出 1 输出 |
| EXCHANGE_XYZ_ESymbolCodes | XYZ 转换时 ,E 类符号适普编码列表 ( 用逗号分隔 ) 。 |
| EXCHANGE_XYZ_YSymbolCodes | XYZ 转换时 ,Y 类符号适普编码列表 ( 用逗号分隔 ) 。 |
| EXCHANGE_XYZ_ParallelLineCodes | XYZ 转换时 , 双线平行线适普编码列表 ( 用逗号分隔 ) 如果平行线编码与主线不一致时 , 可在主码后指定平行线编码 , 用冒号分隔 ( 如果平行线编码与主线编码一致时可不设置 ) |
| EXCHANGE_XYZ_VarWidthLineCodes | XYZ 转换时 , 变宽线适普编码列表 ( 用逗号分隔 ) 。 |
| EXCHANGE_XYZ_NotVarWidthLineCodes | XYZ 转换时 , 非变宽线适普编码列表 ( 用逗号分隔 ) 。 |
| EXCHANGE_XYZ_2PLineCodes | XYZ 转换时 ,2 点定向线适普编码列表 ( 用逗号分隔 ) 。 |
| EXCHANGE_READ_ZDHFieldName | Read 转换时 ,EPS 宗地号字段名称。 |
| EXCHANGE_READ_DLHFieldName | Read 转换时 ,EPS 地类号字段名称。 |
| EXCHANGE_READ_LCFieldName | Read 转换时 ,EPS 楼层字段名称。 |
| EXCHANGE_READ_JGFieldName | Read 转换时 ,EPS 结构代码字段名称。 |
| EXCHANGE_READ_ZDCodes | Read 转换时 ,Read 宗地编码。 |
| EXCHANGE_READ_TBCodes | Read 转换时 ,Read 图斑编码。 |
| EXCHANGE_READ_FWCodes | Read 转换时 ,Read 房屋编码。 |
| EXCHANGE_MDO_ExportNoteMode | MDO 数据转换中的注记输出模式 ,0( 直接转换为 Mdo 的 Text) 1( 所有注记转换为线方式 ) |
| EXCHANGE_MDO_UpdateRegionCode | MDO 数据标准更新区域编码。 |
| EXCHANGE_MDO_UpdateRegionDownLoaderField | MDO 数据标准更新区域中下载人员字段名称。 |
| EXCHANGE_MDO_UpdateRegionDownLoadTimeField | MDO 数据标准更新区域中下载时间字段名称。 |
| EXCHANGE_MDO_UpdateRegionNameField | MDO 数据标准更新区域中区域名称字段名称。 |
- parameterValue 转换参数值

**说明**：

该函数，执行一次，设定一个参数，允许多次执行以设定多个参数。该函数并不实际进行数据转换操作，只有在执行数据输出ExportData或ImportData函数后，才会真正的进行数据转换。

**示例**：

```vbs
下面示例AutoCad的DWG格式输出：
Sub OnClick()
SSProcess.ClearDataXParameter
SSProcess.SetDataXParameter "DataType", "1"
SSProcess.SetDataXParameter "Version", "2000"
SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_500"
SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_500"
SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_500"
SSProcess.SetDataXParameter "ExportPathName", ""
SSProcess.SetDataXParameter "DataBoundMode", "2"
SSProcess.SetDataXParameter "ExportLayerCount", "0"
SSProcess.SetDataXParameter "ZeroLineWidth", "15"
SSProcess.SetDataXParameter "FontWidthScale", "0.7,FontClass_12643=”
SSProcess.SetDataXParameter "FontHeightScale", "0.7,FontClass_12643=”
SSProcess.SetDataXParameter "ThicknessExportMode", "0"
SSProcess.SetDataXParameter "ExplodeObjMakeGroup", "1"
SSProcess.SetDataXParameter "SymbolExplodeMode", "1"
SSProcess.SetDataXParameter "LineExportMode", "1"
SSProcess.SetDataXParameter "AcadLinFileName", "C:\Documents and Settings\lange\Application Data\Autodesk\AutoCAD 2004\R16.0\chs\Support\acadiso.lin"
SSProcess.ExportData
End Sub
下面示例AutoCad的DWG格式导入：
Sub OnClick()
'清空转换参数
SSProcess.ClearDataXParameter
'设置导入文件格式为DWG
SSProcess.SetDataXParameter "DataType", "1"
'设置扩展属性存储在MemoData中
SSProcess.SetDataXParameter "SaveAttrToMemoData", "1"
'开始导入数据
SSProcess.ImportData
' 设置Undo标志
SSProcess.PushUndoMark
'清空选择集
SSProcess.ClearSelection
'清空选择条件
SSProcess.ClearSelectCondition
'设置选择条件为只选择点对象
SSProcess.SetSelectCondition "SSObj_Type", "=", "POINT"
'执行选择过滤
SSProcess.SelectFilter
'选择集对象进行编码匹配
SSProcess.UpdateObjAttrByFeatureCode "FeatureCodeTB_Dxf", "('#'+Feature.Byname)='SSObj_Name' and Feature.Type=0 ", "SSObj_LayerName = Feature.LayerName,SSObj_Code = Feature.Code,SSObj_Color=Feature.LineColor,SSObj_Reverse=Feature.Reverse,SSObj_Name=,SSObj_MemoData="
'清空选择集
SSProcess.ClearSelection
'清空选择条件
SSProcess.ClearSelectCondition
'设置选择条件为只选择点对象
SSProcess.SetSelectCondition "SSObj_Type", "=", "LINE"
'执行选择过滤
SSProcess.SelectFilter
'选择集对象进行编码匹配
SSProcess.UpdateObjAttrByFeatureCode "FeatureCodeTB_Dxf", "Feature.Byname='[AcadThickness]' and Feature.Type<>0", "SSObj_LayerName = Feature.LayerName,SSObj_Code = Feature.Code,SSObj_Color=Feature.LineColor,SSObj_Reverse=Feature.Reverse,SSObj_Name=,SSObj_MemoData="
'清空选择集
SSProcess.ClearSelection
'清空选择条件
SSProcess.ClearSelectCondition
End Sub
```


#### UpdateExportObjListToBuffer 生成转换对象缓存

**函数**：`UpdateExportObjListToBuffer`

**功能**：生成转换对象缓存。

**参数**：

- symbolExplode 符号打散


---

### 图幅操作

#### CreateMapFrame 创建当前数据涉及图幅缓存

**函数**：`CreateMapFrame`

**功能**：创建当前数据涉及图幅缓存。

**语法**：

```vbs
SSProcess.CreateMapFrame()
```

**说明**：

执行该函数，将创建有数据图幅的图幅缓存。
通过方法 GetMapFrameCount 可获取缓存图幅的个数。
通过方法 GetMapFramePoint 可获取指定缓存图幅的坐标。
使用完成后需执行FreeMapFrame释放图幅缓存数据。

**示例**：

```vbs
批量分幅输出DWG数据：
Function ExportDwg( fileName )
'清空转换参数
SSProcess.ClearDataXParameter
SSProcess.SetDataXParameter "DataType", "1"
SSProcess.SetDataXParameter "Version", "2000"
SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_500"
SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_500"
SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_500"
SSProcess.SetDataXParameter "ExportPathName", filename
SSProcess.SetDataXParameter "DataBoundMode", "2"
SSProcess.SetDataXParameter "ExportLayerCount", "0"
SSProcess.SetDataXParameter "ZeroLineWidth", "15"
SSProcess.SetDataXParameter "FontWidthScale", "0.7,FontClass_12643=0.7"
SSProcess.SetDataXParameter "FontHeightScale", "0.7,FontClass_12643=0.7"
SSProcess.SetDataXParameter "ThicknessExportMode", "0"
SSProcess.SetDataXParameter "ExplodeObjMakeGroup", "1"
SSProcess.SetDataXParameter "SymbolExplodeMode", "1"
SSProcess.SetDataXParameter "LineExportMode", "1"
SSProcess.SetDataXParameter "AcadLinFileName", "C:\Documents and Settings\lange\Application Data\Autodesk\AutoCAD 2004\R16.0\chs\Support\acadiso.lin"
SSProcess.ExportData
End Function
Sub OnClick()
pathName = SSProcess.SelectPathName()
If pathName = "" Then
Exit Sub
End If
SSProcess.CreateMapFrame
frameCount = SSProcess.GetMapFrameCount()
For i=0 To frameCount-1
SSProcess.GetMapFrameCenterPoint i, x, y
SSProcess.SetCurMapFrame x, y, 0, ""
frameID = SSProcess.GetCurMapFrame()
mapNumber = SSProcess.GetObjectAttr( CLng(frameID), "[MapNumber]")
If mapNumber <> "" Then
fileName = pathName & mapNumber & ".DWG"
ExportDwg filename
End If
Next
SSProcess.FreeMapFrame
End Sub
```


#### CreateMapFrameByCenterLine 创建指定编码中心线的图幅缓存

**函数**：`CreateMapFrameByCenterLine`

**功能**：创建指定编码中心线的图幅缓存。

**语法**：

```vbs
SSProcess.CreateMapFrameByCenterLine( centerLineCode, paperWidth, paperHeight, marginDist )
```

**参数**：

- centerLineCode 要生成图幅缓存的中心线编码
- paperWidth 图纸有效打印范围宽度,米为单位
- paperHeight 图纸有效打印范围高度,米为单位
- marginDist 图幅间接边重叠区域宽度,米为单位

**说明**：

执行该函数，将创建指定编码中心线的图幅缓存, 图幅将沿指定中心线按斜分幅方式布设。
通过方法 GetMapFrameCount 可获取缓存图幅的个数。
通过方法 GetMapFramePoint 可获取指定缓存图幅的坐标。
使用完成后需执行FreeMapFrame释放图幅缓存数据。

**示例**：

```vbs
自动生成带状图幅：
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.CreateMapFrameByCenterLine 1, 250, 200, 0
frameCount = SSProcess.GetMapFrameCount()
For i=0 To frameCount-1
SSProcess.CreateOneMapFrame i, 2110
Next
SSProcess.FreeMapFrame
End Sub
```


#### CreateMapFrameByPrintRegion 创建指定打印页面方案的图幅缓存

**函数**：`CreateMapFrameByPrintRegion`

**功能**：创建指定打印页面方案的图幅缓存。

**语法**：

```vbs
SSProcess.CreateMapFrameByPrintRegion(pageScheme)
```

**参数**：

- pageScheme


#### CreateMapFrameByRegion 创建指定编码区域内的图幅缓存

**函数**：`CreateMapFrameByRegion`

**功能**：创建指定地物编码区域内的图幅缓存。

**语法**：

```vbs
SSProcess.CreateMapFrameByRegion( code )
```

**参数**：

- code 要生图幅缓存的区域编码

**说明**：

执行该函数，将创建创建指定编码区域内的图幅缓存。
通过方法 GetMapFrameCount 可获取缓存图幅的个数。
通过方法 GetMapFramePoint 可获取指定缓存图幅的坐标。
使用完成后需执行FreeMapFrame释放图幅缓存数据。

**示例**：

```vbs
自动生成测区接合表
Sub OnClick()
SSProcess.CreateMapFrameByRegion 1
frameCount = SSProcess.GetMapFrameCount()
For i=0 To frameCount-1
SSProcess.GetMapFrameCenterPoint i, x, y
SSProcess.SetCurMapFrame x, y, 1, "1:500图幅接合表"
Next
SSProcess.FreeMapFrame
End Sub
```


#### CreateMapFrameByRegionID 创建指定 ID 区域内的图幅缓存

**函数**：`CreateMapFrameByRegionID`

**功能**：创建指定ID区域内的图幅缓存。

**语法**：

```vbs
SSProcess.CreateMapFrameByRegionID(objID)
```

**参数**：

- objID 对象ID

**说明**：

执行此函数将创建指定ID区域内的图幅缓存。

**示例**：

```vbs
Sub OnClick()
GetTFHS
End Sub
Function GetTFHS()
SSParameter.SetParameterSTR "国土" , "获取图幅号字符串", ""
SSParameter.GetParameterSTR "国土" , "获取图幅号地物ID", "", keyValue1
If keyValue1="" Then Exit Function
If isnumeric(keyValue1)=False Then Exit Function
SSProcess.FreeMapFrame
SSProcess.CreateMapFrameByRegionID keyValue1
count=SSProcess.GetMapFrameCount
mapnumbers=""
For i=0 To count-1
SSProcess.GetMapFramePoint i , 0, x0, y0
SSProcess.GetMapFramePoint i , 2, x1, y1
x =(x0+x1)/2 : y=(y0 +y1)/ 2
mapnumber =SSProcess.GetMapFrameNumber (x, y )
If mapnumbers ="" Then
mapnumbers =mapnumber
Else
mapnumbers =mapnumbers & "," &mapnumber
End If
Next
SSProcess.FreeMapFrame
SSParameter.SetParameterSTR "国土" , "获取图幅号字符串", mapnumbers
End Function
```


#### CreateOneMapFrame 创建一个图廓对象

**函数**：`CreateOneMapFrame`

**功能**：创建一个图廓对象。

**语法**：

```vbs
SSProcess.CreateOneMapFrame(frameIndex, frameCode)
```

**参数**：

- frameIndex 索引号,从0开始
- frameCode 创建的图廓对象编码

**说明**：

执行该函数，将在指定缓存图幅处创建一个图廓对象。


#### CutCurMapFrameToEdb 裁剪当前图幅数据到 EDB 文件

**函数**：`CutCurMapFrameToEdb`

**功能**：裁剪当前图幅数据到EDB文件。

**参数**：

- exportEdbFile 文件存储路径
- cutParameters 裁剪方案

**说明**：

使用此函数可以使当前剪裁的数据另存为EDB ,并且可以指定路径和工程名。

**示例**：

```vbs
Sub OnClick()
fileName = "D:\test\wuxi\edb\111.edb"
'清空转换参数
SSProcess.ClearDataXParameter
'设置导入文件格式为ArcGIS PDB
SSProcess.SetDataXParameter "DataType", "1"
SSProcess.SetDataXParameter "DataBoundID", "35"
SSProcess.SetDataXParameter "ExportPathName", ""
'裁剪输出到EDB
notCnotECodes="Codes_Const=9999013;" '不裁剪, 不打散的地物编码
notCnotECodes_Y = "Codes_Const_Symbol_Y=0;" '不裁剪, 不打散的地物包括所有Y类
notCnotECodes_E= "Codes_Const_Symbol_E=0;" '不裁剪, 不打散的地物包括所有E类
parameters = notCnotECodes & notCnotECodes_Y & notCnotECodes_E & doEdoCCodes
SSProcess.CutCurMapFrameToEdb fileName , parameters
SSProcess.SetDataXParameter "AddSystemFieldMode", "0"
'开始导入数据
SSProcess.ExportData
End Sub
```


#### FreeMapFrame 释放所有图幅缓存

**函数**：`FreeMapFrame`

**功能**：释放所有图幅缓存。

**语法**：

```vbs
SSProcess.FreeMapFrame()
```

**说明**：

执行该函数，将释放所有图幅缓存, 与CreateMapFrame或CreateMapFrameByRegion或CreateMapFrameByCenterLine配对使用。


#### GetCurMapFrame 获取当前图幅 ID

**函数**：`GetCurMapFrame`

**功能**：创建当前图幅ID。

**语法**：

```vbs
SSProcess.GetCurMapFrame()
```

**说明**：

执行该函数，将获取当前图幅ID。


#### GetCurMapFrameNumber 获取当前图幅的图幅号

**函数**：`GetCurMapFrameNumber`

**功能**：获取当前图幅的图幅号。

**语法**：

```vbs
SSProcess.GetCurMapFrameNumber()
```

**说明**：

执行此函数，返回当前图幅的图号，如无当前图幅，返回值为空。


#### GetMapFrameCenterPoint 获取指定缓存图幅的中心坐标

**函数**：`GetMapFrameCenterPoint`

**功能**：获取指定缓存图幅的中心坐标。

**语法**：

```vbs
SSProcess.GetMapFrameCenterPoint( frameIndex, x, y)
```

**参数**：

- frameIndex 缓存图幅索引号,从0开始
- x 缓存图幅中心点X坐标
- y 缓存图幅中心点Y坐标

**说明**：

执行该函数，将获取指定缓存图幅的指定中心点的坐标值。


#### GetMapFrameCount 获取缓存图幅的个数

**函数**：`GetMapFrameCount`

**功能**：获取缓存图幅的个数。

**语法**：

```vbs
SSProcess.GetMapFrameCount()
```

**说明**：

执行该函数，将获取缓存图幅的个数。


#### GetMapFrameNumber 获取指定坐标所在的图幅的图幅号

**函数**：`GetMapFrameNumber`

**功能**：获取指定坐标所在图幅的图幅号。

**语法**：

```vbs
SSProcess.GetMapFrameNumber(x, y)
```

**参数**：

- x X坐标值
- y Y坐标值

**说明**：

使用此函数将返回指定坐标处的图幅号。

**示例**：

```vbs
将获得坐标（100,100）处的图幅号：
Sub OnClick()
MapNumber=SSProcess.GetMapFrameNumber( 100,100)
End Sub
```


#### GetMapFramePoint 获取指定缓存图幅的坐标

**函数**：`GetMapFramePoint`

**功能**：获取指定缓存图幅的坐标。

**语法**：

```vbs
SSProcess.GetMapFramePoint( frameIndex, pIndex, x, y)
```

**参数**：

- frameIndex 缓存图幅索引号,从0开始
- pIndex 一个缓存图幅的角点索引,从0开始, 最大为3
- 3--------2
- | |
- 0--------1
- x 缓存图幅角点X坐标
- y 缓存图幅角点Y坐标

**说明**：

执行该函数，将获取指定缓存图幅的指定角点的坐标值。


#### PrintCurFrameMap 打印当前图幅

**函数**：`PrintCurFrameMap`

**功能**：打印当前图幅的地图数据。

**语法**：

```vbs
SSProcess.PrintCurFrameMap()
```

**说明**：

执行该函数，将打印当前图幅的地图数据。


#### 打印指定图幅 打印指定图幅 PrintMap 函数

**函数**：`PrintMap`

**功能**：打印指定图幅。

**语法**：

```vbs
SSProcess.PrintMap(frameIndex)
```

**参数**：

- frameIndex 缓存图幅索引号,从0开始

**说明**：

执行该函数，将打印指定缓存图幅的地图数据。

**示例**：

```vbs
Sub OnClick()
SSProcess.CreateMapFrame  '创建图幅缓存
nCount = SSProcess.GetMapFrameCount  '获取图幅个数
'开始打印
For i=0 To nCount-1
'打印图幅
SSProcess.PrintMap i
Next
End Sub
```


#### PrintMapByCoord 根据坐标范围打印

**函数**：`PrintMapByCoord`

**功能**：根据坐标范围打印。

**语法**：

```vbs
SSProcess.PrintMapByCoord(x,y,width,height,angle)
```

**参数**：

- x 中心X坐标
- y 中心Y坐标
- width 打印宽度
- height 打印高度
- angle 打印角度


#### SetCurMapFrame 设置指定坐标处为当前图幅

**函数**：`SetCurMapFrame`

**功能**：设置指定坐标处为当前图幅。

**语法**：

```vbs
SSProcess.SetCurMapFrame(x,y,addToMapLink,mapLinkTable)
```

**参数**：

- x 要设定图幅的中心点X坐标
- y 要设定图幅的中心点Y坐标
- addToMapLink 是否把图幅加入接合表, 0不加 1 加入
- mapLinkTable 图幅接合表名称

**说明**：

执行该函数，将设置指定坐标处为当前图幅。


---

### 外部函数

#### AddFunctionParameter 增加外部模块函数参数

**函数**：`AddFunctionParameter`

**功能**：设置外部模块函数参数。

**语法**：

```vbs
SSProcess.AddFunctionParameter( parameter )
```

**参数**：

- parameter 由外部函数约定的参数值,如"code=2110,color=RGB(255,0,0)"

**说明**：

执行该函数，可设置外部函数调用所需的参数,可以连续执行该函数实现设置多组参数值。

**示例**：

```vbs
下面示例设置外部模块函数参数值：
Sub OnClick()
SSProcess.AddFunctionParameter "code=2110,color=RGB(255,0,0)"
End Sub
```


#### Arc3pToCenter 三点弧求圆心半径和排序后弧角

**函数**：`Arc3pToCenter`

**功能**：三点弧求圆心、半径和排序后弧角。

**语法**：

```vbs
SSProcess.Arc3pToCenter(p0x, p0y, p1x, p1y, p2x, p2y, centerx, centery, r, ang1, ang2, ang3)
```

**参数**：

- p0x, p0y 圆上第一点坐标值
- p1x, p1y 圆上第二点坐标值
- p2x, p2y 圆上第三点坐标值
- centerx, centery 返回的圆心坐标值
- r 返回的半径值
- ang1, ang2, ang3 返回的弧角值，从大到小排列


#### Circle3pToCenter 三点求圆心和半径

**函数**：`Circle3pToCenter`

**功能**：三点求圆心和半径。

**语法**：

```vbs
SSProcess.Circle3pToCenter(p0x, p0y, p1x, p1y, p2x, p2y, centerx, centery, r)
```

**参数**：

- p0x, p0y 圆上第一点坐标值
- p1x, p1y 圆上第二点坐标值
- p2x, p2y 圆上第三点坐标值
- centerx, centery 返回的圆心坐标值
- r 返回的半径值

**示例**：

```vbs
将根据圆上3点获取其圆心和半径：
Sub OnClick()
p0x=1:p0y= 0
p1x= 0:p1y= 1
p2x=- 1:p2y= 0
SSProcess.Circle3pToCenter p0x,p0y,p1x ,p1y,p2x ,p2y,centerx ,centery, r
End Sub
```


#### ClearFunctionParameter 清空外部模块函数参数列表

**函数**：`ClearFunctionParameter`

**功能**：清空外部模块函数参数列表。

**语法**：

```vbs
SSProcess.ClearFunctionParameter()
```

**说明**：

执行该函数，将清空由函数AddFunctionParameter设置的外部函数参数列表。


#### DecryptData 解密字符串

**函数**：`DecryptData`

**功能**：解密字符串。

**语法**：

```vbs
SSProcess.DecryptData(passWord, data)
```

**参数**：

- password 密码
- data 需解密的字符串

**示例**：

```vbs
Sub OnClick()
strInfo = "山维科技"   '字符串内容
password = "123"   '密码
strInfo_Encrypt = SSProcess.EncryptData( password , strInfo)  '加密字符串
strInfo_Decrypt = SSProcess.DecryptData(password ,strInfo_Encrypt) '解密字符串
strmsg ="字符串内容："&strInfo&chr(13)&_+"加密后内容："&strInfo_Encrypt&chr(13)&_+"解密后内容："&strInfo_Decrypt
Msgbox strmsg, 64, "消息"
End Sub
```


#### EncryptData 加密字符串

**函数**：`EncryptData`

**功能**：加密字符串。

**语法**：

```vbs
SSProcess.EncryptData(passWord, data)
```

**参数**：

- passWord 密码
- data 需加密的字符串

**说明**：

使用此函数将加密字符串。

**示例**：

```vbs
Sub OnClick()
strInfo = "山维科技"  '字符串内容
password = "123"  '密码
strInfo_Encrypt = SSProcess.EncryptData( password , strInfo)  '加密字符串
strInfo_Decrypt = SSProcess.DecryptData(password ,strInfo_Encrypt) '解密字符串
strmsg = " 字符串内容："&strInfo&chr(13)&_+" 加密后内容："&strInfo_Encrypt&chr(13)&_+" 解密后内容："&strInfo_Decrypt
Msgbox strmsg,64," 消息"
End Sub
```


#### ExecuteCommand 执行外部模块函数

**函数**：`ExecuteCommand`

**功能**：调用外部模块函数。

**语法**：

```vbs
SSProcess.ExecuteCommand(libFileName, procName, commandDescription, transSelection)
```

**参数**：

- libFileName 外部模块名称,如果没有指定路径，则认为是放在EPS执行目录下,如"MyFunction.DLL"
- procName 外部模块函数名称,如"ChangeObjColor"
- commandDescription 进程条显示题头,用于说明当前调用函数正在处理的事件,如果为空,则不显示进程条
- transSelection 是否传递当前选择集对象到外部函数， 0（不传递）1（传递）

**说明**：

执行该函数，可调用在外部模块实现的函数,并把由脚本设置的外部函数参数传递到外部函数。
当transSelection为1时，外部函数的定义为如下格式
extern "C" __declspec(dllexport) BOOL
procName(CStringArray *, CGeObjList *, CProgressCtrl*);
当transSelection为0时，外部函数的定义为如下格式
extern "C" __declspec(dllexport) BOOL
procName(CStringArray *, CProgressCtrl*);

**示例**：

```vbs
下面示例调用模块MyFunction.DLL中的ChangeSelectionObjColor函数：
Sub OnClick()
SSProcess.AddFunctionParameter "code==2110,layer=123"
SSProcess.ExecuteCommand "MyFunction.DLL", "ChangeSelectionObjColor", "正在处理选择集对象换色...",1
SSProcess.ClearFunctionParameter
End Sub
```


#### ExecuteFuction 执行外部模块中的函数

**函数**：`ExecuteFuction`

**功能**：执行外部模块中的函数。

**语法**：

```vbs
SSProcess.ExecuteFuction(ModuleAndCmdName, commandDescription)
```

**参数**：

- ModuleAndCmdName 外部模块名称
- commandDescription 进程条显示题头,用于说明当前调用函数正在处理的事件,如果为空,则不显示进程条

**说明**：

执行外部模块中的函数。

**示例**：

```vbs
'面转点处理，参数：原面编码，新生成点编码，层名
Sub OnClick()
typech"2610003", "2610001","HYDP"
End Sub
Function typech(byval oldcode,byval newcode,byval ln)
SSProcess.ClearFunctionParameter
SSProcess.AddFunctionParameter "SearchObject = SSObj_Code = "&oldcode &"&&SSObj_Area< 400" '选择面，并且面的面积<400
SSProcess.AddFunctionParameter "ChangeCode = "&newcode
SSProcess.AddFunctionParameter "LayName = "& ln
SSProcess.ExecuteFuction "SSynthesizer.Syn.ChangeType.MinArea_to_point",""
SSProcess.ClearFunctionParameter
End Function
```


#### ExecuteSDLFunction 执行 SDL 命令

**函数**：`ExecuteSDLFunction`

**功能**：执行SDL命令。

**语法**：

```vbs
SSProcess.ExecuteSDLFunction(SDLCommandInfo, Reason)
```

**参数**：

- SDLCommandinfo SDL命令行
- Reason

**说明**：

该函数将调用指定的SDL命令。

**示例**：

```vbs
Sub OnClick()
'图形范围全视
SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
'图形重新生成
SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
End Sub
```


#### ExecuteSDLFunctionSetup 执行 SDL 功能参数设置

**函数**：`ExecuteSDLFunctionSetup`

**功能**：执行SDL功能参数设置。

**语法**：

```vbs
SSProcess.ExecuteSDLFunctionSetup(SDLName, CmdName, Parameters)
```

**参数**：

- SDLName 功能模块名称
- CmdName 执行命令名称
- Parameters 执行命令所需要的参数


#### LoadXLibrary 加载外部模块

**函数**：`LoadXLibrary`

**功能**：加载外部模块。

**语法**：

```vbs
SSProcess.LoadXLibrary(libFileName)
```

**参数**：

- libFileName 外部模块名称,如果没有指定路径，则认为是放在EPS执行目录下,如"MyFunction.DLL"

**说明**：

执行该函数，可加载外部模块到内存。
当使用ExecuteCommand同时执行多个同一模块内的函数时，可使用LoadXLibrary先加载模块，再调用ExecuteCommand，可加快程序运行速度，在ExecuteCommand执行完后，需调用UnLoadXLibrary卸载模块。

**示例**：

```vbs
下面示例加载外部模块：
Sub OnClick()
SSProcess.LoadXLibrary "MyFunction.DLL"
End Sub
```


#### MapCallBackFunction 执行系统回调函数

**函数**：`MapCallBackFunction`

**功能**：执行系统回调函数。

**语法**：

```vbs
SSProcess.MapCallBackFunction(param, data, flags)
```

**参数**：

- param 回调方法名称
- data 回调方法参数
- flags 预留标志值，一般为0

**说明**：

使用此函数将调用指定的方法或功能

**示例**：

```vbs
下面示例将启动平移功能：
Sub OnClick()
SSProcess.MapCallBackFunction "SDLCommand", "$SDL.SSEdit.Offset",0
End Sub
```


#### MapCallBackFunction1 执行系统回调函数

**函数**：`MapCallBackFunction1`

**功能**：执行系统回调函数。

**语法**：

```vbs
SSProcess.MapCallBackFunction1(param, data, flags)
```

**参数**：

- param 回调方法名称
- data 回调方法参数
- flags 预留标志值，一般为0

**说明**：

使用此函数将调用指定的方法或功能

**示例**：

```vbs
Sub OnClick()
'切换的选择集状态下
SSProcess.MapCallBackFunction1 "SDLCommand", "$SDL,SSEdit,ModIfy", 0
oldCursorStatus = SSProcess.GetCursorStatus
'0 输入光标 6 选择光标
SSProcess.SetCursorStatus 6
End Sub
```


#### UnLoadXLibrary 卸载外部模块

**函数**：`UnLoadXLibrary`

**功能**：卸载外部模块。

**语法**：

```vbs
SSProcess.UnLoadXLibrary( libFileName )
```

**参数**：

- libFileName 外部模块名称,如果没有指定路径，则认为是放在EPS执行目录下,如"MyFunction.DLL"

**说明**：

执行该函数，可卸载由函数LoadXLibrary加载的外部模块。

**示例**：

```vbs
下面示例卸载外部模块：
Sub OnClick()
SSProcess.UnLoadXLibrary "MyFunction.DLL"
End Sub
```


#### 执行外部程序 执行外部程序 WinExec 函数

**函数**：`WinExec`

**功能**：执行外部程序。

**语法**：

```vbs
SSProcess.WinExec(commandLine, flag)
```

**参数**：

- commandLine 外部程序名称
- flag 外部程序启动类型，0隐藏；1默认大小显示并激活；2启动但不激活；3满屏启动并激活 4默认大小显示但不激活

**说明**：

使用此函数将启动外部程序。

**示例**：

```vbs
以满屏启动记事本并激活。
Sub OnClick()
SSProcess.WinExec "notepad",4
End Sub
```


---

### 系统设置

#### AddDispFilterInfo 增加显示可控对象条件信息

**函数**：`AddDispFilterInfo`

**功能**：增加显示可控对象条件信息。

**语法**：

```vbs
SSProcess.AddDispFilterInfo(type, reverse, visible, colorvalid, color, values)
```

**参数**：

- Type 控制类型 0(编码)、1(创建时间)、2(修改时间)、3(被修改对象)、4(选择集对象)、5(数据标识)、6(图层)
- reverse 是否反向选择 0（正向）、1（反向）
- visible 本条设置指定的对象集合的可见与否 0（不可见）、1（可见）
- colorvalid 是否重新配色 0（不配）、1（配）
- color 重新配色的颜色值,分COLORNO(颜色号)、颜色值和RGB(红,绿，蓝)三种指定方式
- values 条件列表
| 参数 | 说明 |
|---|---|
| 类型 | 值方式 |
| 0( 编码 ) | 编码 1, 编码 2,... |
| 1( 创建时间 ) | 时间格式任意 , 如： 2004-1-1-9:10 , 2005-05-05-10:15 一段时间 2004, 2005-05-05 分时秒可顺次前省略 2004-1-1-9:10, 某一时间以后 , 2005-05-05 截止到某一时间 |
| 2( 修改时间 ) | 同上 |
| 3( 被修改对象 ) | 是创建时间与修改时间不等的对象，值域不填 |
| 4( 选择集对象 ) | 可以指定对象的 ID( 如 ID1,ID2,...) 或指选择条件组名 ( 如 SelectionGroup_ 房屋 ) |
| 5( 数据标识 ) | 即 [ 选择集 ] 属性列表中显示的“用户标识 ” 数据标识等于空用“ NULL” 表示 |
| 6( 图层 ) | 层名 1, 层名 2,... |

**说明**：

执行该函数，可实现根据对象的属性特征分类分组显示控制,包括对象的可见性，动态显示颜色。

**示例**：

```vbs
下面示例控制指定编码显示：
Sub OnClick()
SSProcess.ClearDispFilterInfo
SSProcess.AddDispFilterInfo 0, 0, 1, 0, "RGB(255,0,0)", "2010,3010"
End Sub
下面示例控制显示选择组"房屋"中的对象：
Sub OnClick()
SSProcess.ClearDispFilterInfo
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "房屋", "SSObj_Code", "==", "3001,3002,3003,3004"
SSProcess.SetSelectConditionGroup "房屋", "[层数]", ">", "10"
SSProcess.AddDispFilterInfo 4, 0, 1, 0, "RGB(255,0,0)", "SelectionGroup_房屋"
End Sub
```


#### AddFeatureCode 添加编码定义

**函数**：`AddFeatureCode`

**功能**：添加编码定义。

**语法**：

```vbs
SSProcess.AddFeatureCode(code, geoType, symbolType, objName, layerName, lineWidth, color)
```

**参数**：

- code 编码
- geoType 地物几何类型
- symbolType 符号类型
- objName 对象名称
- layerName 层名
- lineWidth 线宽
- color 颜色

**说明**：

执行该函数将添加编码定义。

**示例**：

```vbs
Sub OnClick()
code = "100"  'EPS编码
geoType = "1" '地物几何类型
symbolType= "2"   '符号类型
objName =  "HK" '对象名称
layerName = "HK" '层名
lineWidth = 10 '线宽
color = RGB(255, 0,0) '颜色
SSProcess.AddFeatureCode code,geoType, symbolType,objName ,layerName, lineWidth,color
End Sub
```


#### AddFontClass 添加注记分类定义

**函数**：`AddFontClass`

**功能**：添加注记分类定义。

**语法**：

```vbs
SSProcess.AddFontClass(code, objName, layerName, fontName, fontWidth, fontHeight, color)
```

**参数**：

- code 注记分类号
- objName 对象名称
- layerName 层名
- fontN ame 字体
- fontWidth 注记字高
- fontHeight 注记字色

**说明**：

执行该函数将添加注记分类定义。

**示例**：

```vbs
Sub OnClick()
Code = "hoenking" '注记分类号
objName = "HK"  '对象名称
layerName = "HK"  '层名
fontName = "微软雅黑" '字体
fontWidth = 300 '注记字高
fontHeight = 300   '注记字宽
color = RGB(255, 0,0)  '注记颜色
SSProcess.AddFontClass code,objName, layerName,fontName ,fontWidth, fontHeight,color
End Sub
```


#### AddInputParameter 添加录入参数

**函数**：`AddInputParameter`

**功能**：添加录入参数。

**语法**：

```vbs
SSProcess.AddInputParameter(name, value, type, options, description)
```

**参数**：

- name 参数名称
- value 参数值
- type 类型，0（参数值）、1（颜色）、2（时间）、4（编码）、5（注记）、6（图层）、7（地物类型）、8（线型）、9（线宽）、10（是否）、12（选择文件夹）
- options 枚举值
- description 说明

**说明**：

执行该函数，将弹出对话框，录入相应参数，获取参数值，将值填到扩展属性字段中。

**示例**：

```vbs
下面示例画完房屋(编码：31030131)后弹出对话框录入结构和层数：
Sub OnClick()
SSParameter.GetParameterINT "AfterAddLine", "CurrentObjID", "0", objID
If objID=0 Then Exit Sub
objCode = SSProcess.GetObjectAttr (objID, "SSObj_Code")
If objCode = 31030131
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "结构类型", "砖", 0, "砖,砼,混,木", "输入建筑材料"
SSProcess.AddInputParameter "层数", "1", 0,",2,3,4,5,6,7,8,9,10,11,12,13,14, 15,16,17,18",""
result =SSProcess.ShowInputParameterDlg ("房屋注记")
If result = 1 Then
SSProcess.UpdateScriptDlgParameter 1
jglx = SSProcess.GetInputParameter ("结构类型")
cs = SSProcess.GetInputParameter ("层数")
SSProcess.SetObjectAttr objID, "[实际层数]",cs
SSProcess.SetObjectAttr objID, "[结构类型]",jglx
End If
End If
SSProcess.ObjectDeal objID, "FreeDisplayList", parameters, result
SSProcess.RefreshView
End Sub
```


#### ClearDispFilterInfo 清空显示可控对象条件信息

**函数**：`ClearDispFilterInfo`

**功能**：清空显示可控对象条件信息。

**语法**：

```vbs
SSProcess.ClearDispFilterInfo()
```

**说明**：

执行该函数，将清空由AddDispFilterInfo设置的可控对象条件信息。

**示例**：

```vbs
下面示例清空显示可控对象条件信息：
Sub OnClick()
SSProcess.ClearDispFilterInfo
SSProcess.AddDispFilterInfo 0, 0, 1, 0, "RGB(255,0,0)", "2010,3010"
End Sub
```


#### ClearInputParameter 清除录入参数

**函数**：`ClearInputParameter`

**功能**：清除录入参数。

**语法**：

```vbs
SSProcess.ClearInputParameter()
```

**说明**：

使用此函数清除录入的参数。

**示例**：

```vbs
Dim oldCursorStatus
Dim fromMode
Sub OnInitScript()
mode = 1 ‘=0 无参数对话框 =1 有参数对话框
title="高程插点"
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "高程取值方式" , "三角网", 0, "周边高程数据,三角网,格网" , ""
SSProcess.AddInputParameter "高程点编码" , "831000", 4 , "", ""
SSProcess.ShowScriptDlg mode,title
oldCursorStatus = SSProcess.GetCursorStatus
'0 输入光标 6 选择光标
SSProcess.SetCursorStatus 0
End Sub
Function OnLButtonDown(x , y, spx , spy, flags)
'更新对话框上参数到内存
SSProcess.UpdateScriptDlgParameter 1
fromMode = SSProcess.GetInputParameter ( "高程取值方式" )
If fromMode = "周边高程数据" Then
fromMode = "0"
ElseIf fromMode = "三角网" Then
fromMode = "1"
ElseIf fromMode = "格网" Then
fromMode = "2"
End If
z = SSProcess.GetPointHeight (spx , spy, Clng (fromMode))
If z<9000 Then
SSProcess.PushUndoMark
'更新对话框上参数到内存
SSProcess.UpdateScriptDlgParameter 1
'取编码
code = SSProcess.GetInputParameter ( "高程点编码" )
SSProcess.CreateNewObjByCode code
SSProcess.AddNewObjPoint spx , spy, z, 2, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
'刷屏
SSProcess.RefreshView
Else
MsgBox "指定位置无法提取高程值."
End If
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
'0 输入光标 6 选择光标
SSProcess.SetCursorStatus oldCursorStatus
End Sub
```


#### DeleteFeatureCode 删除编码定义

**函数**：`DeleteFeatureCode`

**功能**：删除编码定义。

**语法**：

```vbs
SSProcess.DeleteFeatureCode(code）
```

**参数**：

- code 地物编码

**说明**：

执行该函数，从模版中删除指定编码的地物。


#### DeleteFontClass 删除注记分类定义

**函数**：`DeleteFontClass`

**功能**：删除注记分类定义。

**语法**：

```vbs
SSProcess.DeleteFontClass(code)
```

**参数**：

- code 注记分类号

**说明**：

执行该函数，从模版中删除指定编码的地物注记。


#### EpsProgressCreate 创建系统进程条

**函数**：`EpsProgressCreate`

**功能**：创建系统进程条。

**语法**：

```vbs
SSProcess.EpsProgressCreate(size, msg)
```

**参数**：

- size 进度条的长度
- msg 进度条的标题

**说明**：

执行该函数，创建一个指定样式的eps进度条，在不指定样式的情况下系统默认为附加在主窗口状态条上。
注意：创建了进度条在使用后务必删除。

**示例**：

```vbs
Sub OnClick()
SSProcess.EpsProgressSetStyle 1
SSProcess.EpsProgressCreate 1000, "进度条"
For i = 0 To 1000
SSProcess.EpsProgressSetPos i
Next
SSProcess.EpsProgressDelete
End Sub
```


#### EpsProgressDelete 删除系统进程条

**函数**：`EpsProgressDelete`

**功能**：删除系统进程条。

**语法**：

```vbs
SSProcess.EpsProgressDelete()
```

**说明**：

执行该函数，删除eps进度条。
注意：创建了进度条才可以删除，没有创建不需要删除。

**示例**：

```vbs
Sub OnClick()
SSProcess.EpsProgressSetStyle 1
SSProcess.EpsProgressCreate 1000, "进度条"
For i = 0 To 1000
SSProcess.EpsProgressSetPos i
Next
SSProcess.EpsProgressDelete
End Sub
```


#### EpsProgressGetPos 获取系统进程条当前位置

**函数**：`EpsProgressGetPos`

**功能**：获取系统进程条当前位置。

**语法**：

```vbs
SSProcess.EpsProgressGetPos()
```

**说明**：

使用此函数将返回系统进程条当前位置。


#### EpsProgressGetRange 获取系统进程条区间

**函数**：`EpsProgressGetRange`

**功能**：清空显示可控对象条件信息。

**语法**：

```vbs
SSProcess.EpsProgressGetRange(nLower, nUpper)
```

**参数**：

- nLower 系统进程条区间最小值
- nUpper 系统进程条区间最大值

**说明**：

使用此函数将返回系统进程条区间。


#### EpsProgressSetPos 设置系统进程条当前位置

**函数**：`EpsProgressSetPos`

**功能**：清空显示可控对象条件信息。

**语法**：

```vbs
SSProcess.EpsProgressSetPos(pos)
```

**参数**：

- pos 系统进程条当前位置

**说明**：

执行该函数，将清空由AddDispFilterInfo设置的可控对象条件信息
下面示例清空显示可控对象条件信息：
使用此函数设置系统进程条当前位置

**示例**：

```vbs
Sub OnClick()
SSProcess.EpsProgressSetStyle 1
SSProcess.EpsProgressCreate 1000, "进度条"
For i = 0 To 1000
SSProcess.EpsProgressSetPos i
Next
SSProcess.EpsProgressDelete
End Sub
```


#### EpsProgressSetRange32 设置系统进程条区间

**函数**：`EpsProgressSetRange32`

**功能**：设置系统进程条区间。

**语法**：

```vbs
SSProcess.EpsProgressSetRange32(nLower, nUpper)
```

**参数**：

- nLower 系统进程条区间最小值
- nUpper 系统进程条区间最大值

**说明**：

使用此函数设置系统进程条区间。

**示例**：

```vbs
Sub OnClick()
my_time =100
SSProcess.EpsProgressSetStyle 1
SSProcess.EpsProgressCreate 100,"正在转换..."
SSProcess.EpsProgressSetStep  1
For i = 0 To 100
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
geocount = 100
If geocount > 0 Then
p1count =50:p2count =100:p3count =200:p4count =20
SSProcess.EpsProgressUpdateMsg "正在处理等高线..."
SSProcess.EpsProgressSetRange32  0, p1count+ 1
SSProcess.EpsProgressSetPos 0
For i=0 To p1count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0, p2count+ 1
SSProcess.EpsProgressUpdateMsg "重新构三角网..."
SSProcess.EpsProgressSetPos  0
For i=0 To p2count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0,p3count+ 1
SSProcess.EpsProgressUpdateMsg "生成等深线..."
SSProcess.EpsProgressSetPos 0
For i=0 To p3count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0, p4count+ 1
SSProcess.EpsProgressUpdateMsg "关闭三角网..."
SSProcess.EpsProgressSetPos 0
For i=0 To p4count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
End If
SSProcess.EpsProgressDelete
SSProcess.RefreshView
End Sub
```


#### EpsProgressSetStep 设置系统进程条步距

**函数**：`EpsProgressSetStep`

**功能**：设置系统进程条步距。

**语法**：

```vbs
SSProcess.EpsProgressSetStep(step)
```

**参数**：

- step 系统进程条步距

**说明**：

该函数用于设置系统进程条的步距。

**示例**：

```vbs
Sub OnClick()
my_time =100
SSProcess.EpsProgressSetStyle 1
SSProcess.EpsProgressCreate 100,"正在转换..."
SSProcess.EpsProgressSetStep  1
For i = 0 To 100
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
geocount = 100
If geocount > 0 Then
p1count =50:p2count =100:p3count =200:p4count =20
SSProcess.EpsProgressUpdateMsg "正在处理等高线..."
SSProcess.EpsProgressSetRange32  0, p1count+ 1
SSProcess.EpsProgressSetPos 0
For i=0 To p1count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0, p2count+ 1
SSProcess.EpsProgressUpdateMsg "重新构三角网..."
SSProcess.EpsProgressSetPos  0
For i=0 To p2count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0,p3count+ 1
SSProcess.EpsProgressUpdateMsg "生成等深线..."
SSProcess.EpsProgressSetPos 0
For i=0 To p3count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0, p4count+ 1
SSProcess.EpsProgressUpdateMsg "关闭三角网..."
SSProcess.EpsProgressSetPos 0
For i=0 To p4count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
End If
SSProcess.EpsProgressDelete
SSProcess.RefreshView
End Sub
```


#### EpsProgressSetStyle 设置进度条样式

**函数**：`EpsProgressSetStyle`

**功能**：设置进度条样式。

**语法**：

```vbs
SSProcess.EpsProgressSetStyle(nStyle)
```

**参数**：

- nStyle 进度条样式：0 附加在主窗口的状态条上 1 可浮动的独立窗口

**说明**：

执行该函数，设置EPS进度条的样式。

**示例**：

```vbs
Sub OnClick()
SSProcess.EpsProgressSetStyle 1
SSProcess.EpsProgressCreate 1000, "进度条"
SSProcess.SetDealMsg "dealMsg"
For i = 0 To 1000
SSProcess.EpsProgressSetPos i
Next
SSProcess.EpsProgressDelete
End Sub
```


#### EpsProgressStepIt 系统进程条单步滚动

**函数**：`EpsProgressStepIt`

**功能**：系统进程条的单步滚动。

**说明**：

执行该函数，将清空由AddDispFilterInfo设置的可控对象条件信息

**示例**：

```vbs
Sub OnClick()
my_time =100
SSProcess.EpsProgressSetStyle 1
SSProcess.EpsProgressCreate 100,"正在转换..."
SSProcess.EpsProgressSetStep  1
For i = 0 To 100
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
geocount = 100
If geocount > 0 Then
p1count =50:p2count =100:p3count =200:p4count =20
SSProcess.EpsProgressUpdateMsg "正在处理等高线..."
SSProcess.EpsProgressSetRange32  0, p1count+ 1
SSProcess.EpsProgressSetPos 0
For i=0 To p1count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0, p2count+ 1
SSProcess.EpsProgressUpdateMsg "重新构三角网..."
SSProcess.EpsProgressSetPos  0
For i=0 To p2count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0,p3count+ 1
SSProcess.EpsProgressUpdateMsg "生成等深线..."
SSProcess.EpsProgressSetPos 0
For i=0 To p3count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0, p4count+ 1
SSProcess.EpsProgressUpdateMsg "关闭三角网..."
SSProcess.EpsProgressSetPos 0
For i=0 To p4count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
End If
SSProcess.EpsProgressDelete
SSProcess.RefreshView
End Sub
```


#### EpsProgressUpdateMsg 更新系统进程条信息

**函数**：`EpsProgressUpdateMsg`

**功能**：更新系统进程条信息。

**语法**：

```vbs
SSProcess.EpsProgressUpdateMsg(msg)
```

**参数**：

- msg 提示信息。

**说明**：

该函数用于更新系统进程条信息。

**示例**：

```vbs
Sub OnClick()
my_time =100
SSProcess.EpsProgressSetStyle 1
SSProcess.EpsProgressCreate 100,"正在转换..."
SSProcess.EpsProgressSetStep  1
For i = 0 To 100
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
geocount = 100
If geocount > 0 Then
p1count =50:p2count =100:p3count =200:p4count =20
SSProcess.EpsProgressUpdateMsg "正在处理等高线..."
SSProcess.EpsProgressSetRange32  0, p1count+ 1
SSProcess.EpsProgressSetPos 0
For i=0 To p1count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0, p2count+ 1
SSProcess.EpsProgressUpdateMsg "重新构三角网..."
SSProcess.EpsProgressSetPos  0
For i=0 To p2count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0,p3count+ 1
SSProcess.EpsProgressUpdateMsg "生成等深线..."
SSProcess.EpsProgressSetPos 0
For i=0 To p3count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
SSProcess.EpsProgressSetRange32 0, p4count+ 1
SSProcess.EpsProgressUpdateMsg "关闭三角网..."
SSProcess.EpsProgressSetPos 0
For i=0 To p4count -1
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
SSProcess.EpsProgressStepIt
End If
SSProcess.EpsProgressDelete
SSProcess.RefreshView
End Sub
```


#### GetCodeAttrTableName 获取编码挂接的属性表

**函数**：`GetCodeAttrTableName`

**功能**：获取编码挂接的属性表。

**语法**：

```vbs
SSProcess.GetCodeAttrTableName(code, tableType)
```

**参数**：

- code 编码
- tableType 类型 0 点属性表、1 线属性表、2 面属性表、3 注记属性表

**说明**：

执行此函数将返回指定图层所挂接的相应属性表名称。

**示例**：

```vbs
下面示例将返回当前工程的DEFAULT层的点属性表名称。
'获取房屋面属性表
Sub OnClick()
code="3103013":tableType = 2
a=SSProcess.GetCodeAttrTableName(code,tableType)
Msgbox a
End Sub
```


#### GetColorInfoTB 获取系统颜色表

**函数**：`GetColorInfoTB`

**功能**：获取当前工程所用的颜色信息表。

**语法**：

```vbs
SSProcess.GetColorInfoTB()
```

**说明**：

获取当前工程所用的颜色信息表。
返回颜色表名。

**示例**：

```vbs
Sub OnClick()
'获取模板中的所有表
strTemplateFile = "基础地理标准_500.mdt"
If strTemplateFile = "0" Then
Msgbox "获取本工程模板文件名失败！"
Exit Sub
End If
strTemplateFile = SSProcess.GetSysPathName(1)&strTemplateFile
Dim tableNames1, TableCount
Dim  TableNames(500),ColorInfoTBNames(500),ColorTBCount
ColorTBCount = 0
SSProcess.OpenAccessMdb strTemplateFile
SSProcess.GetAccessTableNames strTemplateFile, tableNames1
SSFunc.ScanString tableNames1, ",", TableNames, TableCount
'获取颜色信息表:ColorInfoTB
For i = 0 To  TableCount - 1
If Instr(1, TableNames(i), "ColorInfoTB", 1) THEN
ColorInfoTBNames(ColorTBCount) = TableNames(i)
ColorTBCount = ColorTBCount + 1
End If
Next
ColorTB1= SSProcess.GetColorInfoTB
End Sub
```


#### GetCursorPoint 获取当前光标屏幕坐标

**函数**：`GetCursorPoint`

**功能**：获取当前光标屏幕坐标。

**语法**：

```vbs
SSProcess.GetCursorPoint(x,y)
```

**参数**：

- x 当前光标屏幕x 坐标
- y 当前光标屏幕y 坐标

**说明**：

执行该函数，将获取当前光标屏幕坐标。

**示例**：

```vbs
Sub OnClick()
SSProcess.GetCursorPoint x ,y
Msgbox x&"---"&y
End Sub
```


#### GetCursorSurveyPoint 获取当前光标实地坐标

**函数**：`GetCursorSurveyPoint`

**功能**：获取当前光标实地坐标。

**语法**：

```vbs
SSProcess.GetCursorSurveyPoint(x, y, z)
```

**参数**：

- x 目前光标实地x坐标
- y 目前光标实地y坐标
- z 目前光标实地z坐标

**说明**：

执行该函数，将获取当前光标实地坐标。

**示例**：

```vbs
下面实例提取指定点位坐标。
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
title="提取坐标|移动光标回车提取"
SSProcess.ShowScriptDlg mode,title
'SSProcess.ShowScriptUserDefDlg title, dlgTemplateName, dlgWidth, dlgHeight, colCount, titleWidth, valueWidth
Ixycount = 0
'将操作状态切换为选择集
SSProcess.MapCallBackFunction "SDLCommand","$SDL,SSEdit,ModIfy" , 0
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
Dim fileName
fileName = SSProcess.SelectFileName(1, "",0,"坐标文件(*.txt)|*.txt|All Files (*.*)|*.*||")
If fileName ="" Then Exit Sub
Dim fso, tf, i , j
Dim chline, code, layer , color, pointcount
Dim x, y, z , pointtype, name
Set fso = CreateObject("Scripting.FileSystemObject")
Set tf = fso.CreateTextFile(fileName , True)
For j=0 To Ixycount- 1
chline = getCoordinate( j)
tf.WriteLine (chline)
Next
tf.Close
End Sub
Sub OnCancel()
End Sub
Dim Ixycount
Dim getCoordinate(1000)
Function OnKeyDown(nChar , nRepCnt, nFlags)
If nChar =13 Then
SSProcess.GetCursorSurveyPoint x , y, z
x =formatnumber( x,3,-1,0,0)
y =formatnumber( y,3,-1,0,0)
getCoordinate (Ixycount) = x & "," & y
Ixycount =Ixycount + 1
End If
End Function
```


#### GetFeatureCodeInfo 获取编码信息

**函数**：`GetFeatureCodeInfo`

**功能**：清空显示可控对象条件信息。

**语法**：

```vbs
SSProcess.GetFeatureCodeInfo(code, fields)
```

**参数**：

- code 要查找的编码
- fields 要查找的编码信息名称 , 有多个时用逗号分隔 ,编码信息名称见下表
| 参数 | 说明 |
|---|---|
| Code | EPS 编码 |
| Byname | 转换编码 |
| LayerName | 层名 |
| ObjectName | 对象名称 |
| Type | 几何类型 |
| LineType | 连接类型 |
| LineWidth | 线宽 |
| LineColor | 颜色 |
| Filter | 是否过滤 |
| Explode | 是否打散 |
| Reverse | 是否反向 |
| Thickness | 厚度 |
| Others | 其他 |
| AttrType | 属性类型 |
| Reference | 参考 |
| ExtraInfo | 附加信息 |

**说明**：

执行该函数可获取指定编码的定义信息。
返回值为与fields信息个数对应的编码信息 ,有多个时,用逗号分隔, 如果内容为空,则指定编码不存在。

**示例**：

```vbs
下面示例获取编码2110的别名和编码名称信息。
Sub OnClick()
codeinfo = SSProcess.GetFeatureCodeInfo ("2110","Byname,ObjectName")
MsgBox codeinfo
End Sub
```


#### GetFeatureCodeTB 获取系统编码表

**函数**：`GetFeatureCodeTB`

**功能**：获取系统编码表。

**语法**：

```vbs
SSProcess.GetFeatureCodeTB()
```

**说明**：

执行该函数，将返回当前系统使用的编码表名称。

**示例**：

```vbs
下面示例获取当前使用的编码模板表名称：
Sub OnClick()
Dim featureName
featureName = SSProcess.GetFeatureCodeTB()
End Sub
```


#### GetFieldCodeMap 获取属性编码与名称对照表

**函数**：`GetFieldCodeMap`

**功能**：获取属性编码与名称对照表。

**语法**：

```vbs
SSProcess.GetFieldCodeMap(tableName, fieldName, codes, names)
```

**参数**：

- tableName 表名称（可为空）
- fieldname 文件名
- codes 代码
- names 值

**说明**：

获取模板同文件路径下的*.dic文件内容。


#### GetFontClassInfo 获取分类号信息

**函数**：`GetFontClassInfo`

**功能**：清空显示可控对象条件信息。

**语法**：

```vbs
SSProcess.GetFontClassInfo(code, fields）
```

**参数**：

- code 要查找的分类号
- fields 要查找的分类号信息名称 , 有多个时用逗号分隔 ,分类号信息名称见下表：
| 参数 | 说明 |
|---|---|
| FontClass | 分类号 |
| FontColor | 文字颜色 |
| StringAngle | 字串角度 |
| FontName | 字体名 |
| FontWidth | 字宽 |
| FontHeight | 字高 |
| FontInterval | 字隔 |
| FontWeight | 字重 |
| IlaticAngle | 斜体 |
| WordAngle | 字角度 |
| DownAngle | 耸肩 |
| Underline | 下划线 |
| Memo | 备注 |
| LayerName | 层名 |
| FontByname | 字体别名 |
| Filter | 是否过滤 |
| Byname | 别名 |
| HS | 对齐方式 |

**说明**：

执行该函数可获取指定注记分类号的定义信息。
返回值为与fields信息个数对应的分类号信息 ,有多个时,用逗号分隔, 如果内容为空,则指定分类号不存在。

**示例**：

```vbs
下面示例获取注记分类号0的字体名和分类名称信息：
Sub OnClick()
codeinfo = SSProcess.GetFontClassInfo("0", "Fontname,Memo" )
MsgBox codeinfo
End Sub
```


#### GetFrameCode 获取图廓编码

**函数**：`GetFrameCode`

**功能**：获取图廓编码。

**语法**：

```vbs
SSProcess.GetFrameCode()
```

**说明**：

执行该函数，获取系统环境设置下面图廓的编码。

**示例**：

```vbs
Sub OnClick()
SSProcess. GetFrameCode()
End Sub
```


#### GetGeoID 申请对象 ID

**函数**：`GetGeoID`

**功能**：申请对象ID。

**语法**：

```vbs
SSProcess.GetGeoID()
```

**说明**：

该函数为系统设置函数，是指在当前最大地物ID的基础上申请对象ID。

**示例**：

```vbs
下面示例是申请的一个对象ID：
Sub OnClick()
a= SSProcess.GetGeoMaxID
Msgbox a
SSProcess.GetGeoID
b= SSProcess.GetGeoMaxID
Msgbox b
End Sub
```


#### GetGeoMaxID 获取当前最大地物 ID

**函数**：`GetGeoMaxID`

**功能**：GetGeoMaxID 函数

**语法**：

```vbs
SSProcess.GetGeoMaxID()
```

**说明**：

获取当前最大地物ID。

**示例**：

```vbs
Sub OnClick()
a= SSProcess.GetGeoMaxID
Msgbox a
End Sub
```


#### GetGridCellInfo 获取多列参数编辑框单元格信息

**函数**：`GetGridCellInfo`

**功能**：获取多列参数编辑框单元格信息。

**语法**：

```vbs
SSProcess.GetGridCellInfo(row, col)
```

**参数**：

- row 参数编辑框行号
- col 参数编辑框列号

**说明**：

该函数为系统设置函数，用于获取多列参数编辑框单元格信息。

**示例**：

```vbs
下面示例将获取参数编辑框中第1行第3列单元格信息。
Sub OnClick()
SSProcess.SetGridHeadInfo "姓名,性别,学历,工作年限(年）" , "100,50,100,100" , 0, 1, 0
SSProcess.SetGridCellInfo 1, 1 , "张三", ""
SSProcess.SetGridCellInfo 1, 2 , "男", "男,女"
SSProcess.SetGridCellInfo 1, 3 , "本科", ""
SSProcess.SetGridCellInfo 1, 4 , "3", ""
SSProcess.ShowGridEditDlg "查询统计结果"
Msgbox SSProcess.GetGridCellInfo(1, 3)
End Sub
```


#### GetGridColCount 获取多列参数编辑框表格列数

**函数**：`GetGridColCount`

**功能**：获取多列参数编辑框表格列数。

**语法**：

```vbs
SSProcess.GetGridColCount()
```

**说明**：

该函数获取多列参数编辑框表格列数

**示例**：

```vbs
Sub OnClick()
'获取多列参数编辑框表格列数
SSProcess.SetGridHeadInfo "姓名,性别,学历,工作年限（年）" , "100,50,100,100" , 0, 1, 0
SSProcess.SetGridCellInfo 1, 1 , "张一", ""
SSProcess.SetGridCellInfo 1, 2 , "男", "男,女"
SSProcess.SetGridCellInfo 1, 3 , "本科", ""
SSProcess.SetGridCellInfo 1, 4 , "1", ""
SSProcess.SetGridCellInfo 2, 1 , "张二", ""
SSProcess.SetGridCellInfo 2, 2 , "男", "男,女"
SSProcess.SetGridCellInfo 2, 3 , "本科", ""
SSProcess.SetGridCellInfo 2, 4 , "2", ""
SSProcess.SetGridCellInfo 3, 1 , "张三", ""
SSProcess.SetGridCellInfo 3, 2 , "女", "男,女"
SSProcess.SetGridCellInfo 3, 3 , "本科", ""
SSProcess.SetGridCellInfo 3, 4 , "3", ""
SSProcess.SetGridCellInfo 4, 1 , "张四", ""
SSProcess.SetGridCellInfo 4, 2 , "女", "男,女"
SSProcess.SetGridCellInfo 4, 3 , "本科", ""
SSProcess.SetGridCellInfo 4, 4 , "3", ""
SSProcess.SetGridCellInfo 5, 1 , "张五", ""
SSProcess.SetGridCellInfo 5, 2 , "女", "男,女"
SSProcess.SetGridCellInfo 5, 3 , "本科", ""
SSProcess.SetGridCellInfo 5, 4 , "6", ""
SSProcess.ShowGridEditDlg "查询统计结果"
Msgbox SSProcess.GetGridColCount
End Sub
```


#### GetGridRowCount 获取多列参数编辑框表格行数

**函数**：`GetGridRowCount`

**功能**：获取多列参数编辑框表格行数。

**语法**：

```vbs
SSProcess.GetGridRowCount()
```

**示例**：

```vbs
Sub OnClick()
'获取多行参数编辑框表格列数
SSProcess.SetGridHeadInfo "姓名,性别,学历,工作年限（年）" , "100,50,100,100" , 0, 1, 0
SSProcess.SetGridCellInfo 1, 1 , "张一", ""
SSProcess.SetGridCellInfo 1, 2 , "男", "男,女"
SSProcess.SetGridCellInfo 1, 3 , "本科", ""
SSProcess.SetGridCellInfo 1, 4 , "1", ""
SSProcess.SetGridCellInfo 2, 1 , "张二", ""
SSProcess.SetGridCellInfo 2, 2 , "男", "男,女"
SSProcess.SetGridCellInfo 2, 3 , "本科", ""
SSProcess.SetGridCellInfo 2, 4 , "2", ""
SSProcess.SetGridCellInfo 3, 1 , "张三", ""
SSProcess.SetGridCellInfo 3, 2 , "女", "男,女"
SSProcess.SetGridCellInfo 3, 3 , "本科", ""
SSProcess.SetGridCellInfo 3, 4 , "3", ""
SSProcess.SetGridCellInfo 4, 1 , "张四", ""
SSProcess.SetGridCellInfo 4, 2 , "女", "男,女"
SSProcess.SetGridCellInfo 4, 3 , "本科", ""
SSProcess.SetGridCellInfo 4, 4 , "3", ""
SSProcess.SetGridCellInfo 5, 1 , "张五", ""
SSProcess.SetGridCellInfo 5, 2 , "女", "男,女"
SSProcess.SetGridCellInfo 5, 3 , "本科", ""
SSProcess.SetGridCellInfo 5, 4 , "6", ""
SSProcess.ShowGridEditDlg "查询统计结果"
Msgbox SSProcess.GetGridRowCount
End Sub
```


#### GetGridSelRowCount 获取表格当前选择行数

**函数**：`GetGridSelRowCount`

**功能**：获取表格当前选择行数。

**语法**：

```vbs
SSProcess.GetGridSelRowCount()
```


#### GetInputParameter 获取录入参数

**函数**：`GetInputParameter`

**功能**：获取录入参数。

**语法**：

```vbs
SSProcess.GetInputParameter(name)
```

**参数**：

- name 参数名称

**说明**：

执行该函数，将弹出对话框，录入相应参数。


#### GetLayerAttrTableName 获取图层挂接的属性表

**函数**：`GetLayerAttrTableName`

**功能**：获取图层挂接的属性表。

**语法**：

```vbs
SSProcess.GetLayerAttrTableName(layerName, tableType)
```

**参数**：

- layerName 图层名称
- tableType 类型 0点属性表、1线属性表、2面属性表、3注记属性表

**说明**：

执行此函数将返回指定图层所挂接的相应属性表名称。

**示例**：

```vbs
下面示例将返回当前工程的DEFAULT层的点属性表名称。
Sub OnClick()
layerName ="DEFAULT": objType =0
a=SSProcess. GetLayerAttrTableName (layerName, objType)
Msgbox a
End Sub
```


#### GetLayerCount 获取图层数

**函数**：`GetLayerCount`

**功能**：获取图层数。

**语法**：

```vbs
SSProcess.GetLayerCount()
```

**参数**：

- 返回系统图层个数。

**说明**：

执行该函数，返回系统图层个数。

**示例**：

```vbs
下面获取系统图层个数：
Sub OnClick()
Dim count
count = SSProcess.GetLayerCount
End Sub
```


#### GetLayerName 获取图层名称

**函数**：`GetLayerName`

**功能**：获取图层名称。

**语法**：

```vbs
SSProcess.GetLayerName(index)
```

**参数**：

- index 要返回层名的索引号,从0到层数减一

**说明**：

执行该函数，将返回指定索引的图层名称。
返回图层名,如果索引号超出范围,则返回空字符串。

**示例**：

```vbs
下面示例获取指定索引的图层名称：
Sub OnClick()
Dim layerName
layerName = SSProcess.GetLayerName(5)
End Sub
```


#### GetLayerStatus 获取图层状态

**函数**：`GetLayerStatus`

**功能**：获取图层状态。

**语法**：

```vbs
SSProcess.GetLayerStatus(layername, type)
```

**参数**：

- layername 要查看的用户层名
- type 用户层状态 1（是否可见） 2（是否锁定编辑）、4（是否锁定索引）、8（是否已拓扑）、16（是否已做悬挂点处理）

**说明**：

执行该函数，可以获取指定用户层的状态。
返回 0为未锁定， 返回 1为锁定。

**示例**：

```vbs
下面示例获取用户层开关状态：
Sub OnClick()
Dim status
status = SSProcess.GetLayerStatus("控制点", 1)
End Sub
```


#### GetMapBoxLayerName 获取图廓层名

**函数**：`GetMapBoxLayerName`

**功能**：获取图廓层名。

**语法**：

```vbs
SSProcess.GetMapBoxLayerName ()
```

**说明**：

执行此函数将返回图廓层名称。

**示例**：

```vbs
下面示例将返回当前工程的图廓层层名称。
Sub OnClick()
a=SSProcess.GetMapBoxLayerName
End Sub
```


#### SetMapBoxWay 获取分幅方式

**函数**：`SetMapBoxWay`

**功能**：GetMapBoxWay函数

**语法**：

```vbs
SSProcess.GetMapBoxWay()
```

**说明**：

执行此函数将返回分幅方式：0.矩形分幅；1.梯形分幅；2.斜分幅

**示例**：

```vbs
下面示例将返回当前工程的分幅方式。
Sub OnClick()
a=SSProcess.GetMapBoxWay
End Sub
```


#### GetMapRowCol 获取图幅格式

**函数**：`GetMapRowCol`

**功能**：获取图幅格式。

**语法**：

```vbs
SSProcess.SetMapRowCol(row, col)
```

**参数**：

- row 矩形分幅规格高度
- col 矩形分幅规格宽度

**说明**：

该函数设置图幅规格。

**示例**：

```vbs
下面示例将获得当前工程的矩形分幅规格高度和宽度值。
Sub OnClick()
SSProcess.GetMapRowCol row ,col
End Sub
```


#### GetMapScale 获取地图比例尺

**函数**：`GetMapScale`

**功能**：获取地图比例尺。

**语法**：

```vbs
SSProcess.GetMapScale()
```

**参数**：

- 返回地图比例尺分母。

**说明**：

执行该函数，返回系统比例尺，本比例尺为软件设置菜单->系统环境设->图幅信息中的比例尺。

**示例**：

```vbs
下面获取系统比例尺：
Sub OnClick()
Dim MapScale
MapScale= SSProcess.GetMapScale
Msgbox MapScale
End Sub
```


#### GetMapStatus 获取数据库状态

**函数**：`GetMapStatus`

**功能**：获取数据库状态。

**语法**：

```vbs
SSProcess.GetMapStatus(type)
```

**参数**：

- type 数据库状态 2（锁定数据库）、4（锁定地物显示列表）、8（锁定标注显示列表）、16（锁定索引）、32(锁定Undo栈)

**说明**：

执行该函数，可以获取数据库的数据库锁定状态、地物显示列表锁定状态、标注显示列表锁定状态和索引锁定状态。
返回 0为未锁定，返回 1为锁定。

**示例**：

```vbs
下面示例获取数据库锁定状态：
Sub OnClick()
Dim status
status = SSProcess.GetMapStatus (2）
End Sub
下面示例获取地物显示列表锁定状态：
Sub OnClick()
status = SSProcess.GetMapStatus (4）
End Sub
```


#### GetNoteTemplateTB 获取系统注记模板表

**函数**：`GetNoteTemplateTB`

**功能**：获取系统注记模板表名称。

**语法**：

```vbs
SSProcess.GetNoteTemplateTB()
```

**说明**：

执行该函数，将返回当前系统使用的注记分类表名称。

**示例**：

```vbs
下面示例获取当前使用的注记分类表名称：
Sub OnClick()
Dim templateName
templateName = SSProcess.GetNoteTemplateTB()
End Sub
```


#### GetPrivateProfile 读取自定义 Ini 文件参数

**函数**：`GetPrivateProfile`

**功能**：读取自定义Ini文件参数。

**语法**：

```vbs
SSProcess.GetPrivateProfile(strSection, strKey , strDefault, strInIfile)
```

**参数**：

- strSection 参数组名（字符串类型）
- strKey 变量名（字符串类型）
- strDefault 默认值（字符串类型）
- strInIfile 自定义Ini文件的路径（字符串类型）

**说明**：

该函数获取自定义ini文件中的指定参数值。

**示例**：

```vbs
Sub OnClick()
strSection ="ArcGISVesion" :strKey="esriVersion" :strDefault="wh123"
strInIfile ="D:\my_Ini.ini"
Msgbox strKey& "="&SSProcess.GetPrivateProfile (strSection, strKey,strDefault ,strInIfile)
End Sub
```


#### GetSelectionObjExternalPolygon 获取选择集对象的最小外接多边形

**函数**：`GetSelectionObjExternalPolygon`

**功能**：获取选择集对象的最小外接多边形。

**语法**：

```vbs
SSProcess.GetSelectionObjExternalPolygon(gridSize)
```

**参数**：

- gridSize 网格大小

**说明**：

该函数用于获取选择集对象的最小外接多边形。

**示例**：

```vbs
'当前图幅设定
Sub OnClick()
Dim xycoords(10000), icount
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition”SSObj_Type”, "==","POINT,LINE,AREA"
SSProcess.SelectFilter
GeoCount= SSProcess.GetSelGeoCount()
If GeoCount=0 Then
Exit Sub
'获得数据的外接多边形
gridSize= 10
strCoord= SSProcess.GetSelectionObjExternalPolygon(gridSize )
SSFunc.ScanString strCoord,",",xycoords, icount
'添加编码为3的临时要素，外接多边形
SSProcess.CreateNewObjByCode 3
For i =0 To icount-1 step 2
x =xycoords( i)
y =xycoords( i +1)
SSProcess.AddNewObjPoint xycoords(i),xycoords (i + 1),0,0, ""
Next
SSProcess.AddNewObjToSaveObjList
'临时要素最大ID
GeoMaxID= SSProcess.GetGeoMaxID
'获取数据的中心坐标
AreaLabelX= SSProcess.GetObjectAttr(GeoMaxID ,"SSObj_AreaLabelX" )
AreaLabelY= SSProcess.GetObjectAttr(GeoMaxID ,"SSObj_AreaLabelY" )
SSProcess.SaveBufferObjToDatabase
'删除临时要素
SSProcess.DeleteObject GeoMaxID
'设置当前图幅
SSProcess.SetCurMapFrame AreaLabelX,AreaLabelY,0, ""
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
End Sub
```


#### GetSymbolScriptTB 获取系统符号描述表

**函数**：`GetSymbolScriptTB`

**功能**：获取系统符号描述表名称。

**语法**：

```vbs
SSProcess.GetSymbolScriptTB()
```

**说明**：

执行该函数，将返回当前系统使用的符号描述表名称。

**示例**：

```vbs
下面示例获取当前使用的符号描述表名称：
Sub OnClick()
Dim symbolName
symbolName = SSProcess.GetSymbolScriptTB()
End Sub
```


#### KillTimer 删除定时器

**函数**：`KillTimer`

**功能**：删除定时器。

**语法**：

```vbs
SSProcess.KillTimer(nIDEvent)
```

**参数**：

- nlDEvent 定时器的序号

**示例**：

```vbs
scroll = 1
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
title="节目号抽奖程序"
SSProcess.ShowScriptDlg mode,title
scroll = 1
SSView.ZoomExtend 0, 0 , 50, 50 , 1
SSProcess.SetTimer 128801, 2
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
SSProcess.KillTimer 128801
End Sub
Function OnLButtonUp(x , y, spx , spy, flags)
If scroll = 1 Then
scroll = 0
Else
scroll = 1
End If
End Function
Function Draw( pDc )
SSView.GetClientArea minX, minY, maxX , maxY
x = (minX + maxX)/2.0
y = (minY + maxY)/ 2.0
text = FormatNumber(times , 0)
nOldLogicop = SSView.SetROP2 ( 13)
xoff = 18
yoff = 12
pointStrings = x-xoff & "," & y-yoff & "," & x+xoff & "," & y-yoff & "," & x+xoff & "," & y+yoff & "," & x-xoff & "," & y+yoff & "," & x-xoff & "," & y -yoff
SSView.DrawPolyline 0, pointStrings, 11513775 , 100, 1
x = x + 0.25
y = y - 0.25
pointStrings = x-xoff & "," & y-yoff & "," & x+xoff & "," &y-yoff& "," & x+xoff & "," & y+yoff & "," & x-xoff & "," & y+yoff & "," & x-xoff & "," & y -yoff
SSView.DrawPolyline 0, pointStrings, 11511775 , 100, 1
x = x + 0.25
y = y - 0.25
pointStrings = x-xoff & "," & y-yoff & "," & x+xoff & "," & y-yoff & "," & x+xoff & "," & y+yoff & "," & x-xoff & "," & y+yoff & "," & x-xoff & "," & y -yoff
SSView.DrawPolygon 0, pointStrings, 16646143 , 1, 1 , 16777215, 0
x = x - 0.5
y = y + 0.5
SSView.DrawText 0, x, y , times, 0 , "黑体", 256000, 6000, 6000, 0 , 0, 0, 0 , 500, 0 , 0, 1
SSView.DrawText 0, x+ 0.5, y- 0.5, times, 0 , "黑体", 25500, 6000, 6000, 0, 0 , 0, 0 , 500, 0 , 0, 0
SSView.DrawText 0, x+ 1, y- 1, times, 0 , "黑体", 255, 6000, 6000, 0 , 0, 0, 0 , 500, 0 , 0, 0
SSView.SetROP2 nOldLogicop
End Function
Function OnDraw(pDc)
Draw pDC
End Function
times = 0
Function OnTimer(nIDEvent)
If scroll=1 Then
Draw 0
times = times + 1
If times > 55 Then times = 1
End If
End Function
```


#### MoveScreen 移动屏幕到指定位置

**函数**：`MoveScreen`

**功能**：移动屏幕到指定位置。

**语法**：

```vbs
SSProcess.MoveScreen(x, y, movealways)
```

**参数**：

- x 将屏幕移动到的x坐标，以此x值为屏幕中心x坐标
- y 将屏幕移动到的y坐标 ,以此y值为屏幕中心y坐标
- movealways 当指定的x ,y坐标已落在当前窗口内，movealways =0 不执行移屏，movealways =1时则总是移屏到指定x ,y坐标

**说明**：

使用此函数将屏幕移动到指定的位置。

**示例**：

```vbs
下面示例将屏幕移动到坐标为（2,2)的位置。
Sub OnClick()
x=2:y= 2:movealways= 1
SSProcess.MoveScreen x ,y,movealways
End Sub
```


#### ReadEpsDBIni 读取 EPS 工程参数

**函数**：`ReadEpsDBIni`

**功能**：读取EPS工程参数。

**参数**：

- strSection 参数组名
- strKey 变量名
- strDefault 值

**说明**：

该函数获取eps工程中的指定参数值。

**示例**：

```vbs
Sub OnClick()
strSection ="Projection" :strKey="EllipseName" :strDefault="0"
strValue =SSProcess.ReadEpsDBIni (strSection, strKey , strDefault )
Msgbox "strValue="&strValue
End Sub
```


#### ReadEpsGlobalIni 读取 EPSGlobal.Ini 文件参数

**函数**：`ReadEpsGlobalIni`

**功能**：读取EpsGlobalIni文件参数。

**语法**：

```vbs
SSProcess.ReadEpsGlobalIni(strSection, strKey, strDefault)
```

**参数**：

- strSection 分类名
- strKey 关键字段
- strDefault 缺省值

**示例**：

```vbs
获取当前工程使用的台面名称。
Sub OnClick()
Name = SSProcess.ReadEpsGlobalIni ("System" , "CurrentSchemeName" , "1" )
Msgbox Name
End Sub
```


#### ReadEpsIni 读取 EPS.Ini 文件参数

**函数**：`ReadEpsIni`

**功能**：读取EPS.Ini文件参数。

**语法**：

```vbs
SSProcess.ReadEpsIni(strSection, strKey,strDefault)
```

**参数**：

- strSection 段落名称，指读取INI文件的位置名称，当strSection=“测试”后，读取内容为[测试]
- strKey 关键字
- strDefault 默认值，一般为空

**说明**：

执行该函数，读取台面目录下EPS.INI文件中相关内容。

**示例**：

```vbs
下面示例将读取台面下EPS.INI文件中[测试]段落中LastAttr的值：
Sub OnClick()
strvalue=SSProcess.ReadEpsIni("测试", "LastAttr" ,"")
Msgbox strvalue
End Sub
```


#### ReadEpsTemplateIni 读取 EPS 模板参数

**函数**：`ReadEpsTemplateIni`

**功能**：读取EPS模板参数。

**语法**：

```vbs
SSProcess.ReadEpsTemplateIni(strSection, strKey , strDefault)
```

**参数**：

- strSection 分类名
- strKey 关键字段
- strDefault 缺省值

**说明**：

执行该函数，获取EPS模板中的INI参数值。

**示例**：

```vbs
Sub OnClick()
SSProcess. WriteEpsTemplateIni "epsscale","scale1" ,"500"
strValue = SSProcess.ReadEpsTemplateIni("epsscale","scale1", strDefault)
Msgbox strValue
End Sub
```


#### ReadEpsXMLIni 读取 EPS XML 文件参数

**函数**：`ReadEpsXMLIni`

**功能**：读物EPS XML文件参数。

**语法**：

```vbs
SSProcess.ReadEpsXMLIni(nSettingFile, strSection, strKey , strDefault)
```

**参数**：

- nSettingFile
- strSection 分类名
- strKey 关键字段
- strDefault 缺省值

**说明**：

执行该函数，读取EPSXML文件参数。


#### RefreshView 刷新图形窗口

**函数**：`RefreshView`

**功能**：刷新图形窗口。

**语法**：

```vbs
SSProcess.RefreshView()
```

**说明**：

执行该函数，可以刷新当前图形窗口。


#### SendFrameMessage 发送主框架窗口消息

**函数**：`SendFrameMessage`

**功能**：发送主框架窗口消息。

**语法**：

```vbs
SSProcess.SendFrameMessage(Msg, wParam, lParam)
```

**参数**：

- Msg 主框架窗口内容
- wParam 主框架窗口宽度
- lParam 主框架窗口长度

**说明**：

该函数为系统设置函数，执行该函数可发送主框架窗口消息。


#### SendViewMessage 发送视图窗口消息

**函数**：`SendViewMessage`

**功能**：发送视图窗口消息。

**语法**：

```vbs
SSProcess.SendViewMessage(Msg, wParam, lParam)
```

**参数**：

- Msg 视图窗口内容
- wParam 视图窗口宽度
- lParam 视图窗口长度

**说明**：

该函数为系统设置函数，执行该函数可发送视图窗口消息。


#### SetColorInfoTB 设置系统颜色表

**函数**：`SetColorInfoTB`

**功能**：设置系统颜色表。

**语法**：

```vbs
SSProcess.SetColorInfoTB(colorInfoTB)
```

**参数**：

- colorInfoTB 系统注记表名称


#### SetDealMsg 设置当前处理操作进程提示

**函数**：`SetDealMsg`

**功能**：设置当前处理操作进程提示。

**语法**：

```vbs
SSProcess.SetDealMsg(dealMsg)
```

**参数**：

- dealMsg 提示标题

**示例**：

```vbs
Sub OnClick()
SSProcess.EpsProgressSetStyle 1
SSProcess.EpsProgressCreate 1000, "进度条"
SSProcess.SetDealMsg "dealMsg"
For i = 0 To 1000
SSProcess.EpsProgressSetPos i
Next
SSProcess.EpsProgressDelete
End Sub
```


#### SetFeatureCodeInfo 设置编码信息

**函数**：`SetFeatureCodeInfo`

**功能**：设置编码信息。

**语法**：

```vbs
SSProcess.SetFeatureCodeInfo(code, fields, values)
```

**参数**：

- code 编码
- Fields 添加数据的字段名称列表 ,字段间用逗号分隔
- Values 添加数据的字段值列表 ,值间用逗号分隔

**示例**：

```vbs
Sub OnClick()
SSProcess.SetFeatureCodeInfo "3103015", "LayerName", "设施面"
SSProcess.UpdateFeatureCode
End Sub
```


#### SetFeatureCodeTB 设置系统编码表

**函数**：`SetFeatureCodeTB`

**功能**：设置系统编码表。

**语法**：

```vbs
SSProcess.SetFeatureCodeTB(featureTB,symbolscriptTB)
```

**参数**：

- featureTB 编码表名称
- symbolscriptTB 符号描述表名称

**说明**：

执行该函数，将把指定的编码表和符号描述表设为当前系统使用的编码模板表。

**示例**：

```vbs
下面示例设置FeatureCodeTB_500和SymbolScriptTB_500为当前使用的编码模板表：
Sub OnClick()
SSProcess.SetFeatureCodeTB "FeatureCodeTB_500", "SymbolScriptTB_500"
End Sub
```


#### SetFontClassInfo 设置分类号信息

**函数**：`SetFontClassInfo`

**功能**：设置分类号信息。

**语法**：

```vbs
SSProcess.SetFontClassInfo(code, fields, values)
```

**参数**：

- code 分类号编码 ,可查看NoteTemplateTB中的分类号
- fields 字段名称，可查看NoteTemplateTB的英文字段名称
- values 修改后的值

**说明**：

使用此函数将分类号进行修改

**示例**：

```vbs
下面示例将分类号为2990001的编码的层名临时修改为水系注记
Sub OnClick()
SSProcess.SetFontClassInfo "2990001", "LayerName" , "水系注记"
SSProcess.UpdateFontClass
End Sub
```


#### SetFrameCode 设置图廓编码

**函数**：`SetFrameCode`

**功能**：设置图廓编码。

**语法**：

```vbs
SSProcess.SetFrameCode(frmCode)
```

**参数**：

- frmCode 图廓编码

**说明**：

使用此函数将设置工程中的图廓编码。

**示例**：

```vbs
下面示例工程中的图廓编码修改为999900。
Sub OnClick()
SSProcess.SetFrameCode 999900
End Sub
```


#### SetGridCellInfo 设置多列参数编辑框单元格信息

**函数**：`SetGridCellInfo`

**功能**：设置多列参数编辑框单元格信息。

**语法**：

```vbs
SSProcess.SetGridCellInfo(row, col, value, options)
```

**参数**：

- row 第行数
- rol 第列数
- value 值
- options 可选内容，默认值

**说明**：

使用此函数将设置编辑框中某单元格的内容及可选值。

**示例**：

```vbs
下面示例将按设置带有标题、内容的对话框
Sub OnClick()
SSProcess.SetGridHeadInfo "姓名,性别,学历,工作年限（年)" , "100,50,100,100" , 0, 1, 0
SSProcess.SetGridCellInfo 1, 1 , "张三", ""
SSProcess.SetGridCellInfo 1, 2 , "男", "男,女"
SSProcess.SetGridCellInfo  1, 3 , "本科", ""
SSProcess.SetGridCellInfo 1, 4 ,  "3", ""
SSProcess.ShowGridEditDlg "查询统计结果"
End Sub
```


#### SetGridHeadInfo 设置多列参数编辑框表格信息

**函数**：`SetGridHeadInfo`

**功能**：设置多列参数编辑框表格信息。

**语法**：

```vbs
SSProcess.SetGridHeadInfo(columnNames, columnWidths, enableMarkSortedColumn, enableHeader, emptyRow)
```

**参数**：

- columnNames 列名称，多个用逗号隔开
- columnWidths 列宽度，多个用逗号隔开，个数与columnNames一致
- enableMarkSortedColumn 是否允许列排序
- enableHeader 是否隐藏标题，0是，1否
- emptyRow 默认空行数

**示例**：

```vbs
下面示例将按设置带有标题、内容的对话框
Sub OnClick()
SSProcess.SetGridHeadInfo "姓名,性别,学历,工作年限(年）" , "100,50,100,100" , 0, 1, 0
SSProcess.SetGridCellInfo 1, 1 , "张三", ""
SSProcess.SetGridCellInfo 1, 2 , "男", "男,女"
SSProcess.SetGridCellInfo  1, 3 , "本科", ""
SSProcess.SetGridCellInfo 1, 4 ,  "3", ""
SSProcess.ShowGridEditDlg "查询统计结果"
End Sub
```


#### SetInputParameter 修改录入参数

**函数**：`SetInputParameter`

**功能**：修改录入参数。

**语法**：

```vbs
SSProcess.SetInputParameter(name, value)
```

**参数**：

- name 对话框录入参数名称
- value 参数值

**说明**：

使用此函数将在设置指定对话框录入名称的参数值。

**示例**：

```vbs
下面示例将对话框中的显示类型值修改为按功能显示。
Sub OnClick()
mode =1’=0 无参数对话框 =1 有参数对话框
title= "图元分色检查："
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "显示类型" ,"按户显示",0,"按户显示,按功能显示" , ""
SSProcess.ShowScriptDlg mode ,title
SSProcess.SetInputParameter "显示类型" ,"按功能显示"
SSProcess.ShowScriptDlg mode ,title
End Sub
```


#### SetLayerStatus 设置图层状态

**函数**：`SetLayerStatus`

**功能**：设置图层状态。

**语法**：

```vbs
SSProcess.SetLayerStatus(layername,flags,type)
```

**参数**：

- Layername 要设置的层名
- flags 设置或取消状态 0（取消状态）、1（ 设置状态）
- type 用户层状态 1(是否可见) 2（是否锁定编辑）、4（是否锁定索引）、8（是否已拓扑）、16（是否已做悬挂点处理）

**说明**：

执行该函数，可以设置或取消用户层的各种状态。

**示例**：

```vbs
下面示例关闭指定用户层：
Sub OnClick()
SSProcess.SetLayerStatus "控制点", 0, 1
End Sub
下面示例打开指定用户层：
Sub OnClick()
SSProcess.SetLayerStatus "控制点", 1, 1
End Sub
```


#### SetMapBoxLayerName 设置图廓层名

**函数**：`SetMapBoxLayerName`

**功能**：设置图廓层名。

**语法**：

```vbs
SSProcess.SetMapBoxLayerName(layerName)
```

**参数**：

- layerName层名

**说明**：

使用此函数设置当前工程中的图廓层名。

**示例**：

```vbs
Sub OnClick()
SSProcess.SetMapBoxLayerName "图廓层"
End Sub
```


#### SetMapBoxWay 设置分幅方式

**函数**：`SetMapBoxWay`

**功能**：设置分幅方式。

**语法**：

```vbs
SSProcess.SetMapBoxWay(way)
```

**参数**：

- way分幅方式（0：矩形分幅；1：梯形分幅；2：斜分幅；3：北京地籍分幅）

**说明**：

使用此函数设置分幅方式。

**示例**：

```vbs
Sub OnClick()
way =3
SSProcess.SetMapBoxWay way
End Sub
```


#### SetMapRowCol 设置图幅规格

**函数**：`SetMapRowCol`

**功能**：设置图幅规格。

**语法**：

```vbs
SSProcess.SetMapRowCol(row, col)
```

**参数**：

- row 行高(单位cm)
- col 列宽(单位cm)

**说明**：

该函数设置图幅规格。

**示例**：

```vbs
Sub OnClick()
SSProcess.SetMapRowCol 100,100
End Sub
```


#### SetMapScale 设置地图比例尺

**函数**：`SetMapScale`

**功能**：设置地图比例尺。

**语法**：

```vbs
SSProcess.SetMapScale(scale)
```

**参数**：

- scale 比例尺（如：500）

**说明**：

该函数设置地图比例尺。
示例
Sub OnClick()
SSProcess.SetMapScale "500"
End Sub


#### SetMapStatus 设置数据库状态

**函数**：`SetMapStatus`

**功能**：设置数据库状态。

**语法**：

```vbs
SSProcess.SetMapStatus(flags,type)
```

**参数**：

- flags 设置或取消状态 0（取消状态）、1（ 设置状态）
- type 数据库状态 2（锁定数据库）、4（锁定地物显示列表）、8（锁定标注显示列表）、16（锁定索引）、32(锁定Undo栈)

**说明**：

执行该函数，可以设置或取消数据库的数据库锁定状态、地物显示列表锁定状态、标注显示列表锁定状态和索引锁定状态。

**示例**：

```vbs
下面示例设置取消数据库锁定状态：
Sub OnClick()
SSProcess.SetMapStatus 0, 2
End Sub
下面示例设置数据库锁定状态：
Sub OnClick()
SSProcess.SetMapStatus 1, 2
End Sub
```


#### SetNotetemplateTB 设置系统注记模板表

**函数**：`SetNotetemplateTB`

**功能**：设置系统注记模板表。

**语法**：

```vbs
SSProcess.SetNotetemplateTB(notetemplateTB)
```

**参数**：

- notetemplateTB 注记分类表名称

**说明**：

执行该函数，将把指定的注记分类表设为当前系统使用的注记模板表。

**示例**：

```vbs
下面示例设置NoteTemplateTB_500为当前使用注记模板表：
Sub OnClick()
SSProcess.SetNotetemplateTB "NoteTemplateTB_500"
End Sub
```


#### SetTimer 设置定时器

**函数**：`SetTimer`

**功能**：设置定时器。

**语法**：

```vbs
SSProcess.SetTimer(nIDEvent, nElapse)
```

**参数**：

- nIDEvent 定时器标示
- nElapse 定时时间。

**说明**：

此函数功能为设置定时器，需与KillTimer函数同时使用。
执行后将执行定时器。

**示例**：

```vbs
scroll = 1
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
title="节目号抽奖程序"
SSProcess.ShowScriptDlg mode,title
scroll = 1
SSView.ZoomExtend 0, 0 , 50, 50 , 1
SSProcess.SetTimer 128801, 2
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
SSProcess.KillTimer 128801
End Sub
Function OnLButtonUp(x , y, spx , spy, flags)
If scroll = 1 Then
scroll = 0
Else
scroll = 1
End If
End Function
Function Draw( pDc )
SSView.GetClientArea minX, minY, maxX , maxY
x = (minX + maxX)/2.0
y = (minY + maxY)/ 2.0
text = FormatNumber(times , 0)
nOldLogicop = SSView.SetROP2 ( 13)
xoff = 18
yoff = 12
pointStrings = x-xoff & "," & y-yoff & "," & x+xoff & "," & y-yoff & "," & x+xoff & "," & y+yoff & "," & x-xoff & "," & y+yoff & "," & x-xoff & "," & y -yoff
SSView.DrawPolyline 0, pointStrings, 11513775 , 100, 1
x = x + 0.25
y = y - 0.25
pointStrings = x-xoff & "," & y-yoff & "," & x+xoff & "," &y-yoff& "," & x+xoff & "," & y+yoff & "," & x-xoff & "," & y+yoff & "," & x-xoff & "," & y -yoff
SSView.DrawPolyline 0, pointStrings, 11511775 , 100, 1
x = x + 0.25
y = y - 0.25
pointStrings = x-xoff & "," & y-yoff & "," & x+xoff & "," & y-yoff & "," & x+xoff & "," & y+yoff & "," & x-xoff & "," & y+yoff & "," & x-xoff & "," & y -yoff
SSView.DrawPolygon 0, pointStrings, 16646143 , 1, 1 , 16777215, 0
x = x - 0.5
y = y + 0.5
SSView.DrawText 0, x, y , times, 0 , "黑体", 256000, 6000, 6000, 0 , 0, 0, 0 , 500, 0 , 0, 1
SSView.DrawText 0, x+ 0.5, y- 0.5, times, 0 , "黑体", 25500, 6000, 6000, 0, 0 , 0, 0 , 500, 0 , 0, 0
SSView.DrawText 0, x+ 1, y- 1, times, 0 , "黑体", 255, 6000, 6000, 0 , 0, 0, 0 , 500, 0 , 0, 0
SSView.SetROP2 nOldLogicop
End Function
Function OnDraw(pDc)
Draw pDC
End Function
times = 0
Function OnTimer(nIDEvent)
If scroll=1 Then
Draw 0
times = times + 1
If times > 55 Then times = 1
End If
End Function
```


#### SetUnfilterObjColor 设置未可控对象颜色状态

**函数**：`SetUnfilterObjColor`

**功能**：设置未可控对象颜色状态。

**语法**：

```vbs
SSProcess.SetUnfilterObjColor(flag, color)
```

**参数**：

- flag 是否对不在有效设置控制之内的对象集合重配色 0（不配）、1（配）
- color 重新配色的颜色值,分COLORNO(颜色号)、颜色值和RGB(红,绿，蓝)三种指定方式

**说明**：

执行该函数，可实现不在有效设置控制之内的对象集合重新配色显示.

**示例**：

```vbs
下面示例设置不在有效设置控制之内的对象集合的显示颜色：
Sub OnClick()
SSProcess.ClearDispFilterInfo
SSProcess.AddDispFilterInfo 0, 0, 1, 0, "RGB(255,0,0)", "2010,3010"
SSProcess.SetUnfilterObjColor 1, "RGB(0,255,0)"
End Sub
```


#### ShowGridEditDlg 显示多列参数编辑框对话框

**函数**：`ShowGridEditDlg`

**功能**：显示多列参数编辑框对话框。

**语法**：

```vbs
SSProcess.ShowGridEditDlg(title)
```

**参数**：

- title 对话框标题

**说明**：

该函数用于显示多列参数编辑对话框。

**示例**：

```vbs
Sub OnClick()
SSProcess.SetGridHeadInfo "姓名,性别,学历,工作年限（年）", "100,50,100,100", 0, 1, 0
SSProcess.SetGridCellInfo 1, 1 , "张三", ""
SSProcess.SetGridCellInfo 1, 2 , "男", "男,女"
SSProcess.SetGridCellInfo  1, 3 , "本科", ""
SSProcess.SetGridCellInfo 1, 4 , "3", ""
SSProcess.ShowGridEditDlg "查询统计结果"
End Sub
```


#### ShowGridEditDlg1 显示多列参数编辑框对话框

**函数**：`ShowGridEditDlg1`

**功能**：显示多列参数编辑框对话框。

**语法**：

```vbs
SSProcess.ShowGridEditDlg1(title, okBtnTitle, clickOKCloseDlg, dlgWidth, dlgHeight)
```

**参数**：

- title 编辑框的标题
- okBtnTitle 确定或者取消
- clickOKCloseDlg 0, 取消编辑框操作 1, 执行编辑框内操作
- dlgWidth 编辑框的宽
- dlgHeight 编辑框的高

**示例**：

```vbs
Sub OnClick()
title= "山维科技"
okBtnTitle= "确定"
clickOKCloseDlg=1
dlgWidth= 200
dlgHeight= 200
SSProcess.ShowGridEditDlg1 title, okBtnTitle, clickOKCloseDlg , dlgWidth, dlgHeight
End Sub
```


#### ShowGridEditDlg2 显示多列参数编辑框对话框（有模）

**函数**：`ShowGridEditDlg2`

**功能**：显示多列参数编辑框对话框（有模）。

**语法**：

```vbs
SSProcess.ShowGridEditDlg2(title, okBtnTitle, clickOKCloseDlg, dlgWidth, dlgHeight)
```

**参数**：

- title 编辑框的标题
- okBtnTitle 确定或者取消
- clickOKCloseDlg 0, 取消编辑框操作 1, 执行编辑框内操作
- dlgWidth 编辑框的宽
- dlgHeight 编辑框的高


#### ShowInputParameterDlg 显示参数录入对话框

**函数**：`ShowInputParameterDlg`

**功能**：显示参数录入对话框。

**语法**：

```vbs
SSProcess.ShowInputParameterDlg(title)
```

**参数**：

- title 对话框标题

**说明**：

若点击“确定”，返回值为1；若点击“取消”，返回值为0。

**示例**：

```vbs
Sub OnClick()
result =SSProcess.ShowInputParameterDlg ("坐标导入")
End Sub
```


#### ShowInputUserDefParameterDlg 显示自定义参数录入对话框

**函数**：`ShowInputUserDefParameterDlg`

**功能**：显示自定义参数录入对话框。

**语法**：

```vbs
SSProcess.ShowInputUserDefParameterDlg(title, dlgTemplateName, dlgWidth, dlgHeight, colCount, titleWidth, valueWidth)
```

**参数**：

- title 录入对话框标题
- dlgTemplateName 当前工程模板所对应文件夹的属性表.dlg文件
- dlgWidth 录入对话框的宽
- dlgHeight 录入对话框的高
- colCount 对话框内显示的列
- titleWidth 标题宽
- valueWidth 对话框内宽

**示例**：

```vbs
Sub OnClick()
title = "图廓属性表"
dlgTemplateName = "图廓属性表"
dlgWidth = 500
dlgHeight = 500
colCount = 2
titleWidth = 50
valueWidth = 300
SSProcess.ShowInputUserDefParameterDlg title , dlgTemplateName,dlgWidth, dlgHeight, colCount , titleWidth, valueWidth
End Sub
```


#### UpdateFeatureCode 更新编码定义信息到数据库

**函数**：`UpdateFeatureCode`

**功能**：更新编码定义信息到数据库。

**语法**：

```vbs
SSProcess.UpdateFeatureCode(Code,Fileds,Values)
```

**参数**：

- Code 编码
- Fields 添加数据的字段名称列表 ,字段间用逗号分隔
- Values 添加数据的字段值列表 ,值间用逗号分隔

**说明**：

执行该函数，将通过选择集设定的更新编码定义信息更新到数据库
使用此函数将分类号进行修改

**示例**：

```vbs
下面示例将编码为3103015 的地物的层名临时修改为设施面：
Sub OnClick()
SSProcess.SetFeatureCodeInfo "3103015", "LayerName", "设施面"
SSProcess.UpdateFeatureCode
End Sub
```


#### UpdateFontClass 更新注记分类信息到数据库

**函数**：`UpdateFontClass`

**功能**：更新注记分类信息到数据库。

**语法**：

```vbs
SSProcess.UpdateFontClass()
```

**说明**：

执行该函数，将通过选择集设定的注记分类信息更新到数据库。

**示例**：

```vbs
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==" , "NOTE"
SSProcess.SelectFilter
notecount = SSProcess.GetSelNoteCount
For i= 0 To  notecount -1
SSProcess.SetSelNoteValue i , "SSObj_FontClass", "1"
Next
SSProcess.UpdateFontClass '更新存储到数据库
End Sub
```


#### WriteEpsDBIni 写 EPS 工程参数

**函数**：`WriteEpsDBIni`

**功能**：写EPS工程参数。

**语法**：

```vbs
SSProcess.WriteEpsDBIni(strSection, strKey , strValue)
```

**参数**：

- strSection 参数组名
- strKey 变量名
- strValue 值

**说明**：

使用此函数将在eps工程文件中IniInfoTB表写工程参数。

**示例**：

```vbs
下面示例将向IniInfoTB.Ini文件内写组名为epsscale，变量名为scale1，值为500的参数。
Sub OnClick()
SSProcess.WriteEpsDBIni "epsscale","scale1" ,"500"
End Sub
```


#### WriteEpsGlobalIni 写 EPS Global.Ini 文件参数

**函数**：`WriteEpsGlobalIni`

**功能**：写EPS Global.Ini文件参数。

**语法**：

```vbs
SSProcess.WriteEpsGlobalIni(strSection, strKey , strValue)
```

**参数**：

- strSection 参数组名
- strKey 变量名
- strValue 值

**说明**：

使用此函数将在eps软件目录下的EPSGlobal.Ini文件写相应参数。

**示例**：

```vbs
下面示例将向EPSGlobal.Ini文件内写文件参数
Sub OnClick()
SSProcess.WriteEpsGlobalIni "strSection","strKey" ,"strValue"
End Sub
```


#### WriteEpsIni 写 EPS.Ini 文件参数

**函数**：`WriteEpsIni`

**功能**：写EPS.Ini文件参数。

**语法**：

```vbs
SSProcess.WriteEpsIni(strSection, strKey , strValue)
```

**参数**：

- strSection 段落名称，指写入到INI文件后的位置名称，当strSection=“测试”后，写入内容为[测试]。
- strKey 关键字
- strValue 关键字的值。

**说明**：

执行该函数，向软件使用的台面目录下EPS.INI文件中写入内容。

**示例**：

```vbs
下面示例将在台面下的EPS.INI文件内写入如下内容：
[测试]
LastAttr=测试
代码：
Sub OnClick()
SSProcess.WriteEpsIni "测试", "LastAttr" ,"测试"
End Sub
```


#### WriteEpsTemplateIni 写 EPS 模板参数

**函数**：`WriteEpsTemplateIni`

**功能**：写EPS模板参数。

**语法**：

```vbs
SSProcess.WriteEpsTemplateIni(strSection, strKey , strValue)
```

**参数**：

- strSection 参数组名
- strKey 变量名
- strValue 值

**说明**：

使用此函数向当前工程所用mdt模板文件中的IniInfoTB表中写参数。

**示例**：

```vbs
下面示例将向模板中IniInfoTB表内写组名为epsscale，变量名为scale1，值为500的参数：
Sub OnClick()
SSProcess. WriteEpsTemplateIni "epsscale","scale1" ,"500"
End Sub
```


#### WriteEpsXMLIni 写 EPS XML 文件参数

**函数**：`WriteEpsXMLIni`

**功能**：写EPS XML文件参数。

**语法**：

```vbs
SSProcess.WriteEpsXMLIni(nSettingFile,strSection,strKey,strValue)
```

**参数**：

- nSettingFile
- strSection 分类名
- strKey 关键字段
- strValue 值

**说明**：

执行该函数，写EPSXML文件参数。

**示例**：

```vbs
Sub OnClick()
nSettingFile = 0
strSection = "HouseSetup\DatabaseType"
strKey ="default"
strValue = "HoenKing"
SSProcess.WriteEpsXMLIni nSettingFile,strSection,strKey,strValue
aa = SSProcess.ReadEpsXMLIni(nSettingFile,strSection,strDefault)
Msgbox aa
End Sub
```


#### WritePrivateProfile 写自定义 Ini 文件参数

**函数**：`WritePrivateProfile`

**功能**：写自定义Ini文件参数。

**语法**：

```vbs
SSProcess.WritePrivateProfile(strSection, strKey , strValue, strInIfile)
```

**参数**：

- strSection 参数组名
- strKey 变量名
- strValue 值
- strInIfile 自定义ini文件路径

**说明**：

使用此函数将在自定义的Ini文件中写相应参数。

**示例**：

```vbs
Sub OnClick()
strInIfile=SSProcess.GetSysPathName(0)&"/myini.ini"
strSection="System"
strKey="UserTemplate"
strValue="1"
SSProcess.WritePrivateProfile strSection, strKey, strValue, strInIfile
End Sub
```


---

### 选择集操作

#### ClearSelectCondition 清空选择条件

**函数**：`ClearSelectCondition`

**功能**：ClearSelectCondition 函数

**语法**：

```vbs
SSProcess.ClearSelectCondition()
```

**说明**：

执行该函数，将清空函数 SetSelectCondition 设定的选择条件

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelectCondition
End Sub
```


#### ClearSelectConditionGroup 清除指定分组选择条件

**函数**：`ClearSelectConditionGroup`

**功能**：清除函数 SetSelectConditionGroup 设定的分组选择条件。

**语法**：

```vbs
SSProcess.ClearSelectConditionGroup(groupName)
```

**参数**：

- groupName 要清除的分组条件名称

**说明**：

执行该函数，将清空函数 SetSelectConditionGroup 设定的分组名为groupName的选择条件。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelectConditionGroup "房屋范围线"
End Sub
```


#### ClearSelectConditionGroups 清除所有分组选择条件

**函数**：`ClearSelectConditionGroups`

**功能**：清空函数 SetSelectConditionGroup 设定的所有分组选择条件。

**语法**：

```vbs
SSProcess.ClearSelectConditionGroups()
```

**说明**：

执行该函数，将清空函数 SetSelectConditionGroup 设定的所有分组选择条件

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelectConditionGroups
End Sub
```


#### ClearSelection 清空选择集

**函数**：`ClearSelection`

**功能**：ClearSelection 函数

**语法**：

```vbs
SSProcess.ClearSelection()
```

**说明**：

执行该函数，将清空当前选择集中的所有对象

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
End Sub
```


#### ClearSysSelection 清空系统选择集

**函数**：`ClearSysSelection`

**功能**：清空系统选择集。

**语法**：

```vbs
SSProcess.ClearSysSelection()
```

**说明**：

执行该函数，将清空当前系统选择集的对象。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSysSelection
End Sub
```


#### DeleteInSelectionXGeo 删除选择集中的转换地物

**函数**：`DeleteInSelectionXGeo`

**功能**：删除选择集中的转换地物。

**语法**：

```vbs
SSProcess.DeleteInSelectionXGeo ( index )
```

**参数**：

- index 地物索引号

**说明**：

执行该函数，可以对转换过程中地物进行删除

**示例**：

```vbs
需在输出脚本中运行
Sub OnClick()
SSProcess.ClearDataXParameter
SSProcess.SetDataXParameter "DataType", "1"
SSProcess.SetDataXParameter "ImportPathName", DWGfileName
SSProcess.SetDataXParameter "SaveAttrToMemoData", "1"
SSProcess.ImportData
End Sub
Function BeforeExportData()
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount- 1
SSProcess.DeleteInSelectionXGeo i
Next
End Function
```


#### DeleteInSelectionXNote 删除选择集中的转换注记

**函数**：`DeleteInSelectionXNote`

**功能**：删除选择集中的转换注记。

**语法**：

```vbs
SSProcess.DeleteInSelectionXNote (index)
```

**参数**：

- index 注记索引号

**说明**：

执行该函数，可以对转换过程中注记进行删除

**示例**：

```vbs
需在输出脚本中运行
Sub OnClick()
SSProcess.ClearDataXParameter
SSProcess.SetDataXParameter "DataType", "1"
SSProcess.SetDataXParameter "ImportPathName", DWGfileName
SSProcess.SetDataXParameter "SaveAttrToMemoData", "1"
SSProcess.ImportData
End Sub
Function BeforeExportData()
NoteCount = SSProcess.GetSelNoteCount
For i=0 ToNoteCount -1
SSProcess.DeleteInSelectionXNote i
Next
End Function
```


#### GetSelectPolygonObjAttr 获取多边形内选中对象的属性值

**函数**：`GetSelectPolygonObjAttr`

**功能**：获取多边形内选中对象的属性值。

**语法**：

```vbs
SSProcess.GetSelectPolygonObjAttr (index, attrField)
```

**参数**：

- index 对象序号
- attrField 属性值

**说明**：

执行该函数，获取多边形内选中对象的属性值。


#### GetSelectPolygonObjCount 获取多边形内选中对象的个数

**函数**：`GetSelectPolygonObjCount`

**功能**：获取多边形内选中对象的个数。

**语法**：

```vbs
SSProcess.GetSelectPolygonObjCount()
```


#### RemoveSelectPolygonObj 删除多边形内选中的指定对象

**函数**：`RemoveSelectPolygonObj`

**功能**：删除多边形内选中的指定对象。

**语法**：

```vbs
SSProcess.RemoveSelectPolygonObj (index)
```

**参数**：

- index 对象索引号

**说明**：

执行该函数，删除多边形内选中的指定对象 (只是移出选择对象列表,并不实际删除)


#### SelectFilter 选择过滤

**函数**：`SelectFilter`

**功能**：SelectFilter 函数

**语法**：

```vbs
SSProcess.SelectFilter()
```

**说明**：

执行该函数，将根据SetSelectCondition设定的选择条件,选择对象到选择集中。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "3103013"
SSProcess.SelectFilter
End Sub
```


#### SetSelectCondition 设置选择条件

**函数**：`SetSelectCondition`

**功能**：设置选择条件。

**语法**：

```vbs
SSProcess.SetSelectCondition(conditionName, operateMode, conditionValue)
```

**参数**：

- conditionName 条件名称
- 条件分SSOBJ_开头的基本属性，<>括住的几何特性, []括住的扩展属性,{}括住的拓扑关系特性,CallBackFunc_开头的回调函数方式
- SSOBJ_开头的基本属性见下表：
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_GroupID | 对象组 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 COLORNO( 颜色号 ) 、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_ObjectName | 对象汉字名称 , 为 FeatureCodeTB 表中的 ObjectName |
| SSObj_Thickness | 厚度 , 为 FeatureCodeTB 表中的 Thickness |
| SSObj_ExplodeStatus | 打散状态 , 为 FeatureCodeTB 表中的 Explode |
| SSObj_FilterStatus | 过滤状态 , 为 FeatureCodeTB 表中的 Filter |
| SSObj_ReverseStatus | 反向状态 , 为 FeatureCodeTB 表中的 Reverse |
| SSObj_Angle | 角度 , 以弧度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_FontWidth | 字宽，以图上 0.01 毫米为单位 |
| SSObj_FontHeight | 字高，以图上 0.01 毫米为单位 |
| SSObj_FontName | 字体名 |
| SSObj_FontClass | 注记分类号 |
| SSObj_FontAlignment | 对齐方式，包括 0 （中心）、 1 （左上角）、 2 （左下角）、 3 （左中）、 4 （右上角）、 5 （右中）、 6 （右下角）、 7 （上中）、 8 （下中） |
| SSObj_FontPosType | 注记排列方式 , 包括 0 （均分）、 1 （散列） |
| SSObj_FontDirection | 字头朝向，包括 0( 正东 ) 、 1( 正北 ) 、 2( 正西 ) 、 3( 正南 ) 、 4( 切线方向 ) 、 5( 法线方向 ) |
| SSObj_FontWordAngle | 字角度 , 以度为单位 |
| SSObj_FontStringAngle | 串角度 , 以度为单位 |
| SSObj_FontWeight | 字重 |
| SSObj_FontIlaticAngle | 倾斜，包括 0 （不倾斜）、 -1 （左斜）、 1 （右斜） |
| SSObj_FontDownAngle | 耸肩，包括 0 （不耸肩）、 -1 （左耸肩）、 1 （右耸肩） |
| SSObj_FontUnderLine | 下划线，包括 0 （无下划线）、 1 （有下划线） |
| SSObj_FontInterval | 字隔，以图上 0.01 毫米为单位 |
| SSObj_FontString | 注记内容 |
| SSObj_FontStringCount | 注记内容字符个数（汉字算两个字符） |
| SSObj_Area | 面积，以平方米为单位 |
| SSObj_Length | 2D 长度，以米为单位 |
| SSObj_3DLength | 3D 长度，以米为单位 |
| SSObj_PointCount | 空间点数 |
| SSObj_X | 空间点首点 X 坐标 |
| SSObj_Y | 空间点首点 Y 坐标 |
| SSObj_Z | 空间点首点 Z 坐标 |
| SSObj_PointName | 空间点首点点名 |
| SSObj_PointType | 空间点首点点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_X(index) | 指定索引空间点 X 坐标 |
| SSObj_Y(index) | 指定索引空间点 Y 坐标 |
| SSObj_Z(index) | 指定索引空间点 Z 坐标 |
| SSObj_PointName(index) | 指定索引空间点点名 |
| SSObj_PointType(index) | 指定索引空间点 点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_SymType | 符号类型 (P,L,LC,LP,Y,H,E) |
| SSObj_Reverse | 地物方向， 0 （正向）、 1 （反向） |
| SSObj_DrawAreaLabelStatus | 面心点标注状态 , 0( 不注 ) 1( 标注 ) |
| SSObj_AreaLabelX | 面心标注点 X 坐标 |
| SSObj_AreaLabelY | 面心标注点 X 坐标 |
| SSObj_Reverse | 地物方向， 0 （正向）、 1 （反向） |
| SSObj_Status | 地物状态 ,e_Valid = 0x0001, // 有效 e_Explode = 0x0002, // 是否已打散 e_Screen = 0x0004, // 是否被屏蔽 e_Duplicate = 0x0008, // 是否被复制品 e_IsQuote = 0x0010, // 是否索引 e_Hidden = 0x0020, // 是否隐藏 e_IsSelected = 0x0040, // 是否已选择 e_Editable = 0X0080, // 是否可编辑 e_Highlight = 0X0100, // 是否加亮 e_Gray = 0X0200, // 是否变灰 e_Edited = 0X0400, // 是否被修改 e_Refrence = 0X0800, // 参考图形 e_Opaque = 0X1000, // 是否压盖 |
| SSObj_GraphicInfo(flag) | 图形特征 ,flag 可以是以下值 : 25 面心标志点显示开关及相对多边形焦点偏移量 , 格式 : " 是否显示 (0 不显示 ,1 显示 ), 东偏 , 北偏 " ,( 米为单位 ) 26 个性化符号描述 , 格式 : " 是否替换现在符号描述 (0 不替换 ,1 替换 ); 符号描述语句 ( 多个语句用 ; 号分隔 ) " 101 图形缩放比例 , 格式 : " X 向比例 ,Y 向比例 ,Z 向比例 " ( 只支持整数 , 单位为 0.00001 米 ) |
- 几何特性见下表
| 参数 | 说明 |
|---|---|
| <Overlap> | 重叠对象，包括 0 （不重叠）、 1 （有重叠） |
| <Close> | 封闭对象，包括 0 （不封闭）、 1 （封闭） |
| <Clockwise> | 点列方向，包括 0 （逆时针）、 1 （顺时针） |
- 扩展属性为扩展属性表或MemoData中的任一字段，需用[]括住字段名称
- [ExAttr] 扩展属性
- 拓扑关系特性见下表
| 参数 | 说明 |
|---|---|
| {OuterObj} | 外围对象，条件值填写选择过滤条件分组名 |
| {NotOuterObj} | 无外围对象，条件值填写选择过滤条件分组名 |
| {OuterObjSearchRange} | 外围对象最大搜索范围，以米为单位 |
| {NotOuterObjSearchRange} | 无外围对象最大搜索范围，以米为单位 |
| {InnerObj} | 包含对象，条件值填写选择过滤条件分组名 |
| {NotInnerObj} | 无包含对象，条件值填写选择过滤条件分组名 |
| {InnerObjGetPointMode} | 被包含对象取点方式 ,   0( 判断所有点 ) 1( 只判断焦点 ) |
| {CrossObj} | 相交对象，条件值填写选择过滤条件分组名 |
| {NearObj} | 相近对象，条件值填写选择过滤条件分组名 |
| {NotNearObj} | 无相近对象，条件值填写选择过滤条件分组名 |
| {NearObjSearchRange} | 相近对象最大搜索范围，以米为单位 |
| {NotNearObjSearchRange} | 无相近对象最大搜索范围，以米为单位 |
| {OnLineObj} | 选择对象落在指定线上，条件值填写线选择过滤条件分组名 |
| {NotOnLineObj} | 选择对象不落在线上，条件值填写线选择过滤条件分组名 |
| {OnLineObjRange} | 判断点落在线上的范围，以米为单位 |
| {SelectPolygon_GroupName} | 选择多边形分组名，条件值填写选择过滤条件分组名 支持直接指定由 CreateMapFrame,CreateMapFrameByRegion, CreateMapFrameByCenterLine 等函数创建的 图幅缓存作为选择多边形。 格式 : MapFrameIndex_index 如 MapFrameIndex_1 |
| {SelectPolygon_InOutMode} | 多边形选择方式， 1 （选内） 2 （选外） 4 （选择相交），允许多值相加 |
| {SelectPolygon_CallBackFunc} | 过滤多边形内对象回调函数 , 条件值填写回调函数名称 , 如 CallBackFunc_FilterObj, 有岛时 , 只处理外环 说明 : 在回调函数内可应用以下函数操作当前多边形内已选中的对象 GetSelectPolygonObjCount() 获取多边形内选中对象的个数 GetSelectPolygonObjAttr(index, attrField ) 获取多边形内选中对象的属性值 RemoveSelectPolygonObj(index) 删除多边形内选中的指定对象 ( 只是移出选择对象列表 , 并不实际删除 ) |
| {RepeatedInPolygon_GroupName} | 当前选择条件选中的多边形内多次出现的同类对象，条件值填写同类对象过滤条件分组名 |
| {RepeatedInPolygon_FieldName} | 同类对象比较字段名列表 , 用 , 号分隔 |
| {RepeatedInPolygon_ReturnFirst} | 是否返回首个重复对象 (0 不返回 1 返回 ) " |
- 回调函数方式
- 回调函数必须在脚本中定义,函数参数按字符串方式传递,函数的返回值必须强制转成字符串.
| 参数 | 说明 |
|---|---|
| CallBackFunc_ 函数名 | 回调函数方式 , 如 "CallBackFunc_CheckString(SSObj_FontClass,   SSObj_FontString)" 如果同时有拓扑关联对象 , 则允许指定关联对 象的属性 , 以 SSSubObj_ 和 SSSubExtAttr_ 开头 , 如 " CallBackFunc_CheckString( SSObj_FontString, SSSubObj_FontString)" |
- operateMode 条件操作符
| 参数 | 说明 |
|---|---|
| = | 等于 |
| > | 大于 |
| < | 小于 |
| <> | 不等于 |
| LIKE | 包含指定字符 |
| NOT LIKE | 不包含指定字符 |
| CompareNoCase | 忽略大小写等于比较 |
| Dec | 小数位等于 |
- conditionValue 条件值
- 根据条件名称，指定相应的值类型，如果一个条件需要指定多个或关系的值列表，可用"值1,值2,..."的方式描述。

**说明**：

该函数，执行一次，设定一个条件，允许多次执行以设定多个选择条件，有多个条件时，条件的关系为与的关系。
函数执行只是设定选择条件，并不执行选择过滤操作，在设定完选择条件后，必须执行SelectFilter函数，才会真正的进行选择过滤操作

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Color", "==", "RGB(255,0,0)"
SSProcess.SelectFilter
End Sub
```


#### SetSelectConditionGroup 设置分组选择条件

**函数**：`SetSelectConditionGroup`

**功能**：设置分组选择条件。

**语法**：

```vbs
SSProcess.SetSelectConditionGroup( groupName, conditionName, operateMode, conditionValue )
```

**参数**：

- groupName 选择分组名称
- conditionName 条件名称
- 条件分SSOBJ_开头的基本属性，<>括住的几何特性, []括住的扩展属性,{}括住的拓扑关系特性,CallBackFunc_开头的回调函数方式
- SSOBJ_开头的基本属性见下表：
| 参数 | 说明 |
|---|---|
| SSObj_ID | 对象 ID |
| SSObj_GroupID | 对象组 ID |
| SSObj_Code | 编码 |
| SSObj_LayerName | 层名 |
| SSObj_Type | 对象类型，包括 POINT （点）、 LINE （线）、 AREA （面）和 NOTE （注记） 4 种 |
| SSObj_Color | 颜色，分 COLORNO( 颜色号 ) 、颜色值和 RGB( 红 , 绿，蓝 ) 三种指定方式 |
| SSObj_LineType | 线形，包括 0 （点）、 1 （线）、 2 （曲线）、 3 （园弧）和 4 （圆） |
| SSObj_LineWidth | 线宽，以图上 0.01 毫米为单位 |
| SSObj_Name | 对象名称 |
| SSObj_Byname | 别名 , 为 FeatureCodeTB 表中的 Byname |
| SSObj_ObjectName | 对象汉字名称 , 为 FeatureCodeTB 表中的 ObjectName |
| SSObj_Thickness | 厚度 , 为 FeatureCodeTB 表中的 Thickness |
| SSObj_ExplodeStatus | 打散状态 , 为 FeatureCodeTB 表中的 Explode |
| SSObj_FilterStatus | 过滤状态 , 为 FeatureCodeTB 表中的 Filter |
| SSObj_ReverseStatus | 反向状态 , 为 FeatureCodeTB 表中的 Reverse |
| SSObj_Angle | 角度 , 以弧度为单位 |
| SSObj_DataMark | 数据标识 |
| SSObj_MemoData | 备注内容 |
| SSObj_CreateTime | 创建时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_ModifyTime | 修改时间 年 - 月 - 日 时 : 分 : 秒 格式 |
| SSObj_FontWidth | 字宽，以图上 0.01 毫米为单位 |
| SSObj_FontHeight | 字高，以图上 0.01 毫米为单位 |
| SSObj_FontName | 字体名 |
| SSObj_FontClass | 注记分类号 |
| SSObj_FontAlignment | 对齐方式，包括 0 （中心）、 1 （左上角）、 2 （左下角）、 3 （左中）、 4 （右上角）、 5 （右中）、 6 （右下角）、 7 （上中）、 8 （下中） |
| SSObj_FontPosType | 注记排列方式 , 包括 0 （均分）、 1 （散列） |
| SSObj_FontDirection | 字头朝向，包括 0( 正东 ) 、 1( 正北 ) 、 2( 正西 ) 、 3( 正南 ) 、 4( 切线方向 ) 、 5( 法线方向 ) |
| SSObj_FontWordAngle | 字角度 , 以度为单位 |
| SSObj_FontStringAngle | 串角度 , 以度为单位 |
| SSObj_FontWeight | 字重 |
| SSObj_FontIlaticAngle | 倾斜，包括 0 （不倾斜）、 -1 （左斜）、 1 （右斜） |
| SSObj_FontDownAngle | 耸肩，包括 0 （不耸肩）、 -1 （左耸肩）、 1 （右耸肩） |
| SSObj_FontUnderLine | 下划线，包括 0 （无下划线）、 1 （有下划线） |
| SSObj_FontInterval | 字隔，以图上 0.01 毫米为单位 |
| SSObj_FontString | 注记内容 |
| SSObj_FontStringCount | 注记内容字符个数（汉字算两个字符） |
| SSObj_Area | 面积，以平方米为单位 |
| SSObj_Length | 2D 长度，以米为单位 |
| SSObj_3DLength | 3D 长度，以米为单位 |
| SSObj_PointCount | 空间点数 |
| SSObj_X | 空间点首点 X 坐标 |
| SSObj_Y | 空间点首点 Y 坐标 |
| SSObj_Z | 空间点首点 Z 坐标 |
| SSObj_PointName | 空间点首点点名 |
| SSObj_PointType | 空间点首点点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_X(index) | 指定索引空间点 X 坐标 |
| SSObj_Y(index) | 指定索引空间点 Y 坐标 |
| SSObj_Z(index) | 指定索引空间点 Z 坐标 |
| SSObj_PointName(index) | 指定索引空间点点名 |
| SSObj_PointType(index) | 指定索引空间点 点类型，包括 0x0001 （实测的或计算的）、 0x0002 （高程注记否）、 0x0004 （参加建模否）、 0x0008 （新产生的点或旧点）、 0x0010 （转向点或一般点）、 0x0020 （断点）、 0x0040 （直线）、 0x0080 （曲线）、 0x00c0 （圆弧）、 0x1000 （依附点）、 0x2000 （群集点）、 0x4000 （平滑点）、 0x8000 （特征点） |
| SSObj_SymType | 符号类型 (P,L,LC,LP,Y,H,E) |
| SSObj_Reverse | 地物方向， 0 （正向）、 1 （反向） |
| SSObj_DrawAreaLabelStatus | 面心点标注状态 , 0( 不注 ) 1( 标注 ) |
| SSObj_AreaLabelX | 面心标注点 X 坐标 |
| SSObj_AreaLabelY | 面心标注点 X 坐标 |
| SSObj_Reverse | 地物方向， 0 （正向）、 1 （反向） |
| SSObj_Status | 地物状态 ,e_Valid = 0x0001, // 有效 e_Explode = 0x0002, // 是否已打散 e_Screen = 0x0004, // 是否被屏蔽 e_Duplicate = 0x0008, // 是否被复制品 e_IsQuote = 0x0010, // 是否索引 e_Hidden = 0x0020, // 是否隐藏 e_IsSelected = 0x0040, // 是否已选择 e_Editable = 0X0080, // 是否可编辑 e_Highlight = 0X0100, // 是否加亮 e_Gray = 0X0200, // 是否变灰 e_Edited = 0X0400, // 是否被修改 e_Refrence = 0X0800, // 参考图形 e_Opaque = 0X1000, // 是否压盖 |
| SSObj_GraphicInfo(flag) | 图形特征 ,flag 可以是以下值 : 25 面心标志点显示开关及相对多边形焦点偏移量 , 格式 : " 是否显示 (0 不显示 ,1 显示 ), 东偏 , 北偏 " ,( 米为单位 ) 26 个性化符号描述 , 格式 : " 是否替换现在符号描述 (0 不替换 ,1 替换 ); 符号描述语句 ( 多个语句用 ; 号分隔 ) " 101 图形缩放比例 , 格式 : " X 向比例 ,Y 向比例 ,Z 向比例 " ( 只支持整数 , 单位为 0.00001 米 ) |
- 几何特性见下表：
| 参数 | 说明 |
|---|---|
| <Overlap> | 重叠对象，包括 0 （不重叠）、 1 （有重叠） |
| <Close> | 封闭对象，包括 0 （不封闭）、 1 （封闭） |
| <Clockwise> | 点列方向，包括 0 （逆时针）、 1 （顺时针） |
- 扩展属性为扩展属性表或MemoData中的任一字段，需用[]括住字段名称
- [ExAttr] 扩展属性
- 拓扑关系特性见下表：
| 参数 | 说明 |
|---|---|
| {OuterObj} | 外围对象，条件值填写选择过滤条件分组名 |
| {NotOuterObj} | 无外围对象，条件值填写选择过滤条件分组名 |
| {OuterObjSearchRange} | 外围对象最大搜索范围，以米为单位 |
| {NotOuterObjSearchRange} | 无外围对象最大搜索范围，以米为单位 |
| {InnerObj} | 包含对象，条件值填写选择过滤条件分组名 |
| {NotInnerObj} | 无包含对象，条件值填写选择过滤条件分组名 |
| {InnerObjGetPointMode} | 被包含对象取点方式 ,   0( 判断所有点 ) 1( 只判断焦点 ) |
| {CrossObj} | 相交对象，条件值填写选择过滤条件分组名 |
| {NearObj} | 相近对象，条件值填写选择过滤条件分组名 |
| {NotNearObj} | 无相近对象，条件值填写选择过滤条件分组名 |
| {NearObjSearchRange} | 相近对象最大搜索范围，以米为单位 |
| {NotNearObjSearchRange} | 无相近对象最大搜索范围，以米为单位 |
- 回调函数方式：
- 回调函数必须在脚本中定义,函数参数按字符串方式传递，函数的返回值必须强制转成字符串。
| 参数 | 说明 |
|---|---|
| CallBackFunc_ 函数名 | 回调函数方式 , 如 "CallBackFunc_CheckString( SSObj_FontClass,   SSObj_FontString)" 如果同时有拓扑关联对象 , 则允许指定关联对象的属性 , 以 SSSubObj_ 和 SSSubExtAttr_ 开头 , 如 "CallBackFunc_CheckString(SSObj_FontString,   SSSubObj_FontString)" |
- operateMode 条件操作符
| 参数 | 说明 |
|---|---|
| = | 等于 |
| > | 大于 |
| < | 小于 |
| <> | 不等于 |
| LIKE | 包含指定字符 |
| NOT LIKE | 不包含指定字符 |
| CompareNoCase | 忽略大小写等于比较 |
| Dec | 小数位等于 |
- conditionValue 条件值
- 根据条件名称，指定相应的值类型，如果一个条件需要指定多个或关系的值列表，可用"值1,值2,..."的方式描述。

**说明**：

该函数，执行一次，设定分组的一个条件，允许多次执行以设定多个分组的多个选择条件，有多个条件时，条件的关系为与的关系。函数设定的分组选择条件，只供SetSelectCondition函数中设置拓扑关系特性时使用，在执行SelectFilter函数时，不会直接使用分组条件。
在分组选择条件使用完成后，必须调用ClearSelectConditionGroups函数释放空间。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelectConditionGroups
SSProcess.SetSelectConditionGroup "房屋范围线", "SSObj_Code", "==", "4410"
SSProcess.SetSelectConditionGroup "房屋范围线", "SSObj_Color", "==", "RGB(255,0,0)"
End Sub
```


#### SortSelectionObj 选择集内对象排序（从西到东，从北到南）

**函数**：`SortSelectionObj`

**功能**：选择集内对象排序（从西到东，从北到南）。

**参数**：

- rowHeighte 行高

**说明**：

该函数按照设定的行高，进行排序，同行自西向东排序，排完一行再由北向南进行下一行排序。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code","=" ,"6803311"
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount()
'排序
SSProcess.SortSelectionObj 20 'rowHeighte
For i=0 To geoCount- 1
JZDH = Cstr (i+1)
SSProcess.SetSelGeoValue i ,"SSObj_Name", JZDH
SSProcess.AddSelGeoToSaveGeoList i
Next
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### SortSelectionObj1 选择集内对象排序（从东到西，从北到南）

**函数**：`SortSelectionObj1`

**功能**：选择集内对象排序（从东到西，从北到南）。

**语法**：

```vbs
SSProcess.SortSelectionObj1 (rowHeighte)
```

**参数**：

- rowHeighte 行高

**说明**：

该函数按照设定的行高，进行排序，同行自东向西排序，排完一行再由北向南进行下一行排序。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code","=" ,"6803311"
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount()
'排序
SSProcess.SortSelectionObj1 20 'rowHeighte
For i=0 To geoCount- 1
JZDH = Cstr (i+1)
SSProcess.SetSelGeoValue i ,"SSObj_Name" , JZDH
SSProcess.AddSelGeoToSaveGeoList i
Next
SSProcess.SaveBufferObjToDatabase
End Sub
```


#### TransSelectionObjToAttr 选择集图形转属性

**函数**：`TransSelectionObjToAttr`

**功能**：选择集图形转属性。

**语法**：

```vbs
SSProcess.TransSelectionObjToAttr (geoID, fieldName, delObj)
```

**参数**：

- geoID 存储图形的地物ID
- fieldname 存储图形的字段名称
- delObj 是否删除被存储的图形

**说明**：

执行该函数，将通过选择集选中的图形文件以二进制格式存储到指定地物的属性字段中

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "=" , "9203600"
SSProcess.SelectFilter
geoCount = SSProcess.GetSelGeoCount
If geoCount=0 Then
Msgbox "数据中‘竣工数据范围线-9203600’不存在，已取消操作！" :Exit Sub
Elseif geoCount>1 Then
Msgbox "数据中‘竣工数据范围线-9203600’有多个，已取消操作！" :Exit Sub
End If
geoID=SSProcess.GetSelGeoValue( 0, "SSObj_ID" )
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=" , "POINT,LINE,AREA,NOTE"
SSProcess.SetSelectCondition "SSObj_Code", "<>" , "9203600"
SSProcess.SelectFilter
SSProcess.TransSelectionObjToAttr geoID, "图廓内数据" , 1
End Sub
```


#### UpdateSysSelection 更新当前选择到系统选择集

**函数**：`UpdateSysSelection`

**功能**：脚本选择的对象与系统选择集的对象互转。

**语法**：

```vbs
SSProcess.UpdateSysSelection( flag )
```

**参数**：

- flag 选择集更新方式 0(系统选择集选择的内容更新到脚本选择集中) 1(脚本选择集选择的内容更新到系统选择集中)

**说明**：

执行该函数，将实现脚本选择的对象与系统选择集的对象互转。

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSysSelection
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Color", "==", "RGB(255,0,0)"
SSProcess.SelectFilter
SSProcess.UpdateSysSelection 1
End Sub
```


---

### 坐标转换函数

#### LongiLatiToxy54 经纬度转 54 系 XY 坐标

**函数**：`LongiLatiToxy54`

**功能**：经纬度转54系XY坐标

**语法**：

```vbs
SSProcess.LongiLatiToxy54(lon0, b, l, px, py)
```

**参数**：

- lon0 中央子午线
- b 原始数据纬度
- l 原始数据经度
- px 转换后的x坐标
- py 转换后的y坐标

**说明**：

执行该函数，将可以将数据经纬度坐标转换为XY坐标。

**示例**：

```vbs
Sub OnClick()
'选择所有数据
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
geoCount =SSProcess.GetSelGeoCount
For i=0 To geoCount -1
'获取地物的点数
pointCount = SSProcess.GetSelGeoPointCount (i )
SSProcess.LockSelGeoPoint i, 1
'对于每个点进行数据坐标转换
For j=0 To pointCount -1
SSProcess.GetSelGeoPoint i, j, x , y, z , ptype, pname
SSProcess.LongiLatiToxy54 126, y, x, x0, y0
SSProcess.SetSelGeoPoint i, j, x0 , y0, z , ptype, pname
Next
SSProcess.UpdateSelGeoPoint i
SSProcess.LockSelGeoPoint i, 0
Next
End Sub
```


#### LongiLatiToxy80 经纬度转 80 系 XY 坐标

**函数**：`LongiLatiToxy80`

**功能**：经纬度转80系XY坐标

**语法**：

```vbs
SSProcess.LongiLatiToxy80(lon0, b, l, px, py)
```

**参数**：

- lon0 中央子午线
- b 原始纬度
- l 原始经度
- px 转换后的x坐标
- py 转换后的y坐标

**说明**：

执行该函数，将可以将数据经纬度坐标转换为XY坐标

**示例**：

```vbs
Sub OnClick()
'选择所有数据
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
geoCount =SSProcess.GetSelGeoCount
For i=0 To geoCount -1
'获取地物的点数
pointCount = SSProcess.GetSelGeoPointCount (i )
SSProcess.LockSelGeoPoint i, 1
'对于每个点进行数据坐标转换
For j=0 To pointCount -1
SSProcess.GetSelGeoPoint i, j, x , y, z , ptype, pname
SSProcess.LongiLatiToxy80 126, y, x, x0, y0
SSProcess.SetSelGeoPoint i, j, x0 , y0, z , ptype, pname
Next
SSProcess.UpdateSelGeoPoint i
SSProcess.LockSelGeoPoint i, 0
Next
End Sub
```


#### LongiLatiToxyCGCS2000 经纬度转 2000 系 XY 坐标

**函数**：`LongiLatiToxyCGCS2000`

**功能**：打开文件选择对话框，选择文件名。

**参数**：

- lon0 中央子午线经度，度为单位
- b 原始纬度
- l 原始经度
- px 转换后的坐标x
- py 转换后的坐标y

**说明**：

执行该函数，将把CGCS2000椭球的大地坐标换算成该椭球对应投影带下的高斯坐标。
函数执行成功，返回高斯平面坐标。

**示例**：

```vbs
Sub OnClick()
'获取中央子午线经度
CentralMeridian = GetStringValueFromIniInfoTB("Projection","CentralMeridian")
' GetStringValueFromIniInfoTB 为自定义函数，用于从工程EDB的IniInfoTB表中获取中央子午线经度
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount-1
'--------------------坐标转换-----------------
pointCount = SSProcess.GetSelGeoPointCount (i)
SSProcess.LockSelGeoPoint i, 1
For j=0 To pointCount-1
SSProcess.GetSelGeoPoint i, j, x, y, z, ptype, pname
'经纬度（度）换算成CGCS2000高斯坐标
SSProcess.LongiLatiToxyCGCS2000 CentralMeridian, y, x, x0, y0
SSProcess.SetSelGeoPoint i, j, x0, y0, z, ptype, pname
Next
SSProcess.UpdateSelGeoPoint i
SSProcess.LockSelGeoPoint i, 0
Next
End Sub
```


#### LongiLatiToxyWGS84 经纬度转 80 系 XY 坐标

**函数**：`LongiLatiToxyWGS84`

**功能**：经纬度转80系XY坐标

**语法**：

```vbs
SSProcess.LongiLatiToxyWGS84 (lon0, b, l, px, py)
```

**参数**：

- lon0 中央子午线
- b 纬度值
- l 经度值
- px WGS84坐标系x值
- py WGS84坐标系y值

**说明**：

该函数为坐标转换函数，执行该函数可以把经纬度转换为WGS84坐标系中坐标。

**示例**：

```vbs
Sub OnInitScript()
mode = 1 ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="经纬度转WGS84系XY坐标"
SSProcess.AddInputParameter "中央子午线" ,  "117",  0,"" , "设置当地所在中央子午线"
SSProcess.AddInputParameter "纬度值B" , "48", 0 , "", "输入要转换的纬度值"
SSProcess.AddInputParameter "经度值L" , "240", 0 , "", "输入要转换的经度值"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSProcess.UpdateScriptDlgParameter 1
CentralMeridian = SSProcess.GetInputParameter ("中央子午线")
b = SSProcess.GetInputParameter("纬度值B")
l = SSProcess.GetInputParameter("经度值L")
SSProcess.LongiLatiToxyWGS84 lon0, b, l , pX, pY
Msgbox "pX:"& pX&chr(13)& "pY:"&pY
End Sub
Sub OnCancel()
End Sub
```


#### LongiLatiToxyz54 经纬度转 54 系 XYZ 坐标

**函数**：`LongiLatiToxyz54`

**功能**：经纬度转54系XYZ坐标

**语法**：

```vbs
SSProcess.LongiLatiToxyz54(lon0, b, l, h, px, py, pz)
```

**参数**：

- lon0 中央子午线
- b 纬度值
- l 经度值
- h 高度值
- px 54坐标系中X值
- py 54坐标系中Y值
- pz 54坐标系中Z值

**说明**：

该函数为坐标转换函数，执行该函数可以把经纬度转换为54坐标系中坐标。

**示例**：

```vbs
Sub OnInitScript()
mode = 1 ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="经纬度转54系XYZ坐标"
SSProcess.AddInputParameter "中央子午线" ,  "117",  0,"" , "设置当地所在中央子午线"
SSProcess.AddInputParameter "纬度值B" , "48", 0 , "", "输入要转换的纬度值"
SSProcess.AddInputParameter "经度值L" , "240", 0 , "", "输入要转换的经度值"
SSProcess.AddInputParameter "高度值H" , "56", 0 , "", "输入要转换的高度值"
SSProcess.ShowScriptDlg mode ,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSProcess.UpdateScriptDlgParameter 1
CentralMeridian = SSProcess.GetInputParameter ( "中央子午线" )
b = SSProcess.GetInputParameter ( "纬度值B" )
l = SSProcess.GetInputParameter ( "经度值L" )
h= SSProcess.GetInputParameter ( "高度值H" )
SSProcess.LongiLatiToxyz54 lon0 , b, l , h, pX , pY, pZ
Msgbox "pX:"& pX&chr(13)& "pY:"&pY&chr( 13)&"pZ:" &pZ
End Sub
Sub OnCancel()
End Sub
```


#### LongiLatiToxyz80 经纬度转 80 系 XYZ 坐标

**函数**：`LongiLatiToxyz80`

**功能**：经纬度转80系XYZ坐标

**语法**：

```vbs
SSProcess.LongiLatiToxyz80 (lon0, b, l, h, px, py, pz )
```

**参数**：

- lon0 中央子午线
- b 纬度值
- l 经度值
- h 高度值
- px 2000坐标系中X值
- py 2000坐标系中Y值
- pz 2000坐标系中Z值

**说明**：

该函数为坐标转换函数，执行该函数可以把经纬度转80系XYZ坐标。

**示例**：

```vbs
Sub OnInitScript()
mode = 1 ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="经纬度转80系XYZ坐标"
SSProcess.AddInputParameter "中央子午线" , "117", 0,"" , "设置当地所在中央子午线"
SSProcess.AddInputParameter "纬度值B" , "48", 0 , "", "输入要转换的纬度值"
SSProcess.AddInputParameter "经度值L" , "240", 0 , "", "输入要转换的经度值"
SSProcess.AddInputParameter "高度值H" , "56", 0 , "", "输入要转换的高度值"
SSProcess.ShowScriptDlg mode ,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSProcess.UpdateScriptDlgParameter 1
CentralMeridian = SSProcess.GetInputParameter ( "中央子午线" )
b = SSProcess.GetInputParameter ( "纬度值B" )
l = SSProcess.GetInputParameter ( "经度值L" )
h= SSProcess.GetInputParameter ( "高度值H" )
SSProcess.LongiLatiToxyz80 lon0 , b, l , h, pX , pY, pZ
Msgbox "pX:"& pX&chr(13)& "pY:"&pY&chr( 13)&"pZ:" &pZ
End Sub
Sub OnCancel()
End Sub
```


#### LongiLatiToxyzCGCS2000 经纬度转 2000 系 XYZ 坐标

**函数**：`LongiLatiToxyzCGCS2000`

**功能**：经纬度转2000系XYZ坐标

**语法**：

```vbs
SSProcess.LongiLatiToxyzCGCS2000 (lon0, b, l, h, px, py, pz)
```

**参数**：

- lon0 中央子午线
- b 纬度值
- l 经度值
- h 高度值
- px 2000坐标系中X值
- py 2000坐标系中Y值
- pz 2000坐标系中Z值

**说明**：

该函数为坐标转换函数，执行该函数可以把经纬度转2000系XYZ坐标。

**示例**：

```vbs
Sub OnInitScript()
mode = 1 ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="经纬度转2000系XYZ坐标"
SSProcess.AddInputParameter "中央子午线" , "117", 0,"" , "设置当地所在中央子午线"
SSProcess.AddInputParameter "纬度值B" , "48", 0 , "", "输入要转换的纬度值"
SSProcess.AddInputParameter "经度值L" , "240", 0 , "", "输入要转换的经度值"
SSProcess.AddInputParameter "高度值H" , "56", 0 , "", "输入要转换的高度值"
SSProcess.ShowScriptDlg mode ,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSProcess.UpdateScriptDlgParameter 1
CentralMeridian = SSProcess.GetInputParameter ( "中央子午线" )
b = SSProcess.GetInputParameter ( "纬度值B" )
l = SSProcess.GetInputParameter ( "经度值L" )
h= SSProcess.GetInputParameter ( "高度值H" )
SSProcess.LongiLatiToxyzCGCS2000 lon0 , b, l , h, pX , pY, pZ
Msgbox "pX:"& pX&chr(13)& "pY:"&pY&chr( 13)&"pZ:" &pZ
End Sub
```


#### LongiLatiToxyzWGS84 经纬度转 80 系 XYZ 坐标

**函数**：`LongiLatiToxyzWGS84`

**功能**：经纬度转80系XYZ坐标

**语法**：

```vbs
SSProcess.LongiLatiToxyzWGS84 (lon0, b, l, h, px, py, pz )
```

**参数**：

- lon0 中央经度
- b 纬度
- l 经度
- h 大地高
- px 转换后的X坐标
- py 转换后的Y坐标
- pz 转换后的Z坐标

**说明**：

使用此函数将根据指定的经纬度blh值返回80坐标系xyz值。

**示例**：

```vbs
108带的经纬度值转换为80坐标系的xyz值。
Sub OnClick()
lon0="108" :b=108:l= 35:h= 25
SSProcess.LongiLatiToxyzWGS84 lon0 ,b,l ,h,pX ,pY,pZ
End Sub
```


#### TransBLtoXY 经纬度转换 XY 坐标

**函数**：`TransBLtoXY`

**功能**：经纬度转换XY坐标

**语法**：

```vbs
SSProcess.TransBLtoXY( ke, bb, ll, pX, pY, pR )
```

**参数**：

- ke 带宽
- bb 需要转换的经度
- ll 需要转换的纬度
- pX 转换后的X坐标
- pY 转换后的Y坐标
- pR 返回的椭球半径

**说明**：

使用此函数可以将数据的经纬度坐标转换为XY坐标。

**示例**：

```vbs
Sub OnClick()
name = "TransBLtoXY "
Dim pB, pL
pB=40.01:pL=114.02
SSProcess.TransBLtoXY 3,pB, pL, xx, yy,pR
Msgbox name&":"& chr(13)&"pB = " &pB & "pL = "&pL & chr(13)&" xx = "&xx& "yy = "&yy
End Sub
```


#### TransCoord_4p 四参数坐标转换

**函数**：`TransCoord_4p`

**功能**：使用4参数，进行坐标系间的坐标换算。

**语法**：

```vbs
SSProcess.TransCoord_4p(lon0, xoffset, yoffset, srcCoordType, tagCoordType, srcEllipsoid, tagEllipsoid, dx, dy, da, dm, x, y, pX, pY)
```

**参数**：

- lon0 中央子午线(度.分秒)
- xoffset X（纵）坐标偏移量
- yoffset Y（横）坐标偏移量
- srcCoordType 源坐标类型：0 大地坐标(BLH) BL度.分秒 H米；1 大地坐标(BLH) BL度.分 H米；2 大地坐标(BLH) BL度 H米；3 平面坐标(xyh) 米
- srcCoordType 目标坐标类型：0 大地坐标(BLH) BL度.分秒 H米 1 大地坐标(BLH) BL度.分 H米 2 大地坐标(BLH) BL度 H米 3 平面坐标(xyh) 米
- srcEllipsoid 源坐标椭球基准：0 北京-54坐标系；1 西安-80坐标系；2 WGS-84坐标系；3 WGS-72坐标系；4 CGCS-2000坐标系
- tagEllipsoid 目标坐标椭球基准：0 北京-54坐标系；1 西安-80坐标系；2 WGS-84坐标系；3 WGS-72坐标系；4 CGCS-2000坐标系
- dx X坐标偏移(米)
- dy Y坐标偏移(米)
- da 轴旋转角(秒)
- dm 尺度
- x 源坐标
- y 源坐标
- pX 目标坐标
- pY 目标坐标

**说明**：

使用此函数可以直接对当前工程数据进行修改，或将在BeforeExportData/BeforeSaveImportData函数中调用，在数据导出/导入环节进行坐标换算。

**示例**：

```vbs
将当前工程数据（2000坐标系）输出为一个80坐标系的EDB。
Sub OnClick()
'清空转换参数
SSProcess.ClearDataXParameter
'设置导入文件格式为EPS EDB
SSProcess.SetDataXParameter "DataType", "21"
' 符号打散方式。 0（自动打散）、 1（根据编码表设定打散）、 2（全部不打散）
SSProcess.SetDataXParameter "SymbolExplodeMode", "2"
'开始导入数据
SSProcess.ExportData
End Sub
Function BeforeExportData()
lon0 = 105:xoffset = 0:yoffset = 500000:srcCoordType = 3:tagCoordType = 3:srcEllipsoid = 4:tagEllipsoid = 1:dx= -51.689:dy = 1.650:da = 0.00000211:dm =1.0000722
'点线面处理
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount-1
pointCount = SSProcess.GetSelGeoPointCount (i)
SSProcess.LockSelGeoPoint i, 1
For j=0 To pointCount-1
SSProcess.GetSelGeoPoint i, j, x, y, z, ptype, pname
'四参数转换，参数见上
SSProcess.TransCoord_4p lon0, xoffset, yoffset, srcCoordType, tagCoordType, srcEllipsoid, tagEllipsoid, dx, dy, da, dm, x, y, pX, pY
SSProcess.SetSelGeoPoint i, j, Px, Py, z, ptype, pname
Next
SSProcess.UpdateSelGeoPoint i
SSProcess.LockSelGeoPoint i, 0
Next
'注记处理
noteCount = SSProcess.GetSelNoteCount
For i=0 To noteCount-1
pointCount = SSProcess.GetSelNotePointCount(i)
For j = 0 To pointCount - 1
SSProcess.GetSelNotePoint i, j, x, y, z, ptype, name
'四参数转换，参数见上
SSProcess.TransCoord_4p lon0, xoffset, yoffset, srcCoordType, tagCoordType, srcEllipsoid, tagEllipsoid, dx, dy, da, dm, x, y, pX, pY
SSProcess.SetSelNotePoint i, j, Px, Py, z, ptype, name
Next
SSProcess.UpdateSelNotePoint i
Next
End Function
```


#### TransCoord_7p 七参数坐标转换

**函数**：`TransCoord_7p`

**功能**：使用7参数，进行坐标系间的坐标换算。

**语法**：

```vbs
SSProcess.TransCoord_7p(lon0, xoffset, yoffset, srcCoordType, tagCoordType, srcEllipsoid, tagEllipsoid, dx, dy, dz, dxa, dya, dza, dm, x, y, z, pX, pY, pZ)
```

**参数**：

- lon0 中央子午线(度.分秒)
- xoffset 纵坐标偏移量
- yoffset 横坐标偏移量
- srcCoordType 源坐标类型：0 大地坐标(BLH) BL度.分秒 H米；1 大地坐标(BLH) BL度.分 H米；2 大地坐标(BLH) BL度 H米；3 平面坐标(xyh) 米
- srcCoordType 目标坐标类型：0 大地坐标(BLH) BL度.分秒 H米；1 大地坐标(BLH) BL度.分 H米；2 大地坐标(BLH) BL度 H米；3 平面坐标(xyh) 米
- srcEllipsoid 源坐标椭球基准：0 北京-54坐标系；1 西安-80坐标系；2 WGS-84坐标系；3 WGS-72坐标系；4 CGCS-2000坐标系
- tagEllipsoid 目标坐标椭球基准：0 北京-54坐标系；1 西安-80坐标系；2 WGS-84坐标系； 3 WGS-72坐标系；4 CGCS-2000坐标系
- dx X坐标偏移(米)
- dy Y坐标偏移(米)
- dz Z坐标偏移(米)
- dxa X轴旋转角(秒)
- dya Y轴旋转角(秒)
- dza Z轴旋转角(秒)
- dm 尺度
- x 源X坐标(转换前)
- y 源Y坐标
- z 源Z坐标
- pX 目标坐标(转换后)
- pY 目标坐标
- pZ 目标坐标

**说明**：

使用此函数可以直接对当前工程数据进行修改，或将在BeforeExportData/BeforeSaveImportData函数中调用，在数据导出/导入环节进行坐标换算。


#### TransCoordXYZ 坐标投影换算

**函数**：`TransCoordXYZ`

**功能**：坐标体系转换计算主函数。

**语法**：

```vbs
SSProcess.TransCoordXYZ(planeSysName, heightSysName, reverse, transxy, transz, surveyDH, ByRef x, ByRef y, ByRef z)
```

**参数**：

- planeSysName 平面坐标系名称
- heightSysName 高程基准名称
- reverse 是否反向换算，0（正向换算），1（反向换算）
- transxy 是否进行平面坐标换算，0（不进行平面坐标换算），1（进行平面坐标换算）
- transz 是否进行高程换算，0（不进行高程换算），1（进行高程换算）
- x,y,z 转换源坐标，参数为引用方式传递，

**说明**：

该函数是EPS的系统消息函数，在执行SSProcess.TransSelectionObjCoord函数后，将会自动激活该函数进行坐标换算处理。
在进行坐标换算处理时，必须在脚本中同时定义该函数，并在函数中实现坐标换算模型。

**示例**：

```vbs
把选择集对象进行四参数转换。
Function TransCoordXYZ(ByVal planeSysName, ByVal heightSysName, ByVal reverse, ByVal transxy, ByVal transz, ByVal surveyDH, ByRef x, ByRef y, ByRef z)
Dim dx,dy,m,a
dx = 123.5
dy = 57.33
m = 1.00000000098
a = 0.00000007
x = dx + m*x*cos(a) + m*y*sin(a)
y = dy + m*y*cos(a) - m*x*sin(a)
End Function
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
SSProcess.TransSelectionObjCoord ""，""，0, 1, 0
SSProcess.ClearSelection
End Sub
```


#### TransSelectionObjCoord 对选择集对象作坐标体系换算

**函数**：`TransSelectionObjCoord`

**功能**：对选择集对象作坐标体系换算。

**语法**：

```vbs
SSProcess.TransSelectionObjCoord(planeSysName, heightSysName, reverse, transXY, transZ)
```

**参数**：

- planeSysName 平面坐标系名称
- heightSysName 高程基准名称
- reverse 是否反向换算，0（正向换算），1（反向换算）
- transXY 是否进行平面坐标换算，0（不进行平面坐标换算），1（进行平面坐标换算）
- transz 是否进行高程换算，0（不进行高程换算），1（进行高程换算）

**说明**：

执行该函数，将会激活系统消息函数TransCoordXYZ，具体的转换模型在TransCoordXYZ函数中指定，因此必须同时定义TransCoordXYZ。

**示例**：

```vbs
把选择集对象进行四参数转换。
Function TransCoordXYZ(ByVal planeSysName, ByVal heightSysName, ByVal reverse, ByVal transxy, ByVal transz, ByVal surveyDH, ByRef x, ByRef y, ByRef z)
Dim dx,dy,m,a
dx = 123.5
dy = 57.33
m = 1.00000000098
a = 0.00000007
x = dx + m*x*cos(a) + m*y*sin(a)
y = dy + m*y*cos(a) - m*x*sin(a)
End Function
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SelectFilter
SSProcess.TransSelectionObjCoord "","",0, 1, 0
SSProcess.ClearSelection
End Sub
```


#### TransXYtoBL XY 坐标转换经纬度

**函数**：`TransXYtoBL`

**功能**：XY坐标转换经纬度

**语法**：

```vbs
SSProcess.TransXYtoBL(ke, x, y, pB, pL, pR)
```

**参数**：

- ke 带宽
- x 需要转换的X坐标
- y 需要转换的Y坐标
- pB 转换完后的纬度
- pL 转换完后的经度
- pR 返回椭球半径

**说明**：

执行该函数，将坐标转换为经纬度。

**示例**：

```vbs
Sub OnClick()
x = 760.204:y=352.613
SSProcess.TransXYtoBL  3, x, y, pB, pL, pR
End Sub
```


#### xy2000ToLongiLati 2000 系 XY 坐标转换经纬度

**函数**：`xyz2000ToLongiLati`

**功能**：2000系XYZ坐标转换经纬度

**语法**：

```vbs
SSProcess.xyz2000ToLongiLati (lon0, x, y, h, pB, pL, pH)
```

**参数**：

- lon0 中央子午线经度，度为单位。
- X 高斯坐标X
- y 高斯坐标Y
- h 高斯坐标H
- pB 大地坐标纬度。
- pL 大地坐标经度。
- pH 大地坐标高程。

**说明**：

执行该函数，将把高斯坐标换算成该椭球对应投影带下的 CGCS2000椭球的大地坐标。
函数执行成功，返回CGCS2000椭球的大地坐标 。


#### xy54ToLongiLati 系 XY 坐标转换经纬度

**函数**：`xy54ToLongiLati`

**功能**：54系XY坐标转换经纬度

**语法**：

```vbs
SSProcess.xy54ToLongiLati(lon0, x, y, pB, pL)
```

**参数**：

- lon0 中央子午线
- x 54系x坐标值
- y 54系x坐标值
- pB 纬度值
- pL 经度值

**说明**：

该函数将54系XY坐标转换为经纬度。

**示例**：

```vbs
Sub OnInitScript()
mode =1’=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="54系XY坐标转经纬度"
SSProcess.AddInputParameter "中央子午线","117",0,"","设置当地所在中央子午线"
SSProcess.AddInputParameter "X坐标","3122389.551",0,"","输入要转换的X坐标值"
SSProcess.AddInputParameter "Y坐标","501523.305",0,"","输入要转换的Y坐标值"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSProcess.UpdateScriptDlgParameter 1
CentralMeridian = SSProcess.GetInputParameter ("中央子午线")
x = SSProcess.GetInputParameter ("X坐标")
y = SSProcess.GetInputParameter ("Y坐标")
SSProcess.xy54ToLongiLati CentralMeridian, x, y, pB, pL
Msgbox "PB:"&pB&chr(13)&"PL:"&pL
End Sub
Sub OnCancel()
End Sub
```


#### xy80ToLongiLati 系 XY 坐标转换经纬度

**函数**：`xy80ToLongiLati`

**功能**：80系XY坐标转换经纬度

**语法**：

```vbs
SSProcess.xy80ToLongiLati(lon0, x, y, pB, pL)
```

**参数**：

- lon0 中央子午线
- x 80系x坐标值
- y 80系x坐标值
- pB 纬度值
- pL 经度值

**说明**：

该函数将80系XY坐标转换为经纬度。

**示例**：

```vbs
Sub OnInitScript()
mode =1’=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="80系XY坐标转经纬度"
SSProcess.AddInputParameter "中央子午线","117",0,"","设置当地所在中央子午线"
SSProcess.AddInputParameter "X坐标","3122389.551",0,"","输入要转换的X坐标值"
SSProcess.AddInputParameter "Y坐标","501523.305",0,"","输入要转换的Y坐标值"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSProcess.UpdateScriptDlgParameter 1
CentralMeridian = SSProcess.GetInputParameter ("中央子午线")
x = SSProcess.GetInputParameter ("X坐标")
y = SSProcess.GetInputParameter ("Y坐标")
SSProcess.xy80ToLongiLati CentralMeridian, x, y, pB, pL
Msgbox "PB:"&pB&chr(13)&"PL:"&pL
End Sub
Sub OnCancel()
End Sub
```


#### xyWGS84ToLongiLati WGS84 系 XY 坐标转换经纬度

**函数**：`xyWGS84ToLongiLati`

**功能**：WGS84系XY坐标转换经纬度

**语法**：

```vbs
SSProcess.xyWGS84ToLongiLati(lon0, x, y, pB, pL)
```

**参数**：

- lon0 中央子午线
- x WGS84系x坐标值
- y WGS84系x坐标值
- pB 纬度值
- pL 经度值

**说明**：

该函数将WGS84系XY坐标转换为经纬度。

**示例**：

```vbs
Sub OnInitScript()
mode =1’=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="WGS84系XY坐标转经纬度"
SSProcess.AddInputParameter "中央子午线","117",0,"","设置当地所在中央子午线"
SSProcess.AddInputParameter "X坐标","3122389.551",0,"","输入要转换的X坐标值"
SSProcess.AddInputParameter "Y坐标","501523.305",0,"","输入要转换的Y坐标值"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSProcess.UpdateScriptDlgParameter 1
CentralMeridian = SSProcess.GetInputParameter ("中央子午线")
x = SSProcess.GetInputParameter ("X坐标")
y = SSProcess.GetInputParameter ("Y坐标")
SSProcess.xyWGS84ToLongiLati CentralMeridian, x, y, pB, pL
Msgbox "PB:"&pB&chr(13)&"PL:"&pL
End Sub
Sub OnCancel()
End Sub
```


#### xyz2000ToLongiLati 2000 系 XYZ 坐标转换经纬度

**函数**：`xyz2000ToLongiLati`

**功能**：2000系XYZ坐标转换经纬度

**语法**：

```vbs
SSProcess.xyz2000ToLongiLati (lon0, x, y, h, pB, pL, pH)
```

**参数**：

- lon0 中央子午线经度，度为单位。
- X 高斯坐标X
- y 高斯坐标Y
- h 高斯坐标H
- pB 大地坐标纬度。
- pL 大地坐标经度。
- pH 大地坐标高程。

**说明**：

执行该函数，将把高斯坐标换算成该椭球对应投影带下的 CGCS2000椭球的大地坐标。
函数执行成功，返回CGCS2000椭球的大地坐标 。


#### xyz54ToLongiLati 54 系 XYZ 转换经纬度

**函数**：`xyz54ToLongiLati`

**功能**：54系XYZ转换经纬度

**语法**：

```vbs
SSProcess.xyz54ToLongiLati (lon0, x, y, z, pB, pL, pH)
```

**参数**：

- lon0 中央子午线
- x X坐标
- y Y坐标
- z Z坐标
- pB 经度
- pL 纬度
- pH 高程

**示例**：

```vbs
Sub OnClick()
Y =466000
X =3536000
SSProcess.xyz54ToLongiLati 117 , X, Y , z, pB , pL, pH
SSProcess.LongiLatiToxyzCGCS2000 117, pB, pL , pH , pX, pY, pZ
Msgbox "pB="&pB&"pL="& pL&"pH=" &pH &"pX=" &pX& "pY="&pY&"pZ="& pZ
End Sub
```


#### xyz80ToLongiLati 80 系 XYZ 坐标转换经纬度

**函数**：`xyz80ToLongiLati`

**功能**：四参数坐标转换

**语法**：

```vbs
SSProcess.xyz80ToLongiLati (lon0, x, y, z, pB, pL, pH)
```

**参数**：

- lon0 中央子午线
- x X坐标
- y Y坐标
- z Z坐标
- pB 经度
- pL 纬度
- pH 高程

**示例**：

```vbs
Sub OnClick()
Y =466000
X =3536000
SSProcess.xyz80ToLongiLati 117, X, Y , z, pB , pL, pH
SSProcess.LongiLatiToxyzCGCS2000 117, pB, pL , pH , pX, pY, pZ
Msgbox "pB="&pB&"pL="& pL&"pH=" &pH &"pX=" &pX& "pY="&pY&"pZ="& pZ
End Sub
```


#### xyzWGS84ToLongiLati WGS84 系 XYZ 坐标转换经纬度

**函数**：`xyzWGS84ToLongiLati`

**功能**：WGS84系XYZ坐标转换经纬度

**语法**：

```vbs
SSProcess.xyzWGS84ToLongiLati (lon0, x, y, z, pB, pL, pH)
```

**参数**：

- lon0 中央子午线
- X WGS84坐标系下，x坐标
- Y WGS84坐标系下，y坐标
- Z WGS84坐标系下，z坐标
- pB WGS84系点坐标对应经度
- pL WGS84系点坐标对应纬度
- pH WGS84系点坐标对应高程

**说明**：

此函数功能为将WGS84系点坐标转换为经纬度的表示方法。

**示例**：

```vbs
Sub OnClick()
SSProcess.xyzWGS84ToLongiLati 1,350000,2500000, 200 ,pB,pL ,pH
Msgbox "pB=" & pB & "*" &"pL=" & pL& "*"& "pH=" & pH
End Sub
```


---

### 数学函数

#### Cross_L 直线与线集求交

**函数**：`Cross_L`

**功能**：直线与线集求交。

**语法**：

```vbs
SSProcess.Cross_L ( x1, y1, x2, y2, crossIDs )
```

**参数**：

- x1 直线点1的x坐标
- y1 直线点1的y坐标
- x2 直线点2的x坐标
- y2 直线点2的y坐标
- crossIDs 与直线相交的线集ID

**说明**：

执行此函数可求直线与线集求交


#### Cross_P 两线求交

**函数**：`Cross_P`

**功能**：两线求交。

**语法**：

```vbs
SSProcess.Cross_P ( resx, resy, xy11x, xy11y, xy12x, xy12y,xy21x, xy21y, xy22x, xy22y )
```

**参数**：

- resx 返回交点的X坐标
- resy 返回交点的Y坐标
- xy11x 直线1点1的X坐标
- xy11y 直线1点1的Y坐标
- xy12x 直线1点2的X坐标
- xy12y 直线1点2的Y坐标
- xy21x 直线2点1的X坐标
- xy21y 直线2点1的Y坐标
- xy22x 直线2点2的X坐标
- xy22y 直线2点2的Y坐标

**说明**：

执行此函数，返回两条直线（段）的交点坐标。

**示例**：

```vbs
Sub OnClick()
Dim resx,resy
xy11x= 10
xy11y= 10
xy12x= 20
xy12y= 20
xy21x= 35
xy21y= 30
xy22x= 40
xy22y = 45
SSProcess.Cross_P resx, resy, xy11x , xy11y, xy12x, xy12y ,xy21x, xy21y, xy22x, xy22y
Msgbox resx&" "& resy
End Sub
```


#### DistPerpEnd 求点到线段的垂距，垂足点及相对关系

**函数**：`DistPerpEnd`

**功能**：求点到线段的垂距，垂足点及相对关系。

**语法**：

```vbs
SSProcess.DistPerpend ( resx, resy, pResRelation, x, y, x1, y1, x2, y2 )
```

**参数**：

- resx,resy 垂足点坐标
- pResRelation 点与垂足点的相对关系：
- -1 垂点在(p1,p2)的p1端外
- 1 垂点在(p1,p2)的p1点上
- 0 垂点在(p1,p2)的中间
- 2 垂点在(p1,p2)的p2点上
- -2 垂点在(p1,p2)的p2端外
- x,y 点坐标
- x1,y1、x2,y2 线段的两端点坐标

**说明**：

执行该函数，将返回垂足点坐标及点与垂足点之间的关系。

**示例**：

```vbs
Sub OnClick()
x=764.487:y=360.533:x1=760.204:y1=352.613:x2=765.535:y2=379.631
SSProcess.DistPerpend resx, resy, pResRelation, x, y, x1, y1, x2, y2
Msgbox resx&" "&resy&" "&pResRelation
End Sub
```


#### GetColorIndex 根据颜色值获取颜色号

**函数**：`GetColorIndex`

**功能**：根据颜色值获取颜色号。

**语法**：

```vbs
SSProcess.GetColorIndex ( rgbColor )
```

**参数**：

- rgbColor 颜色值
- 返回对应的颜色号

**示例**：

```vbs
Sub Onclick()
mycolor = RGB (255,255, 255)
colorNumber = SSProcess.GetColorIndex ( mycolor )'根据颜色值获取颜色号
Msgbox colorNumber
End Sub
```


#### GetColorValue 根据颜色号获取颜色值

**函数**：`GetColorValue`

**功能**：根据颜色号获取颜色值。

**语法**：

```vbs
SSProcess.GetColorValue ( indexColor )
```

**参数**：

- indexColor 颜色序号

**说明**：

根据颜色号获取颜色值

**示例**：

```vbs
Sub OnClick()
colorNumber = SSProcess.GetColorValue (7)
End Sub
```


#### IsPolygonInPolygon 判断多边形与多边形关系

**函数**：`IsPolygonInPolygon`

**功能**：判断多边形与多边形关系。

**语法**：

```vbs
SSProcess.IsPolygonInPolygon ( polygonID1, polygonID2, limit )
```

**参数**：

- polygonID1 多边形1的ID
- polygonID2 多边形2的ID
- limit 间隙最小距，米为单位

**说明**：

执行该函数，返回两个多边形的相互关系，返回值：0-两多边形无包含关系，无相交部分；2-两多边形为包含关系；3-两多边形相交。

**示例**：

```vbs
Sub OnClick()
polygonID1 =51
polygonID2 =39
limit =0.001
PolygonRelation =SSProcess.IsPolygonInPolygon( polygonID1, polygonID2, limit )
Msgbox PolygonRelation
End Sub
```


#### IsPolylineInPolygon 判断线与多边形关系

**函数**：`IsPolylineInPolygon`

**功能**：判断线与多边形关系。

**语法**：

```vbs
SSProcess.IsPolylineInPolygon ( lineID, polygonID, limit )
```

**参数**：

- lineID 线ID
- polygonID 多边形ID
- limit 间隙最小距，米为单位

**说明**：

执行该函数，判断线与多边形关系，返回值：0-线在多边形外，无相交部分；1-线在多边形内，无相交部分；2-线在多边形上，与多边形的某条边完全重合；3-线与多边形有相交部分。

**示例**：

```vbs
Sub OnClick()
'任意画线和多边形，假设线 ID为8，多边形ID为3，limit为1.
BZ = SSProcess.IsPolylineInPolygon (8, 3 , 1)
Msgbox BZ
End Sub
```


#### IsPtInPoly 判断点与多边形关系

**函数**：`IsPtInPoly`

**功能**：判断点与多边形关系。

**语法**：

```vbs
SSProcess.IsPtInPoly ( x, y, polygonID, limit )
```

**参数**：

- x 点的X坐标
- y 点的Y坐标
- polygonID 多边形的ID
- limit 限差

**说明**：

执行该函数，判断线与多边形关系，若返回值为0 ，表示点在多边形外。若返回值为1 ，表示点在多边形内。若返回值为2 ，表示点在多边形上。

**示例**：

```vbs
Sub OnClick()
pid=99  '点ID
SSProcess.getobjectpoint pid,0,x,y,z,ptype,name  '获取点坐标
PolygonID=23   '面ID
Limit=0.1  '限差
P2ARelation=SSProcess.IsPtInPoly(x,y,polygonID,limit)
Msgbox P2ARelation
End Sub
```


#### IsPtOnLine 判断点与线关系

**函数**：`IsPtOnLine`

**功能**：判断点与线关系。

**语法**：

```vbs
SSProcess.IsPtOnLine ( x, y, lineID, limit )
```

**参数**：

- x 点的X坐标
- y 点的y坐标
- lineID 线对象的ID号
- limit 间隙最小距，米为单位

**说明**：

使用此函数可以判断点与线的关系。返回值：0-点在线外；1-点在线上。

**示例**：

```vbs
Sub OnClick()
pid=99  '点ID
SSProcess.getobjectpoint pid,0,x,y,z,ptype,name  '获取点坐标
lineID=23 '线ID
Limit=0.1 '限差
aa= SSProcess.IsPtOnLine( x , y, lineID , limit )
Msgbox aa
End Sub
```


#### PerpEnd_P 求垂足点

**函数**：`PerpEnd_P`

**功能**：求垂足点。

**语法**：

```vbs
SSProcess.Perpend_P ( resx, resy, xyx, xyy, xy1x, xy1y,xy2x, xy2y )
```

**参数**：

- resx 生成垂足点的X坐标
- resy 生成垂足点的Y坐标
- xyx 直线外点的X坐标
- xyy 直线外点的Y坐标
- xy1x 直线第一点X坐标
- xy1y 直线第一点Y坐标
- xy2x 直线第二点X坐标
- xy2y 直线第二点Y坐标

**说明**：

求直线外一点到直线的垂足坐标。

**示例**：

```vbs
Sub OnClick()
x=764.487:y=360.533:x1=760.204:y1=352.613:x2=765.535:y2=379.631
SSProcess. Perpend_P resx, resy, x, y, x1, y1, x2, y2
Msgbox resx&" "&resy
End Sub
```


---

### 系统设置函数

#### CloseScriptDlg 关闭脚本运行对话框

**函数**：`CloseScriptDlg`

**功能**：关闭脚本运行对话框。

**语法**：

```vbs
SSProcess.CloseScriptDlg()
```

**说明**：

该函数用于关闭脚本运行对话框。

**示例**：

```vbs
Sub OnInitScript()
mode =1’=0 无参数对话框 =1 有参数对话框
title="测试"
'设置对话框
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSProcess.CloseScriptDlg()'关闭脚本运行对话框
End Sub
Sub OnCancel()
End Sub
```


#### GetCursorStatus 获取当前光标状态

**函数**：`GetCursorStatus`

**功能**：获取当前光标状态。

**语法**：

```vbs
SSProcess.GetCursorStatus()
```

**说明**：

使用此函数可以判断当前光标的状态，如果光标为选择状态则返回值为0，如果光标为输入状态则返回值为6。

**示例**：

```vbs
Sub OnClick()
oldCursorStatus = SSProcess.GetCursorStatus()
Msgbox oldCursorStatus
End Sub
```


#### SetCursorStatus 设置当前光标状态

**函数**：`SetCursorStatus`

**功能**：设置当前光标状态。

**语法**：

```vbs
SSProcess.SetCursorStatus (status)
```

**参数**：

- status 光标状态，取值为0,1,2,3,4,5,6,7

**说明**：

执行该函数设置光标状态。

**示例**：

```vbs
Sub OnClick()
oldCursorStatus = SSProcess.GetCursorStatus
SSProcess.SetCursorStatus 3
newCursorStatus = SSProcess.GetCursorStatus
Msgbox "oldCursorStatus=" & oldCursorStatus & "newCursorStatus =" & newCursorStatus
End Sub
```


#### ShowScriptDlg 显示脚本运行对话框

**函数**：`ShowScriptDlg`

**功能**：显示脚本运行对话框。

**语法**：

```vbs
SSProcess.ShowScriptDlg (mode, title)
```

**参数**：

- mode 对话框模式；0无参数对话框，1有参数对话框
- title 对话框标题

**说明**：

执行此函数弹出参数对话框，对话框参数使用AddInputParameter 进行添加。

**示例**：

```vbs
Sub OnInitScript()
mode = 1’=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "范围线编码", "1", 0, "", ""
SSProcess.AddInputParameter "输出路径", "C:\", 12, "", ""
SSProcess.ShowScriptDlg mode,title
End Sub
```


#### ShowScriptUserDefDlg 显示自定义脚本运行对话框

**函数**：`ShowScriptUserDefDlg`

**功能**：显示自定义脚本运行对话框。

**语法**：

```vbs
SSProcess.ShowScriptUserDefDlg(title, dlgTemplateName, dlgWidth, dlgHeight, colCount, titleWidth, valueWidth)
```

**参数**：

- title 对话框标题名
- dlgTemplateName 对话框模板名，DLG格式对话框名
- dlgWidth 对话框初始宽度
- dlgHeight 对话框初始高度
- colCount 对话框列数
- titleWidth 对话框标题宽度
- valueWidth 对话框值输入栏宽度

**说明**：

执行此函数可以调用使用EPS定制的DLG对话框样式，DLG对话框放置的目录为与模板同目录的同名文件夹内。

**示例**：

```vbs
Sub Onclick()
title="承包地块调查"
dlgTemplateName="RLCM_承包地块面属性表"
dlgWidth=650
dlgHeight=600
colCount=2
titleWidth=150
valueWidth=150
result = SSProcess.ShowScriptUserDefDlg(title, dlgTemplateName, dlgWidth, dlgHeight, colCount, titleWidth, valueWidth)
End Sub
```


#### UpdateScriptDlgParameter 更新脚本运行对话框参数

**函数**：`UpdateScriptDlgParameter`

**功能**：更新脚本运行对话框参数。

**语法**：

```vbs
SSProcess.UpdateScriptDlgParameter (mode)
```

**参数**：

- mode 参数更新方式，0表示更新到内存，1表示不更新。

**说明**：

该函数为系统设置函数，用于更新脚本运行对话框参数。

**示例**：

```vbs
将脚本运行对话框参数更新到内存。
Sub OnInitScript()
mode = 1 ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "太阳" , "1!", 0 , "", "" '添加录入参数
SSProcess.AddInputParameter "月亮" , "1!", 0 , "", ""
SSProcess.ShowScriptDlg mode , title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "太阳" , "A1!", 0 , "", ""
SSProcess.AddInputParameter "月亮" , "B1!", 0 , "", ""
SSProcess.UpdateScriptDlgParameter
End Sub
Sub OnCancel()
End Sub
```


---

### 系统函数

#### GetCurrentTime1 获取当前时间

**函数**：`GetCurrentTime1`

**功能**：获取当前时间。

**语法**：

```vbs
SSProcess.GetCurrentTime1()
```

**说明**：

执行该函数，将获取当前时间

**示例**：

```vbs
Sub Onclick()
my_time =100  '设置单步所需时间（毫秒）
SSProcess.EpsProgressSetStyle 1  '设定滚动条样式
SSProcess.EpsProgressCreate 100," 正在转换..."  '创建滚动条
SSProcess.EpsProgressSetStep 1  '设置滚动条步长
start_time = SSProcess.GetCurrentTime1 '获取开始时间
For i = 0 To 100
SSProcess.EpsProgressStepIt
SSProcess.Sleep my_time
Next
finish_time = SSProcess.GetCurrentTime1  '获取结束时间
SSProcess.EpsProgressDelete  '删除滚动条
SSProcess.RefreshView '刷新屏幕
Msgbox "时间差： "&finish_time - start_time,64, "消息"
End Sub
```


#### ShowSelectMenu 弹出选择菜单

**函数**：`ShowSelectMenu`

**功能**：弹出选择菜单。

**语法**：

```vbs
SSProcess.ShowSelectMenu(selParameters )
```

**参数**：

- selParameters 选择菜单列表，多个参数之间用“,”分割

**说明**：

执行该函数，将弹出选择菜单

**示例**：

```vbs
Sub OnClick()
str = SSProcess.ShowSelectMenu ("A,B,C,D,E,F" )
Msgbox str
End Sub
```


#### ShowSelectMenu1 弹出分级选择菜单

**函数**：`ShowSelectMenu1`

**功能**：获取当前时间。

**语法**：

```vbs
SSProcess.ShowSelectMenu1( title, selParameters, color1, color2)
```

**参数**：

- title 菜单标题
- selParameters 选择菜单列表 ，多个参数之间用“，”分割
- color1 背景颜色1（与背景颜色2形成渐变背景）
- color2 背景颜色2

**说明**：

执行该函数，将弹出分级选择菜单

**示例**：

```vbs
Sub OnClick()
str = SSProcess.ShowSelectMenu1 ("选择","A,B,C,D,E,F",RGB(255,0,0),RGB(255,255,0) )
Msgbox str
End Sub
```


#### Sleep 休眠指定时间

**函数**：`Sleep`

**功能**：休眠指定时间。

**语法**：

```vbs
SSProcess.Sleep(milliseconds)
```

**参数**：

- milliseconds 休眠的时间，单位为毫秒

**说明**：

执行该函数，将当前执行挂起一定时间。
注意：该函数的参数milliseconds 为毫秒，故欲延迟1秒需设置为1000。

**示例**：

```vbs
Sub OnClick()
Dim strFilenames (1000)
mdbName=SSProcess.GetTemplateFileName
strDicPath=left(mdbName ,len(mdbName )-4)
strListname=strDicPath&"/LIST.TXT"
Set ws=CreateObject("Wscript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject")
ws.run "%comspec% /c cd "+chr(34)& strDicPath&chr(34)+ _
" & dir *.DIC /b >"+chr(34)& strListname&chr(34)
SSProcess.Sleep 2000
Set listTXT=fso.OpenTextFile(strListname , 1)
nCount=0
Do While Not listTXT.AtEndOfStream
listLine =listTXT.ReadLine
If listLine <>"" Then
strFilenames (i)=trim (listLine)
i =i+ 1
nCount =nCount+ 1
End If
Loop
listTXT.Close
End Sub
```


---

## Function 对象参考

### 系统消息函数

#### incluede 引用外部脚本 #

**函数**：`#incluede`

**功能**：#inclued函数

**示例**：

```vbs
用VBScript建立脚本“msgbox.vbs”，存放在“PDB出入库”的脚本组中；另建立脚本“输入输出.vbs”，脚本代码如下，执行脚本“输入输出.vbs”，可运行外部脚本“Msgbox.vbs”。
脚本一“msgbox.vbs”
Sub OnClick()
temp = inputbox( "请输入华氏温度" ,"温度转换",80)
End Sub
脚本二“输入输出.vbs”
RunScript "VBScript", "PDB出入库" ,"Msgbox"
#include "msgbox.vbs"
Sub OnClick()
Msgbox "温度为"&Celsius( temp)&"摄氏度"
End Sub
Function Celsius(fDegrees )
Celsius = (fDegrees - 32) * 5 / 9
End Function
```


#### BeforeExportData 输出外部数据之前

**函数**：`BeforeExportData`

**功能**：写入外部数据之前。

**说明**：

在输出外部数据之前调用该函数对数据进行处理后再输出。

**示例**：

```vbs
输出ArcGIS的MDB格式前将高斯直角坐标转为经纬度
Function BeforeExportData()
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount-1
pointCount = SSProcess.GetSelGeoPointCount (i)
SSProcess.LockSelGeoPoint i, 1
For j=0 To pointCount-1
SSProcess.GetSelGeoPoint i, j, x, y, z, ptype, pname
SSProcess.xy80ToLongiLati 114, x, y, y0, x0
SSProcess.SetSelGeoPoint i, j, x0, y0, z, ptype, pname
Next
SSProcess.UpdateSelGeoPoint i
SSProcess.LockSelGeoPoint i, 0
Next
End Function
Sub OnClick()
'清空转换参数
SSProcess.ClearDataXParameter
'设置导入文件格式为ArcGIS PDB
SSProcess.SetDataXParameter "DataType", "22"
SSProcess.SetDataXParameter "AddSystemFieldMode", "0"
'开始导入数据
SSProcess.ExportData
End Sub
```


#### BeforeSaveImportData 导入外部数据存库之前

**函数**：`BeforeSaveImportData`

**功能**：导入外部数据存库之前。

**示例**：

```vbs
经纬度ArcGIS的MDB格式数据导入，并在导入前把经纬度换算为高斯直角坐标
Function BeforeSaveImportData()
geoCount = SSProcess.GetSelGeoCount
For i=0 To geoCount-1
pointCount = SSProcess.GetSelGeoPointCount (i)
SSProcess.LockSelGeoPoint i, 1
For j=0 To pointCount-1
SSProcess.GetSelGeoPoint i, j, x, y, z, ptype, pname
SSProcess.LongiLatiToxy80 114, y, x, x0, y0
SSProcess.SetSelGeoPoint i, j, x0, y0, z, ptype, pname
Next
SSProcess.UpdateSelGeoPoint i
SSProcess.LockSelGeoPoint i, 0
Next
End Function
Sub OnClick()
'清空转换参数
SSProcess.ClearDataXParameter
'设置导入文件格式为ArcGIS PDB
SSProcess.SetDataXParameter "DataType", "22"
SSProcess.SetDataXParameter "SaveAttrToMemoData", "1"
'开始导入数据
SSProcess.ImportData
End Sub
```


#### LineProcess 编辑功能快捷键消息

**函数**：`LineProcess`

**功能**：编辑功能快捷键消息。

**参数**：

- nFlags 快捷键标识

**说明**：

通过该函数可定义用户按下相应快捷键时，做出何种相应。

**示例**：

```vbs
colorchange= False
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="请按W 键，会有惊喜哟！ "
SSProcess.ShowScriptDlg mode ,title
End Sub
Function LineProcess(nFlags )
If nFlags = 140 Then
Msgbox "LineProcess测试：你按了 W 键！"
Else
Msgbox "太淘气了为什么不按 W键~"
End If
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
4.2.1.5 点击对话框取消键消息
OnCancel函数
OnCancel函数
点击对话框取消键消息。
OnCancel()
```

**说明**：

点击对话框取消键消息。

**示例**：

```vbs
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
Msgbox "OnCancel()"
End Sub
```


#### OnClick 有模脚本入口函数

**函数**：`OnClick`

**功能**：有模脚本入口函数。


#### OnContextMenu 上下文菜单消息

**函数**：`OnContextMenu`

**功能**：上下文菜单消息。

**参数**：

- x
- y


#### OnDraw 系统绘画消息

**函数**：`OnDraw`

**功能**：系统绘画消息。

**参数**：

- pdc 屏幕指针


#### OnExitScript 无模脚本出口函数

**函数**：`OnExitScript`

**功能**：无模脚本出口函数。

**说明**：

此函数通常与入口函数成对出现，可在函数内设置退出对话框时触发的动作，如：清空所填参数、弹出对话框等操作。

**示例**：

```vbs
退出无模脚本弹出提示并清空参数
Sub OnExitScript()
Msgbox "退出对话框，清空参数" '弹窗显示内容
SSProcess.ClearInputParameter  '清空所填参数
End Sub
```


#### OnGridItemChanged 表格控件行列值改变消息

**函数**：`OnGridItemChanged`

**功能**：表格控件行列值改变消息。

**参数**：

- Row 行
- col 列


#### OnGridItemDbClick 表格控件双击消息

**函数**：`OnGridItemDbClick`

**功能**：表格控件双击消息。

**参数**：

- row


#### OnGridSelChanged 表格控件选择行改变消息

**函数**：`OnGridSelChanged`

**功能**：表格控件选择行改变消息。

**参数**：

- Row 行


#### OnInitScript 无模脚本入口函数

**函数**：`OnInitScript`

**功能**：无模脚本入口函数。

**参数**：

- mode 有无参数对话框， =1 有参数对话框； =0 无参数对话框
- title 参数对话框名称。

**说明**：

此函数通常位于所创建的无模脚本开始处，用于在无模脚本中设置是否弹出对话框、窗口名称、窗口中参数内容等。

**示例**：

```vbs
输入点位坐标
Sub OnInitScript()
mode = 1  ‘=0 无参数对话框 =1 有参数对话框
title="输入坐标点击确定(有带号坐标)" '对话框名称
SSProcess.ClearInputParameter '清空所填参数
SSProcess.AddInputParameter "X(北25..)","" , 0 ,"" , "2500000左右" '参数名称及注释
SSProcess.AddInputParameter "Y(东36..)","" , 0, "" , "36000000左右" '参数名称及注释
SSProcess.ShowScriptDlg mode,title
End Sub
```


#### OnKeyDown 键盘按键击下消息

**函数**：`OnKeyDown`

**功能**：键盘按键击下消息。

**参数**：

- nChar 键盘的字符码值
- nRepCnt 击键的次数
- nFlags 包含扫描码，转换码等
| 参数 | 说明 |
|---|---|
| nFlags | 键盘 |
| ~ | 41 |
| 1 | 2 |
| 2 | 3 |
| 3 | 4 |
| 4 | 5 |
| 5 | 6 |
| 6 | 7 |
| 7 | 8 |
| 8 | 9 |
| 9 | 10 |
| 0 | 11 |
| - | 12 |
| + | 13 |
| Backspace | 14 |
| home | 327 |
| Q | 16 |
| W | 17 |
| E | 18 |
| R | 19 |
| T | 20 |
| Y | 21 |
| U | 22 |
| I | 23 |
| O | 24 |
| P | 25 |
| [ | 26 |
| ] | 27 |
| \ | 43 |
| PgUp | 329 |
| CapsLock | 352 |
| A | 30 |
| S | 31 |
| D | 32 |
| F | 33 |
| G | 34 |
| H | 35 |
| J | 36 |
| K | 37 |
| L | 38 |
| : | 39 |
| " | 40 |
| Enter | 28 |
| PgDn | 337 |
| 左 Shift | 42 |
| Z | 44 |
| X | 45 |
| C | 46 |
| V | 47 |
| B | 48 |
| N | 49 |
| M | 50 |
| < | 51 |
| > | 52 |
| ? | 53 |
| 右 shift | 54 |
| 上箭头 | 328 |
| End | 335 |
| Ctrl | 29 |
| 左箭头 | 331 |
| 下箭头 | 336 |
| 右箭头 | 333 |

**说明**：

使用此函数可用键盘上的任意键执行换颜色。

**示例**：

```vbs
将用“Y”键来执行1线换颜色：
colorchange=False
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="更换1线颜色！"
SSProcess.ShowScriptDlg mode ,title
End Sub
Function OnKeyDown(nChar , nRepCnt, nFlags)
Msgbox nFlags
If nFlags = 21 Then
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
nCount =SSProcess.GetSelGeoCount
For i = 0 To nCount -1
SSProcess.SetSelGeoValue i , "SSObj_Color", "RGB(255,255,0)"
Next
SSProcess.RefreshView
End If
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### OnKeyUp 键盘按键抬起消息

**函数**：`OnKeyUp`

**功能**：键盘按键抬起消息。

**参数**：

- nChar 键盘的字符码值
- nRepCnt 击键的次数
- nFlags 包含扫描码，转换码等：
| 参数 | 说明 |
|---|---|
| nFlags | 键盘 |
| ~ | 49153 |
| 1 | 49154 |
| 2 | 49155 |
| 3 | 49156 |
| 4 | 49157 |
| 5 | 49158 |
| 6 | 49159 |
| 7 | 49160 |
| 8 | 49161 |
| 9 | 49162 |
| 0 | 49163 |
| - | 49164 |
| + | 49165 |
| Backspace | 49166 |
| home | 49479 |
| Q | 49168 |
| W | 49169 |
| E | 49170 |
| R | 49171 |
| T | 49172 |
| Y | 49173 |
| U | 49174 |
| I | 49175 |
| O | 49176 |
| P | 49177 |
| [ | 49178 |
| ] | 49179 |
| \ | 49196 |
| PgUp | 49481 |
| CapsLock | 49210 |
| A | 49182 |
| S | 49183 |
| D | 49184 |
| F | 49185 |
| G | 49186 |
| H | 49187 |
| J | 49188 |
| K | 49189 |
| L | 49190 |
| : | 49191 |
| " | 49192 |
| Enter | 49180 |
| PgDn | 49489 |
| 左 Shift | 49194 |
| Z | 49196 |
| X | 49197 |
| C | 49198 |
| V | 49199 |
| B | 49200 |
| N | 49201 |
| M | 49202 |
| < | 49203 |
| > | 49204 |
| ? | 49205 |
| 又 shift | 49206 |
| 上箭头 | 49480 |
| End | 49487 |
| Ctrl | 49181 |
| 左箭头 | 49483 |
| 下箭头 | 49488 |
| 右箭头 | 49485 |

**说明**：

使用此函数可用键盘上的任意键执行换颜色。

**示例**：

```vbs
将用“Y”键来执行1线换颜色：
colorchange= False
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="更换1线颜色！"
SSProcess.ShowScriptDlg mode ,title
End Sub
Function OnKeyUp(nChar , nRepCnt, nFlags)
Msgbox nFlags
If nFlags = 49173 Then
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "1"
SSProcess.SelectFilter
nCount =SSProcess.GetSelGeoCount
For i = 0 To nCount -1
SSProcess.SetSelGeoValue i , "SSObj_Color", "RGB(255,0,0)"
Next
SSProcess.RefreshView
End If
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### OnLButtonDblClk 鼠标左键双击消息

**函数**：`OnLButtonDblClk`

**功能**：鼠标左键双击消息。

**参数**：

- x 返回屏幕横坐标
- y 返回屏幕纵坐标
- spx 返回实际的空间横坐标
- spy 返回实际的空间纵坐标
- flags 标记


#### OnLButtonDown 鼠标左键击下消息

**函数**：`OnLButtonDown`

**功能**：鼠标左键击下消息。

**参数**：

- x 返回屏幕横坐标
- y 返回屏幕纵坐标
- spx 返回实际的空间横坐标
- spy 返回实际的空间纵坐标
- flags 标记

**说明**：

该函数为系统消息函数，用于无模脚本中，鼠标左键击下即触发该函数。

**示例**：

```vbs
Dim arX(10000),arY(10000), arX0(10000),arY0(10000),arXYCount
Sub OnInitScript()
mode =0’=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode,title
End Sub
Function OnLButtonDown(x, y, spx, spy, flags)
arX(arXYCount)=spx:arY(arXYCount)= spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x, y, spx, spy, flags)
If arXYCount>0 Then
SSProcess.PasteBackgroundImage
If arXYCount =1 Then
SSView.DrawLine pDc, arX(0), arY(0),spx,spy, RGB(255,0,255),300, linestyle
Elseif arXYCount >1 Then
SSView.DrawLine pDc, arX(arXYCount-1), arY(arXYCount-1),spx,spy, RGB(255,0,255),300, linestyle
draws spx,spy
End If
End If
End Function
Function draws (spx,spy)
For i =1 To arXYCount-1
SSView.DrawLine pDc, arX(i-1), arY(i-1), arX(i), arY(i), RGB(255,0,255),300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### OnLButtonUp 鼠标左键抬起消息

**函数**：`OnLButtonUp`

**功能**：鼠标左键抬起消息。

**参数**：

- x 返回屏幕横坐标
- y 返回屏幕纵坐标
- spx 返回实际的空间横坐标
- spy 返回实际的空间纵坐标
- flags 标记

**说明**：

该函数为系统消息函数，用于无模脚本中，鼠标左键抬起即触发该函数。

**示例**：

```vbs
Dim arX(10000),arY(10000), arX0(10000),arY0(10000),arXYCount
Sub OnInitScript()
mode =0’=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode,title
End Sub
Function OnLButtonUp(x, y, spx, spy, flags)
arX(arXYCount)=spx:arY(arXYCount)= spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x, y, spx, spy, flags)
If arXYCount>0 Then
SSProcess.PasteBackgroundImage
If arXYCount =1 Then
SSView.DrawLine pDc, arX(0), arY(0),spx,spy, RGB(255,0,255),300, linestyle
elseIf arXYCount >1 Then
SSView.DrawLine pDc, arX(arXYCount-1), arY(arXYCount-1),spx,spy, RGB(255,0,255),300, linestyle
draws spx,spy
End If
End If
End Function
Function draws (spx,spy)
For i =1to arXYCount-1
SSView.DrawLine pDc, arX(i-1), arY(i-1), arX(i), arY(i), RGB(255,0,255),300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### OnMButtonDown 鼠标中键击下消息

**函数**：`OnMButtonDown`

**功能**：鼠标中键击下消息。

**参数**：

- x 返回屏幕横坐标
- y 返回屏幕纵坐标
- spx 返回实际的空间横坐标
- spy 返回实际的空间纵坐标
- flags 标记

**说明**：

该函数为系统消息函数，用于无模脚本中，鼠标中键击下即触发该函数。

**示例**：

```vbs
Dim arX(10000),arY(10000), arX0(10000),arY0(10000),arXYCount
Sub OnInitScript()
mode =0’=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode,title
End Sub
Function OnMButtonDown(x, y, spx, spy, flags)
arX(arXYCount)=spx:arY(arXYCount)= spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x, y, spx, spy, flags)
If arXYCount>0 Then
SSProcess.PasteBackgroundImage
If arXYCount =1 Then
SSView.DrawLine pDc, arX(0), arY(0),spx,spy, RGB(255,0,255),300, linestyle
Elseif arXYCount >1 Then
SSView.DrawLine pDc, arX(arXYCount-1), arY(arXYCount-1),spx,spy, RGB(255,0,255),300, linestyle
draws spx,spy
End If
End If
End Function
Function draws (spx,spy)
For i =1 To arXYCount-1
SSView.DrawLine pDc, arX(i-1), arY(i-1), arX(i), arY(i), RGB(255,0,255),300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### OnMButtonUp 鼠标中键抬起消息

**函数**：`OnMButtonUp`

**功能**：鼠标中键抬起消息。

**参数**：

- x 返回屏幕横坐标
- y 返回屏幕纵坐标
- spx 返回实际的空间横坐标
- spy 返回实际的空间纵坐标
- flags 标记

**说明**：

该函数为系统消息函数，用于无模脚本中，鼠标中键抬起即触发该函数。


#### OnMouseMove 鼠标移动消息

**函数**：`OnMouseMove`

**功能**：鼠标移动消息。

**参数**：

- x 返回屏幕横坐标
- y 返回屏幕纵坐标
- spx 返回实际的空间横坐标
- spy 返回实际的空间纵坐标
- flags 标记

**说明**：

该函数为系统消息函数，鼠标移动即触发函数，一般不单独适用。

**示例**：

```vbs
Dim a ,b,c,d ,e
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
title="返回坐标"
SSProcess.ShowScriptDlg mode,title
End Sub
Function OnLButtonDown(x , y, spx , spy, flags)
Msgbox a
Msgbox b
Msgbox c
Msgbox d
Msgbox e
End Function
Function OnMouseMove(x , y, spx , spy, flags)
a= x
b= y
c= spx
d= spy
e= flags
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### OnMouseWheel 鼠标滚轮消息

**函数**：`OnMouseWheel`

**功能**：鼠标滚轮消息。

**参数**：

- x 屏幕横坐标
- y 屏幕纵坐标
- spx 实际横坐标
- spy 实际纵坐标
- flags 判断标志
- zDelta 鼠标滚轮滚动一下高程的变化量

**说明**：

该函数为系统消息函数，鼠标移动即触发函数，一般不单独适用。

**示例**：

```vbs
Dim  a ,b,c,d ,e
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
title="返回坐标"
SSProcess.ShowScriptDlg mode,title
End Sub
Function OnLButtonDown(x , y, spx , spy, flags)
Msgbox a
Msgbox b
Msgbox c
Msgbox d
Msgbox e
End Function
Function OnMouseWheel(x , y, spx , spy, flags, zDelta )
a= x
b= y
c= spx
d= spy
e= zDelta
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### OnOK 点击对话框确定键消息

**函数**：`OnOK`

**功能**：点击对话框确定键消息。

**说明**：

响应点击对话框确定键消息

**示例**：

```vbs
Sub OnInitScript()
mode = 0  ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
Msgbox "OnOk()"
End Sub
Sub OnCancel()
End Sub
```


#### OnParameterSetup 参数设置消息

**函数**：`OnParameterSetup`

**功能**：参数设置消息。


#### OnPropertyChanged 参数发生改变

**函数**：`OnPropertyChanged`

**功能**：参数发生改变。


#### OnRButtonDblClk 鼠标右键双击消息

**函数**：`OnRButtonDblClk`

**功能**：鼠标右键双击消息。

**参数**：

- x 返回屏幕横坐标
- y 返回屏幕纵坐标
- spx 返回实际的空间横坐标
- spy 返回实际的空间纵坐标
- flags 标记


#### OnRButtonDown 鼠标右键击下消息

**函数**：`OnRButtonDown`

**功能**：鼠标右键击下消息。

**参数**：

- x 返回屏幕横坐标
- y 返回屏幕纵坐标
- spx 返回实际的空间横坐标
- spy 返回实际的空间纵坐标
- flags 标记

**说明**：

执行该语句，可以获取鼠标点右键落下时刻屏幕坐标和鼠标点位置数据坐标，主要应用在无模脚本中，在做交互式操作时使用。

**示例**：

```vbs
Dim arX(10000),arY(10000), arX0(10000),arY0(10000),arXYCount
Sub OnInitScript()
mode =0 ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode,title
End Sub
Function OnRButtonDown (x , y, spx , spy, flags)
arX(arXYCount)=spx:arY(arXYCount)= spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x, y, spx, spy, flags)
If arXYCount>0 Then
SSProcess.PasteBackgroundImage
If arXYCount =1 Then
SSView.DrawLine pDc, arX(0), arY(0),spx,spy, RGB(255,0,255),300, linestyle
Elseif arXYCount >1 Then
SSView.DrawLine pDc, arX(arXYCount-1), arY(arXYCount-1),spx,spy, RGB(255,0,255),300, linestyle
draws spx,spy
End If
End If
End Function
Function draws (spx,spy)
For i =1 To arXYCount-1
SSView.DrawLine pDc, arX(i-1), arY(i-1), arX(i), arY(i), RGB(255,0,255),300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### OnRButtonUp 鼠标右键抬起消息

**函数**：`OnRButtonUp`

**功能**：鼠标右键抬起消息。

**参数**：

- X 返回的屏幕X坐标
- Y 返回的屏幕Y坐标
- Spx 返回当前鼠标点位置数据X坐标
- Spy 返回当前鼠标点位置数据Y坐标
- Flags 标记

**说明**：

执行该语句，可以获取鼠标点右键抬起时刻屏幕坐标和鼠标点位置数据坐标，主要应用在无模脚本中，在做交互式操作时使用。

**示例**：

```vbs
Dim arX(10000),arY(10000), arX0(10000),arY0(10000),arXYCount
Sub OnInitScript()
mode =0 ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode,title
End Sub
Function OnRButtonUp(x , y, spx , spy, flags)
arX(arXYCount)=spx:arY(arXYCount)= spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x, y, spx, spy, flags)
If arXYCount>0 Then
SSProcess.PasteBackgroundImage
If arXYCount =1 Then
SSView.DrawLine pDc, arX(0), arY(0),spx,spy, RGB(255,0,255),300, linestyle
Elseif arXYCount >1 Then
SSView.DrawLine pDc, arX(arXYCount-1), arY(arXYCount-1),spx,spy, RGB(255,0,255),300, linestyle
draws spx,spy
End If
End If
End Function
Function draws (spx,spy)
For i =1to arXYCount-1
SSView.DrawLine pDc, arX(i-1), arY(i-1), arX(i), arY(i), RGB(255,0,255),300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### OnSelectionChange 选择集对象改变消息

**函数**：`OnSelectionChange`

**功能**：选择集对象改变消息。

**说明**：

在无模脚本中，加入此函数，当选择集对象发生变化时自动进入此函数。

**示例**：

```vbs
当选择集对象发生变化后自动弹出提示。
Sub OnInitScript()
mode = 0
title="功能名称"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
Sub OnSelectionChange()
Msgbox "OK!"
End Sub
```


#### OnTimer 响应定时器

**函数**：`OnTimer`

**功能**：响应定时器。

**参数**：

- nIDEvent 定时器标示

**说明**：

使用此函数将执行定时器

**示例**：

```vbs
包含执行定时器。
Sub OnInitScript()
Mode = 0
title=" 节目号抽奖程序 "
SSProcess.ShowScriptDlg mode,title
SSProcess.SetTimer 128801, 2
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
Function OnTimer( nIDEvent )
End Function
```


#### RunScript 运行外部脚本

**函数**：`RunScript`

**功能**：运行外部脚本。

**参数**：

- language 语言名称
- scriptGroup 脚本分组名称
- scriptName 脚本名称

**说明**：

在本脚本文件中引用已编写好的脚本文件。（详见EPS脚本语言规则）

**示例**：

```vbs
用VBScript建立脚本“msgbox.vbs”，存放在“PDB出入库”的脚本组中；另建立脚本“输入输出.vbs”，脚本代码如下，执行脚本“输入输出.vbs”，可运行外部脚本“msgbox.vbs”。
脚本一“msgbox.vbs”
Sub OnClick()
temp = inputbox( "请输入华氏温度" ,"温度转换",80)
End Sub
脚本二“输入输出.vbs”
RunScript "VBScript", "PDB出入库" ,"msgbox"
#include "msgbox.vbs"
Sub OnClick()
Msgbox "温度为"&Celsius( temp)&"摄氏度"
End Sub
Function Celsius(fDegrees )
Celsius = (fDegrees - 32) * 5 / 9
End Function
```


---

## SSView 对象参考

### 屏幕绘制函数

#### Draw3DLine 画 3 维线

**函数**：`Draw3DLine`

**功能**：画三维线。

**语法**：

```vbs
SSView.Draw3DLine(pDc,x0,y0 ,z0,x1 ,y1,z1 ,color, width,linestyle)
```

**参数**：

- pDc 指定的窗口DC
- x0 绘制线段的第一点X坐标
- y0 绘制线段的第一点y坐标
- z0 绘制线段的第一点Z坐标
- x1 绘制线段的第二点X坐标
- y1 绘制线段的第二点Y坐标
- z1 绘制线段的第二点Z坐标
- color 颜色值
- width 线宽
- linestyle 线型

**说明**：

使用此函数将在指定位置绘制带三维坐标的线段。

**示例**：

```vbs
通过鼠标移动绘制线段：
Function OnMouseMove(x, y,spx , spy, flags)
x0=spx:y0 = spy:z0=spy: x1= spx+1000:y1= spy +1000:x2 = spx: y2= spy+1000:z2=spy+1000
color=255
width=10
linestyle=0
SSView.Draw3DLine pDc,x0,y0 ,z0,x1 ,y1,z1 ,color, width,linestyle
End Function
Sub OnInitScript()
mode =0  ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### DrawArc 画圆弧

**函数**：`DrawArc`

**功能**：画圆弧。

**语法**：

```vbs
SSView.DrawArc(pDc,x0,y0 ,x1,y1 ,x2,y2 ,color, width,linestyle)
```

**参数**：

- pDc 指定的窗口DC
- x0 绘制圆弧的第一点X坐标
- y0 绘制圆弧的第一点y坐标
- x1 绘制圆弧的第二点X坐标
- y1 绘制圆弧的第二点Y坐标
- x2 绘制圆弧的第三点X坐标
- y2 绘制圆弧的第三点Y坐标
- color 颜色值
- width 线宽
- linestyle 线型

**说明**：

使用此函数将在指定位置绘制圆弧。

**示例**：

```vbs
通过鼠标移动绘制圆弧：
Function OnMouseMove(x , y,spx , spy, flags)
x0=spx:y0 = spy: x1= spx+1000:y1= spy +1000:x2 = spx: y2= spy+1000
color=255
width=10
linestyle=0
SSView.DrawArc pDc,x0,y0 ,x1,y1 ,x2,y2 ,color, width,linestyle
End Function
Sub OnInitScript()
mode =0  ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### DrawCircle 画圆

**函数**：`DrawCircle`

**功能**：画圆。

**语法**：

```vbs
SSView.DrawCircle(pDc ,x0,y0 ,x1,y1 ,x2,y2 ,color, width,linestyle)
```

**参数**：

- pDc 指定的窗口DC
- x0 绘制圆的第一点X坐标
- y0 绘制圆的第一点y坐标
- x1 绘制圆的第二点X坐标
- y1 绘制圆的第二点Y坐标
- x2 绘制圆的第三点X坐标
- y2 绘制圆的第三点Y坐标
- color 颜色值
- width 线宽
- linestyle 线型

**说明**：

使用此函数将在指定位置绘制圆。

**示例**：

```vbs
通过鼠标移动绘制圆：
Function OnMouseMove(x , y,spx , spy, flags)
x0=spx:y0 = spy: x1= spx+1000:y1= spy +1000:x2 = spx: y2= spy+1000
color=255
width=10
linestyle=0
SSView.DrawCircle pDc ,x0,y0 ,x1,y1 ,x2,y2 ,color, width,linestyle
End Function
Sub OnInitScript()
mode =0  ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### pDc 画图片 DrawImage

**函数**：`DrawImage`

**功能**：画图片。

**语法**：

```vbs
SSView.DrawImage(pDc,imgType,imgName ,minX, minY,maxX,maxY,width ,height, drawToBk)
```

**参数**：

- pDc 指定的窗口DC
- imgType 图片类型
- imgName 图片路径名称
- minx 绘图最小X坐标
- minY 绘图最小Y坐标
- maxX 绘图最大X坐标
- maxY 绘图最大Y坐标
- width 宽度
- height 高度
- drawToBk

**说明**：

使用此函数将在指定范围内加载图片。

**示例**：

```vbs
通过鼠标移动加载图片：
Function OnMouseMove(x , y, spx , spy, flags)
minX=spx:minY= spy:maxX=spy+ 2000:maxY =spy+1000:imgType = 1
imgName="C:\1.jpg"
width=200:height= 100 :drawToBk= 3
SSView.DrawImage pDc,imgType,imgName ,minX, minY,maxX,maxY,width ,height, drawToBk
End Function
Sub OnInitScript()
mode =0  ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### DrawLine 画线

**函数**：`DrawLine`

**功能**：画线。

**语法**：

```vbs
SSView.DrawLine(pDc, x0, y0, x1, y1, color, width, linestyle)
```

**参数**：

- pDc 窗口标识（使用时无需修改默认即可）
- x0 起点的横坐标（屏幕坐标）
- y0 起点的纵坐标（屏幕坐标）
- x1 终点的横坐标（屏幕坐标）
- y1 终点的纵坐标（屏幕坐标）
- color 线的颜色值（RGB值）
- width 线宽
- linestyle 线型

**说明**：

该函数为屏幕绘图函数，用于在屏幕上绘制点。

**示例**：

```vbs
Dim arX(10000), arY(10000), arX0(10000), arY0(10000), arXYCount
Sub OnInitScript()
mode =0  ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode ,title
End Sub
Function OnLButtonDown(x , y, spx , spy, flags)
arX(arXYCount)=spx :arY(arXYCount )= spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x , y, spx , spy, flags)
If arXYCount>0 Then
SSProcess.PasteBackgroundImage
If arXYCount =1 Then
SSView.DrawLine pDc , arX( 0), arY( 0),spx, spy, RGB(255,0, 255),300, linestyle
ElseIf arXYCount >1 Then
SSView.DrawLine pDc, arX( arXYCount-1), arY(arXYCount-1),spx, spy, RGB(255,0,255),300, linestyle
draws spx ,spy
End If
End If
End Function
Function draws (spx,spy)
For i =1 To arXYCount-1
SSView.DrawLine pDc , arX( i-1), arY(i-1), arX( i), arY(i), RGB(255,0, 255),300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### DrawObj 绘指定对象

**函数**：`DrawObj`

**功能**：绘指定对象。

**语法**：

```vbs
SSView.DrawObj(pDc, handle)
```

**参数**：

- pDc 窗口标识（使用时无需修改默认即可）
- handle 对象句柄

**说明**：

该函数为屏幕绘图函数，用于在屏幕上绘制指定对象。


#### DrawPoint 画点

**函数**：`DrawPoint`

**功能**：画点。

**语法**：

```vbs
SSView.DrawPoint(pDc, x, y, color, width)
```

**参数**：

- pDc 窗口标识（使用时无需修改默认即可）
- x 点的横坐标（屏幕坐标）
- y 点的纵坐标（屏幕坐标）
- color 线的颜色值（RGB值）
- width 点的大小

**说明**：

该函数为屏幕绘图函数，用于在屏幕上绘制点。

**示例**：

```vbs
Dim arX(10000), arY(10000), arX0(10000), arY0(10000), arXYCount
Sub OnInitScript()
mode =0 ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode ,title
End Sub
Function OnLButtonDown(x , y, spx , spy, flags)
arX(arXYCount)=spx :arY(arXYCount )= spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x , y, spx , spy, flags)
SSView.GetClientArea minX , minY, maxX, maxY
haomi =SSView.GetMMtoME
x1=minX+haomi *20:y1 =minY+ haomi*20
point1 =x1&","& y1
point2 =x1+haomi *20&","& minY+haomi*20
point3 =x1+haomi *20&","& minY+haomi*40
point4 =x1&","& minY+haomi*40
pointStrings =point1&","& point2&"," &point3& ","&point4&","& point1
SSView.DrawPoint pDc , x1+haomi *10, y1 +haomi* 30, RGB( 255,0,0), 300
SSView.DrawPolyline pDc , pointStrings, RGB(255, 0,255),300, linestyle
SSView.DrawPolygon pDc , pointStrings, RGB(255, 0,255),"10", linestyle, RGB(255,0,255), fillstyle
If arXYCount >0 Then
SSProcess.PasteBackgroundImage
If arXYCount =1 Then
SSView.DrawLine pDc , arX( 0), arY( 0),spx, spy, RGB(255,0, 255),300, linestyle
ElseIf arXYCount >1 Then
SSView.DrawLine pDc , arX( arXYCount-1), arY(arXYCount-1),spx, spy, RGB(255,0,255), 300, linestyle
draws spx ,spy
End If
End If
End Function
Function draws (spx,spy)
For i =1 To arXYCount- 1
SSView.DrawLine pDc , arX( i-1), arY(i-1), arX( i), arY(i), RGB(255,0, 255),300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### DrawPolygon 画多边形

**函数**：`DrawPolygon`

**功能**：画多边形。

**语法**：

```vbs
SSView.DrawPolygon(pDc, pointStrings, linecolor, linewidth, linestyle, fillcolor, fillstyle)
```

**参数**：

- pDc 窗口标识（使用时无需修改默认即可）
- pointStrings 点列坐标串（坐标为屏幕坐标，格式为：x1,y1,x2,y2,x3,y3）
- linecolor 线的颜色值（RGB值）
- linewidth 线的宽度值
- linestyle 线型（不知其缺省值）
- fillcolor 填充颜色值（RGB值）
- fillstyle 填充样式（不知其缺省值）

**说明**：

该函数为屏幕绘图函数，用于在屏幕上绘制多边形。

**示例**：

```vbs
该函数为屏幕绘图函数，用于在屏幕上绘制多边形。
Dim arX(10000), arY(10000), arX0(10000), arY0(10000), arXYCount
Sub OnInitScript()
mode =0  ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode ,title
End Sub
Function OnLButtonDown(x , y, spx , spy, flags)
arX(arXYCount)=spx :arY(arXYCount )= spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x , y, spx , spy, flags)
SSView.GetClientArea minX , minY, maxX, maxY
haomi =SSView.GetMMtoME
x1=minX+haomi *20:y1 =minY+ haomi*20
point1 =x1&","& y1
point2 =x1+haomi *20&","& minY+haomi*20
point3 =x1+haomi *20&","& minY+haomi*40
point4 =x1&","& minY+haomi*40
pointStrings =point1&","& point2&"," &point3& ","&point4&","& point1
SSView.DrawPoint pDc , x1+haomi *10, y1 +haomi* 30, RGB( 255,0,0), 300
SSView.DrawPolyline pDc , pointStrings, RGB(255, 0,255),300, linestyle
SSView.DrawPolygon pDc , pointStrings, RGB(255, 0,255),"10", linestyle, RGB(255,0,255), fillstyle
If arXYCount >0 Then
SSProcess.PasteBackgroundImage
If arXYCount =1 Then
SSView.DrawLine pDc , arX( 0), arY( 0),spx, spy, RGB(255,0, 255),300, linestyle
ElseIf arXYCount >1 Then
SSView.DrawLine pDc, arX(arXYCount-1), arY(arXYCount-1),spx, spy, RGB(255,0,255), 300, linestyle
draws spx ,spy
End If
End If
End Function
Function draws (spx,spy)
For i =1 To arXYCount- 1
SSView.DrawLine pDc , arX( i-1), arY(i-1), arX( i), arY(i), RGB(255,0, 255),300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### DrawPolyline 画多点线

**函数**：`DrawPolyline`

**功能**：画多点线。

**语法**：

```vbs
SSView.DrawPolyline(pDc, pointStrings, color, width, linestyle)
```

**参数**：

- pDc 窗口标识（使用时无需修改默认即可）
- pointStrings 点列坐标串（坐标为屏幕坐标，格式为：x1,y1,x2,y2,x3,y3…）
- color 线的颜色值（RGB值）
- width 线的宽度值
- linestyle 线型（不知其缺省值）

**说明**：

该函数为屏幕绘图函数，用于在屏幕上绘制多点线。

**示例**：

```vbs
Dim arX(10000), arY(10000), arX0(10000), arY0(10000), arXYCount
Sub OnInitScript()
mode =0  ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode ,title
End Sub
Function OnLButtonDown(x , y, spx , spy, flags)
arX(arXYCount)=spx :arY(arXYCount )= spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x , y, spx , spy, flags)
SSView.GetClientArea minX , minY, maxX, maxY
haomi = SSView.GetMMtoME
x1=minX+haomi *20:y1 =minY+ haomi*20
point1 =x1&","& y1
point2 =x1+haomi *20&","& minY+haomi*20
point3 =x1+haomi *20&","& minY+haomi*40
point4 =x1&","& minY+haomi*40
pointStrings =point1&","& point2&"," &point3& ","&point4&","& point1
SSView.DrawPoint pDc , x1+haomi *10, y1 +haomi* 30, RGB( 255,0,0), 300
SSView.DrawPolyline pDc , pointStrings, RGB(255, 0,255),300, linestyle
SSView.DrawPolygon pDc, pointStrings, RGB(255, 0,255), "10", linestyle, RGB(255,0,255), fillstyle
If arXYCount >0 Then
SSProcess.PasteBackgroundImage
If arXYCount =1 Then
SSView.DrawLine pDc , arX( 0), arY( 0),spx, spy, RGB(255,0, 255),300, linestyle
ElseIf arXYCount >1 Then
SSView.DrawLine pDc, arX(arXYCount-1), arY(arXYCount-1),spx, spy, RGB(255,0,255), 300, linestyle
draws spx ,spy
End If
End If
End Function
Function draws (spx,spy)
For i =1 To arXYCount- 1
SSView.DrawLine pDc , arX( i-1), arY(i-1), arX( i), arY(i), RGB(255,0, 255),300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### DrawRuntimeObjAttrInfo 动态属性标注

**函数**：`DrawRuntimeObjAttrInfo`

**功能**：动态属性标注。

**语法**：

```vbs
SSView.DrawRuntimeObjAttrInfo(pDC, hObjList, offsetX, offsetY, strTextFormat, strFontInfo, Color)
```

**参数**：

- pDc 屏幕指针
- hObjList 标注的对象列表
- offsetX 标注的注记X方向偏移距离
- offset 标注的注记Y方向偏移距离
- strTextFormat 标注对象的属性
- strFontInfo 标注的注记内容
- color 标注的RGB颜色值

**说明**：

该函数实现地物的动态属性标注功能。

**示例**：

```vbs
Sub OnInitScript()
mode =0 ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus  0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnPopButtonClick()
btn_id= "GxPointName_Name"
sdl_id= SSProcess.GetSDLRuntimeCmdID("管线专用辅助功能 \管线点号标注.vbs")
SSProcess.RegisterPopButtonID sdl_id,btn_id
check= SSProcess.IsPopButtonChecked(btn_id )
if check=0 Then
SSProcess.CheckPopButton btn_id,1
SSProcess.CreateScriptHandle btn_id,"管线专用辅助功能\管线点号标注 .vbs"
Else
SSProcess.CheckPopButton btn_id,0
SSProcess.CloseScriptHandle btn_id
End If
SSProcess.RefreshView
End Sub
Function OnDraw(pDc)
gxPoints= SSArray.Create("CGeObjList", "")
SSView.GetClientArea minX,minY,maxX ,maxY
rect= minX&"," &minY& ","&maxX&","& maxY
SSProcess.SetSelectConditionGroup"condition1", "SSObj_Code","INRANGE" ,"1600000,1700000"
SSProcess.SearchRectrect,0,"condition1" ,gxPoints
SSProcess.ClearSelectConditionGroup "condition1"
strTextInfo1 ="[物探点号]"
SSView.DrawRuntimeObjAttrInfo pDc,gxPoints,10, 10, strTextInfo1,"" ,-000255000
SSArray.ClosegxPoints
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### DrawText 画文本

**函数**：`DrawText`

**功能**：画文本。

**语法**：

```vbs
SSView.DrawText(pDc, x, y, text, alignment, font, color, width, height, stringangle, wordangle, italicangle, downangle, fontWeight, fontInterval, underLine, textBKMode)
```

**参数**：

- pDc 屏幕指针
- x 返回屏幕横坐标
- y 返回屏幕纵坐标
- text 注记内容
- alignment 对齐方式
- font 字体名
- color 注记颜色
- width 注记宽度
- height 注记高度
- stringangle 字串角度
- wordangle 字角度
- italicangle 是否斜体：0，否；1，是
- downangle 是否耸肩：0，否；1，是
- fontWeight 字重
- fontInterval 字隔
- underline 是否下划线：0，否；1，是
- textBKMode 返回实际的空间纵坐标

**说明**：

该函数为系统消息显示函数，用于无模脚本中，鼠标移动即触发该函数。

**示例**：

```vbs
text = "注记"
Function OnMouseMove(x , y,spx , spy, flags)
SSProcess.PasteBackgroundImage
If text <>"" Then
SSProcess.DrawText pDc,spx,spy ,text, 0," 黑体", RGB(255,255, 0),150,150, 0,0,0, 0,500,10, 0,2
End If
End Function
Sub OnInitScript()
mode =0 ‘=0 无参数对话框 =1 有参数对话框
title="功能名称"
SSProcess.ShowScriptDlg mode,title
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### GetClientArea 获取图形窗口实地范围

**函数**：`GetClientArea`

**功能**：获取图形窗口实地范围。

**语法**：

```vbs
SSView.GetClientArea(minX, minY, maxX, maxY)
```

**参数**：

- minX 图形窗口最小X坐标
- minY 图形窗口最小Y坐标
- maxX 图形窗口最大X坐标
- maxY 图形窗口最大Y坐标
- minX、minY、maxX、maxY为获取的窗口屏幕最小、最大坐标。

**说明**：

此函数功能为获取图形窗口的实地范围坐标。

**示例**：

```vbs
Sub OnClick()
SSProcess.GetClientArea minX, minY, maxX, maxY
Msgbox minX&","&minY&","&maxX&","&maxY
End Sub
```


#### GetClientRect 获取图形窗口屏幕范围

**函数**：`GetClientRect`

**功能**：获取图形窗口屏幕范围。

**语法**：

```vbs
SSView.GetClientRect(l, t, r, b)
```

**参数**：

- l （lift）窗口范围参数 :屏幕显示区域尺寸，左侧值，通常为0
- t （top）窗口范围参数 :屏幕显示区域尺寸，顶部值，通常为0
- r （right）窗口范围参数 :屏幕显示区域尺寸，右侧值相当于窗口宽度，随窗口变化
- b （bottom）窗口范围参数 :屏幕显示区域尺寸，底部值相当于窗口高度，随窗口变化
- 由于窗口屏幕范围是相对窗口屏幕左上角而言的，因此左上角坐标为（0，0）。

**说明**：

此函数功能为获取窗口屏幕范围的坐标。窗口屏幕坐标指定窗口屏幕的左上角和右下角，由于窗口屏幕坐标是相对窗口屏幕的左上角而言的，因此左上角坐标为（0，0）。

**示例**：

```vbs
Sub OnInitScript()
title="GetClientRect"
'清除对象框参数
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "点击<确定>获取图形窗口屏幕范围" ,"", 1, "",""
SSProcess.AddInputParameter "屏幕左上角坐标" ,"", 0, "", ""
SSProcess.AddInputParameter "屏幕右下角坐标" ,"", 0, "", ""
dlgWidth=220:dlgHeight=150:colCount=1:titleWidth=100:valueWidth=90 SSProcess.ShowScriptUserDefDlg title ,"GetClientRect" ,dlgWidth,dlgHeight,
colCount ,titleWidth, valueWidth
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
SSView.GetClientRect l, t, r , b  '获取图形窗口屏幕范围
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "点击<确定>获取图形窗口屏幕范围" ,"", 1, "",""
SSProcess.AddInputParameter "屏幕左上角坐标" ,l& ","&t, 0, "", ""
SSProcess.AddInputParameter "屏幕右下角坐标" ,r& ","&b, 0, "", ""
SSProcess.UpdateScriptDlgParameter 0
End Sub
Sub OnCancel()
End Sub
```


#### GetKeyState 获取按键状态

**函数**：`GetKeyState`

**功能**：获取按键状态。

**语法**：

```vbs
SSView.GetKeyState(virtKey)
```

**参数**：

- virtKey 键盘键值

**说明**：

获取键盘状态，需要将16进制通过计算机转换成10进制

**示例**：

```vbs
键盘F1（112）
Sub OnClick()
'获取按键状态
ss =SSView.GetKeyState( 112)
Msgbox ss
End Sub
```


#### GetMCtoCC 图纸上 0.1 毫米代表的屏幕上的像素

**函数**：`GetMCtoCC`

**功能**：写入外部数据之前。

**语法**：

```vbs
SSView.GetMCtoCC()
```

**示例**：

```vbs
Sub OnClick()
'图纸上0.1毫米代表的屏幕上的像素
cc =SSView.GetMCtoCC()
Msgbox cc
End Sub
```


#### GetMMtoME 获取当前屏幕每毫米对应实地米数

**函数**：`GetMMtoME`

**功能**：获取当前屏幕每毫米对应实地米数。

**语法**：

```vbs
SSView.GetMMtoME()
```

**示例**：

```vbs
Sub OnClick()
'获取当前屏幕每毫米对应实地米数
mm =SSView.GetMMtoME()
Msgbox mm
End Sub
```


#### MakeBackgroundImage 复制当前屏幕到剪切板

**函数**：`MakeBackgroundImage`

**功能**：复制当前屏幕到剪切板。

**语法**：

```vbs
SSView.MakeBackgroundImage()
```

**说明**：

该函数用于复制当前屏幕到剪粘板。


#### MoveScreenByPoint 移动屏幕到指定位置

**函数**：`MoveScreenByPoint`

**功能**：移动屏幕到指定位置。

**语法**：

```vbs
SSView.MoveScreenByPoint(x, y, redraw)
```

**参数**：

- x 设置需移动屏幕到指定位置的X坐标
- y 设置需移动屏幕到指定位置的Y坐标
- redraw 设置是否重绘屏幕

**说明**：

该函数用于移动屏幕到指定位置。

**示例**：

```vbs
Sub OnClick()
SSView.MoveScreenByPoint 55555, 5555, redraw
End Sub
```


#### PasteBackgroundImage 拷贝剪切板内容到当前屏幕

**函数**：`PasteBackgroundImage`

**功能**：拷贝剪切板内容到当前屏幕。

**语法**：

```vbs
SSView.PasteBackgroundImage()
```

**说明**：

该函数用于拷贝剪粘板内容到当前屏幕。


#### SetPenWidth 设置线笔宽

**函数**：`SetPenWidth`

**功能**：写入外部数据之前。

**语法**：

```vbs
SSView.SetPenWidth(dWidth)
```

**参数**：

- dWidth 设置屏幕绘制时的线笔宽度

**说明**：

该函数用于设置屏幕绘制时的线笔宽度。

**示例**：

```vbs
Dim arX(10000),arY(10000), arX0(10000),arY0(10000),arXYCount
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
SSProcess.SetCursorStatus 0
SSProcess.ClearSysSelection
SSProcess.ClearInputParameter
title="屏幕绘图"
SSProcess.ShowScriptDlg mode,title
End Sub
SSProject.SetActiveMap mapHandle
Function OnLButtonDown(x, y, spx, spy, flags)
arX(arXYCount) =spx:arY(arXYCount) = spy
arXYCount=arXYCount+1
End Function
Function OnMouseMove(x, y, spx, spy, flags)
SSView.MoveScreenByPoint spx,spy,1
SSView.GetClientArea minX, minY, maxX, maxY
haomi= SSView.GetMMtoME
x1=minX+haomi*20:y1=minY+haomi*20
point1=x1&","&y1
point2=x1+haomi*20&","&minY+haomi*20
point3=x1+haomi*20&","&minY+haomi*40
point4=x1&","&minY+haomi*40
pointStrings=point1&","&point2&","&point3&","&point4&","&point1
SSView.DrawPoint pDc, x1+haomi*10, y1+haomi*30, RGB(255,0,0), 300
SSView.DrawPolyline pDc, pointStrings, RGB(255,0,255), 300, linestyle
SSView.DrawPolygon pDc, pointStrings, RGB(255,0,255), "10", linestyle, RGB(255,0,255), fillstyle
If arXYCount>0 Then
SSProcess.PasteBackgroundImage
If arXYCount = 1 Then
SSView.DrawLine pDc, arX(0) , arY(0) ,spx,spy, RGB(255,0,255), 300, linestyle
ElseIf arXYCount > 1 Then
SSView.SetPenWidth 80
SSView.DrawLine Dc, arX(arXYCount-1), arY(arXYCount-1), px,spy, RGB(255,0,255), 300, linestyle
draws spx,spy
End If
End If
End Function
Function draws (spx,spy)
For i = 1 To arXYCount-1
SSView.DrawLine pDc, arX(i-1) , arY(i-1) , arX(i) , arY(i) , RGB(255,0,255), 300, linestyle
Next
End Function
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCancel()
End Sub
```


#### SetPointSize 设置点笔宽

**函数**：`SetPointSize`

**功能**：设置点笔宽。

**语法**：

```vbs
SSView.SetPointSize(dSize)
```

**参数**：

- dSize 点笔宽像素

**说明**：

使用此函数将指定绘制点的大小。

**示例**：

```vbs
设置点笔宽：
Sub OnClick()
SSView.SetPointSize 10
SSView.DrawPoint pDc, 10, 30, RGB(255,0,0), 300
End Sub
```


#### SetROP2 设置绘图异或方式

**函数**：`SetROP2`

**功能**：设置绘图异或方式。

**语法**：

```vbs
SSView.SetROP2(nLogicop)
```

**参数**：

- nLogicop

**说明**：

该函数用于设定当前绘图前景色的混合模式。

**示例**：

```vbs
scroll = 1
Sub OnInitScript()
mode = 0 ‘=0 无参数对话框 =1 有参数对话框
title="节目号抽奖程序"
SSProcess.ShowScriptDlg mode,title
scroll = 1
SSView.ZoomExtend 0, 0 , 50, 50 , 1
SSProcess.SetTimer 128801, 2
End Sub
Sub OnExitScript()
End Sub
Sub OnOK()
End Sub
Sub OnCanel()
SSProcess.KillTimer 128801
End Sub
Function OnLButtonUp(x , y, spx , spy, flags)
If scroll = 1 Then
scroll = 0
Else
scroll = 1
End If
End Function
Function Draw( pDc )
SSView.GetClientArea minX, minY, maxX , maxY
x = (minX + maxX)/2.0
y = (minY + maxY)/ 2.0
text = FormatNumber(times , 0)
nOldLogicop = SSView.SetROP2 ( 13)
xoff = 18
yoff = 12
pointStrings = x-xoff & "," & y-yoff & "," & x+xoff & "," & y-yoff & "," & x+xoff & "," & y+yoff & "," & x-xoff & "," & y+yoff & "," & x-xoff & "," & y -yoff
SSView.DrawPolyline 0, pointStrings, 11513775 , 100, 1
x = x + 0.25
y = y - 0.25
pointStrings = x-xoff & "," & y-yoff & "," & x+xoff & "," & y-yoff & "," & x+xoff & "," & y+yoff & "," & x-xoff & "," & y+yoff & "," & x-xoff & "," & y -yoff
SSView.DrawPolyline 0, pointStrings, 11511775 , 100, 1
x = x + 0.25
y = y - 0.25
pointStrings = x-xoff & "," & y-yoff & "," & x+xoff & "," & y-yoff & "," & x+xoff & "," & y+yoff & "," & x-xoff & "," & y+yoff & "," & x-xoff & "," & y -yoff
SSView.DrawPolygon 0, pointStrings, 16646143 , 1, 1 , 16777215, 0
x = x - 0.5
y = y + 0.5
SSView.DrawText 0, x, y , times, 0 , "黑体", 256000, 6000, 6000, 0 , 0, 0, 0 , 500, 0 , 0, 1
SSView.DrawText 0, x+ 0.5, y- 0.5, times, 0 , "黑体", 25500, 6000, 6000, 0, 0 , 0, 0 , 500, 0 , 0, 0
SSView.DrawText 0, x+ 1, y- 1, times, 0 , "黑体", 255, 6000, 6000, 0 , 0, 0, 0 , 500, 0 , 0, 0
SSView.SetROP2 nOldLogicop
End Function
Function OnDraw( pDc )
Draw pDC
End Function
times = 0
Function OnTimer( nIDEvent )
If scroll=1 Then
Draw 0
times = times+ 1
If  times > 55 Then
times = 1
End If
End Function
```


#### ZoomExtend 缩放屏幕到指定位置

**函数**：`ZoomExtend`

**功能**：缩放屏幕到指定范围。

**语法**：

```vbs
SSView.ZoomExtend(xMin,yMin,xMax ,yMax, redraw)
```

**参数**：

- xMin 最小X坐标
- yMin 最小Y坐标
- xMax 最大X坐标
- yMax 最大Y坐标
- redraw 返回参数

**说明**：

该函数设定屏幕显示范围。

**示例**：

```vbs
Sub OnClick()
xMin= 500
yMin= 500
xMax= 990
yMax= 690
SSView.ZoomExtend xMin,yMin,xMax ,yMax, redraw
SSProcess.RefreshView
End Sub
```


---

## SSArray 对象参考

### 数组函数

#### hArrayDest 追加数组元素 Append

**函数**：`Append`

**功能**：追加数组元素。

**语法**：

```vbs
SSArray.Append(hArrayDest, hArraySrc)
```

**参数**：

- hArrayDest
- hArraySrc


#### arrayType 删除数组 Close

**函数**：`Close`

**功能**：删除数组。

**语法**：

```vbs
SSArray.Close (arrayType)
```

**参数**：

- arrayType 数值类型

**示例**：

```vbs
Function OnDraw( pDc )
gxPoints = SSArray.Create( "CGeObjList", "")
SSView.GetClientArea minX , minY, maxX, maxY
rect = minX & ","& minY & ","& maxX & ","& maxY
strTextInfo = "<SSObj_PointName>"
SSProcess.SetSelectConditionGroup "condition1", "SSObj_ModifyTime", ">", g_SwitchTime1
SSProcess.SetSelectConditionGroup "condition1", "SSObj_ModifyTime", "<", g_SwitchTime2
SSProcess.SearchRect rect , 0, "condition1", gxPoints
SSView.DrawRuntimeObjAttrInfo pDc , gxPoints, 10, 0, strTextInfo, "", -000255000
SSProcess.SearchRect rect , 1, "condition1", gxPoints
SSView.DrawRuntimeObjAttrInfo pDc , gxPoints, 10, 0, strTextInfo, "", -000255000
SSProcess.SearchRect rect , 2, "condition1" , gxPoints
SSView.DrawRuntimeObjAttrInfo pDc , gxPoints, 10, 0, strTextInfo, "", -000255000
SSArray.Close gxPoints
End Function
```


#### hArrayDest 拷贝数组元素 Copy

**函数**：`Copy`

**功能**：拷贝数组元素。

**语法**：

```vbs
SSArray.Copy (hArrayDest, hArraySrc)
```

**参数**：

- hArrayDest
- hArraySrc


#### arrayType 创建数组 Create

**函数**：`Create`

**功能**：创建数组。

**语法**：

```vbs
SSArray.Create (arrayType, arrayName)
```

**参数**：

- arrayType 数值类型
- arrayName 数组名称


#### arrayType 根据类型与名称获取数组句柄 GetHandle

**函数**：`GetHandle`

**功能**：获取数组句柄。

**语法**：

```vbs
SSArray.GetHandle(arrayType, arrayName)
```

**参数**：

- arrayType 数组类型
- arrayName 数组名称


#### hArray 获取数组中元素数量 GetSize

**函数**：`GetSize`

**功能**：获取数组中元素数量。

**语法**：

```vbs
SSArray.GetSize(hArray)
```

**参数**：

- hArray


#### hArray 数组排序 Order

**函数**：`Order`

**功能**：数组排序。

**语法**：

```vbs
SSArray.Order(hArray, bIncrease)
```

**参数**：

- hArray
- bIncrease


#### hArray 数组元素反序 Reverse

**函数**：`Reverse`

**功能**：数组元素反序。

**语法**：

```vbs
SSArray.Reverse(hArray)
```

**参数**：

- hArray


#### hArray 设置数组中元素数量 SetSize

**函数**：`SetSize`

**功能**：设置数组中元素数量。

**语法**：

```vbs
SSArray.SetSize(hArray, nSize, nGrow)
```

**参数**：

- hArray 数组名
- nSize
- nGrow


#### hArrayValues 数组按值排序 SortByValue

**函数**：`SortByValue`

**功能**：数组按值排序。

**语法**：

```vbs
SSArray.SortByValue(hArrayValues, hArrayKeys, bIncrease)
```

**参数**：

- hArrayValues
- hArrayKeys
- bIncrease


---

## SSParmeter 对象参考

### 共享参数函数

#### GetParameterDATE 获取日期型参数

**函数**：`GetParameterDATE`

**功能**：获取日期型参数。

**参数**：

- strSection 组名
- strKeyName 约定的关键字
- defaultValut 缺省值
- keyValue 返回值

**说明**：

执行该函数，获取系统传回的所对应的特定值（日期型）。


#### GetParameterDBL 获取浮点型参数

**函数**：`GetParameterDBL`

**功能**：获取浮点型参数。

**参数**：

- strSection 组名
- strKeyName 约定的关键字
- defaultValut 缺省值
- keyValue 返回值

**说明**：

执行该函数，获取系统传回的所对应的特定值（浮点型）。


#### GetParameterINT 获取整型参数

**函数**：`GetParameterINT`

**功能**：获取整型参数。

**参数**：

- strSection 组名
- strKeyName 约定的关键字；"CurrentNoteID"为获取当前的注记对象ID，"CurrentObjID" 为获取当前的点、线、面对象ID； "ModifyTypeEnum"为获取修改类型
- defaultValut 缺省值
- keyValue 返回值

**说明**：

执行该函数，获取系统传回的所对应的特定值（整型）。

**示例**：

```vbs
Sub OnClick()
Dim arID(100),nCount
SSParameter.GetParameterINT "AfterQuickNote", "CurrentNoteID", 0, noteID
If noteID = 0 Then Exit Sub
'获取文字注记的注记分类号
objCode = SSProcess.GetObjectAttr (noteID, "SSObj_FontClass")
If objCode = "3990022" Then
'.....
End If
End Sub
```


#### GetParameterSTR 获取字符型参数

**函数**：`GetParameterSTR`

**功能**：获取字符型参数。

**参数**：

- strSection 组名
- strKeyName 约定的关键字
- defaultValut 缺省值
- keyValue 返回值

**说明**：

执行该函数，获取系统传回的所对应的特定值（字符型）。

**示例**：

```vbs
Sub OnClick()
SSParameter.GetParameterSTR "AfterSelectionItemChange", "TagetObjIDs", "", tagetObjIDs
SSParameter.GetParameterINT "AfterSelectionItemChange", "ModifyTypeEnum", -1, modifyTypeEnum
SSParameter.GetParameterSTR "AfterSelectionItemChange", "ModifyGroup", "", modifyGroup
SSParameter.GetParameterSTR "AfterSelectionItemChange", "ModifyItemName", "", modifyItemName
Dim tagetObjIDList(100000), tagetCount
If tagetObjIDs<>"" Then
If modifyTypeEnum=55 Then '扩展属性修改
ScanString tagetObjIDs, ",", tagetObjIDList, tagetCount
For i=0 To tagetCount-1
geoID = tagetObjIDList(i)
If modifyItemName="房屋层数" Then
floorCount = SSProcess.GetObjectAttr (geoID, "[房屋层数]" )
objArea = SSProcess.GetObjectAttr (geoID, "SSObj_Area" ): objArea =formatnumber (objArea,3)
SSProcess.SetObjectAttr geoID, "[建筑面积]" , CDbl(floorCount) * Cdbl(objArea)
SSProcess.UpdateSysSelection 0 '更新选择集
SSProcess.UpdateSysSelection 1
End If
Next
End If
End If
End Sub
```


#### GetParameterVOID 获取指针型参数

**函数**：`GetParameterVOID`

**功能**：获取指针型参数。

**参数**：

- strSection 组名
- strKeyName 约定的关键字
- keyValue 返回值

**说明**：

执行该函数，获取系统传回的所对应的特定值（指针型）。


#### RemoveParameter 删除参数

**函数**：`RemoveParameter`

**功能**：删除参数。

**参数**：

- strSection 组名
- strKeyName 约定的关键字

**说明**：

执行该函数，删除指定组指定关键字位置的参数值。


#### SetParameterDATE 设置日期型参数

**函数**：`SetParameterDATE`

**功能**：设置日期型参数。

**参数**：

- strSection 参数组名
- strKeyName 参数名称
- keyValue 参数值

**说明**：

执行该函数，获取系统传回的所对应的特定值（日期型）。

**示例**：

```vbs
Sub OnClick()
SSParameter.SetParameterDATE "a", "1" , "2013-01-01"
SSParameter.GetParameterDATE "a","1" ,"1990-01-01", mDate
Msgbox mDate
End Sub
```


#### SetParameterDBL 设置浮点型参数

**函数**：`SetParameterDBL`

**功能**：设置浮点型参数。

**参数**：

- strSection 参数组名
- strKeyName 参数名称
- keyValue 参数值

**说明**：

使用此函数可以使参数的类型设置为浮点型，可以存储浮点型的值。

**示例**：

```vbs
Sub OnClick()
SSParameter.GetParameterSTR "PrintInfo", "PaperName", "", PaperName
'纸张长宽
SSParameter.GetParameterINT "PrintInfo", "SIZEX", 0, PapersizeX
SSParameter.GetParameterINT "PrintInfo", "SIZEY", 0, PapersizeY
'页边距
SSParameter.GetParameterINT "PrintInfo", "EdgeRectLeft", 0, EdgeRectleft
SSParameter.GetParameterINT "PrintInfo", "EdgeRectRight", 0, EdgeRectright
SSParameter.GetParameterINT "PrintInfo", "EdgeRectTop", 0 , EdgeRecttop
SSParameter.GetParameterINT "PrintInfo", "EdgeRectBottom", 0 , EdgeRectbottom
'基于图廓的白边
SSParameter.GetParameterINT "PrintInfo", "DataIn", 0, m_DataIn
maxX= SSProject.GetMapInfo(SSProject.GetActiveMap, "maxX")
maxY= SSProject.GetMapInfo(SSProject.GetActiveMap, "maxY")
'计算出定位的将要生成的图形的位置X、y
SSParameter.SetParameterDBL "Location", "PositionX", maxX+50
SSParameter.SetParameterDBL "Location", "PositionY", maxY
End Sub
```


#### SetParameterINT 设置整型参数

**函数**：`SetParameterINT`

**功能**：设置整型参数。

**参数**：

- strSection 组名
- strKeyName 分组名
- keyValue 设置的值

**说明**：

执行该函数，设置整型参数。

**示例**：

```vbs
Sub OnClick()
Dim valueList(100),fieldCount
edbname= SSProcess.GetProjectFileName
SSProcess.OpenAccessMdbedbname
sql= "Select ID,[实际层数],[ 楼名] From [房屋面属性表] WHERE [实际层数] <>'*' Order By CDBL([实际层数 ])"
SSProcess.OpenAccessRecordsetedbname,sql
count= SSProcess.GetAccessRecordCount(edbname ,sql)
'写返回记录数
SSParameter.SetParameterINT "FeatureListQuery","RecordCount", count
'写返回字段数
SSParameter.SetParameterINT "FeatureListQuery","FieldCount", 3
SSProcess.CloseAccessRecordsetedbname,sql
SSProcess.CloseAccessMdbedbname
End Sub
```


#### SetParameterSTR 设置字符型参数

**函数**：`SetParameterSTR`

**功能**：设置字符型参数。

**参数**：

- strSection 组名
- strKeyName 分组名
- keyValue 设置的值

**说明**：

执行该函数，设置字符型参数。

**示例**：

```vbs
Sub OnClick()
SSParameter.SetParameterSTR "南宁国土", "获取图幅号字符串 ",  ""
SSParameter.GetParameterSTR "南宁国土", "获取图幅号地物 ID", "", keyValue1
End Sub
```


#### SetParameterVOID 设置指针型参数

**函数**：`SetParameterVOID`

**功能**：设置指针型参数。

**参数**：

- strSection 组名
- strKeyName 分组名
- keyValue 设置的值

**说明**：

执行该函数，设置指针型参数。

**示例**：

```vbs
Sub OnClick()
SSParameter.SetParameterVOID "a", "1", 1234
SSParameter.GetParameterVOID "a","1",vv
Msgbox vv
End Sub
```


---

## SSFunc 对象参考

### 扩展函数库

#### atof 字符转双精度型，自动处理非法数据

**函数**：`atof`

**功能**：字符转双精度型。

**语法**：

```vbs
SSFunc.atof (str)
```

**参数**：

- str 需转换的字符

**说明**：

该函数用于字符转双精度类型，自动处理非法数据。

**示例**：

```vbs
Sub OnClick()
str= “10.1”
str_dbl=SSFunc.atof(str)
Msgbox str+ str
Msgbox str_dbl+str_dbl
End Sub
```


#### atol 字符转整型，自动处理非法数据

**函数**：`atol`

**功能**：字符转整型，自动处理非法数据。

**语法**：

```vbs
SSFunc.atol (str)
```

**参数**：

- str 字符串

**说明**：

该函数将字符型转整型，自动处理非法数据。

**示例**：

```vbs
Sub OnClick()
strEmpty = "NULL"
iEmpty = SSFunc.atol(strEmpty ) '字符转整型
Msgbox "strEmpty =" &strEmpty&chr(13)&"iEmpty =" &iEmpty ,64,"消息"
End Sub
```


#### AutoTrimExtend 自动延伸裁剪

**函数**：`AutoTrimExtend`

**功能**：自动延伸裁剪。

**语法**：

```vbs
SSFunc.AutoTrimExtend(strTargetCode, strBorderCode, strParameters)
```

**参数**：

- strTargetCode 目标地物编码，为空时使用系统选择集
- strBorderCode 边界线编码，支持SSObj_Type=AREA && [扩展属性]=A 等标准语法
- strParameters 边界类型(0或1或2),是否在边界线上插点,限距,是否显示进度条

**说明**：

执行该函数，自动对地物进行延伸裁减。

**示例**：

```vbs
'道路中心线悬挂处理
Sub OnClick()
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition"SSObj_Code", "=","4201012,4201022,4202012,4202022,4203012,4203022,4204002,4205002,4206002,4209012,4209022,4302002,4303002,4304002,4305012,4305022,4305032,4306002"
SSProcess.SelectFilter
SSProcess.UpdateSysSelection 1 '脚本选择集选择的内容更新到系统选择集中
'返回值：处理的地物个数
nDealCount= SSFunc.AutoTrimExtend("", strRoadCentreLineCode,"1,1,8,1")
'清空选择
SSProcess.ExecuteSDLFunction"$SDL.SSEdit.SelectionEmpty",0
'图形重新生成
SSProcess.ExecuteSDLFunction"$SDL.SSProject.Display.RedrawExtend",0
Msgbox "道路悬挂点处理完成！"
End Sub
```


#### CloseLogFile 关闭 LOG 文件

**函数**：`CloseLogFile`

**功能**：关闭LOG文件。

**语法**：

```vbs
SSFunc.CloseLogFile(strFile, strPostfix, nShowLogFile)
```

**参数**：

- strFile 设置LOG文件的路径
- strPostfix 设置LOG文件的后缀名
- nShowLogFile 设置是否显示LOG文件，取值为0和1（0为不显示；1为显示）

**说明**：

执行该函数，将关闭LOG日志文件。

**示例**：

```vbs
Sub OnClick()
loglines="地名属性信息表（EXCEL）导入 "&"^" &chr(13)&chr( 10)&"^" &"**************** *****" _ +"^" &"地名属性 _边贸区口岸_边贸区口岸点信息表 .xlsx"&"^" &"**************" _+"^" &chr(13)&"^" &"------------错误记录---------------" &"^"& "错误内容···"_ +"^" &"-----------"
Dim lineArray(10000)
edbName =SSProcess.GetProjectFileName
SSFunc.OpenLogFile GetParent (edbName)& "error.txt", ".log"
SSFunc.ScanString loglines , "^", lineArray, lineCount
For i = 0 To lineCount-1
SSFunc.WriteLogFile lineArray (i)
Next
SSFunc.CloseLogFile GetParent (edbName)& "error.txt", ".log", 1
End Sub
'-------------------------- 获取文件的父目录 -----------------------------------
Function GetParent(strFile )
Dim fileArr(20),PfileArr()
SSFunc.ScanString strFile , "\", fileArr, fileCount
Redim PfileArr(fileCount-1)
For i = 0 To fileCount-2
PfileArr (i)=fileArr (i)
Next
GetParent=Join(PfileArr ,"\" )
End Function
```


#### CreateLegend 根据选择集生成图例

**函数**：`CreateLegend`

**功能**：根据选择集生成图例。

**语法**：

```vbs
SSFunc.CreateLegend(strSchemeName, dbX, dbY)
```

**参数**：

- strSchemeName 方案名称
- dbX 生成图例位置的X坐标值，对应 EPS中的 Y值
- dbY 生成图例位置的Y坐标值，对应 EPS中的 X值

**说明**：

使用此函数将选择集内包含的地物在指定位置生成图例。

**示例**：

```vbs
选择集内指定位置生成图例。
Sub OnClick()
SSFunc.CreateLegend "生成图例", 182 , 705
End Sub
```


#### CrossLines 折线与折现集求交

**函数**：`CrossLines`

**功能**：折线与折线集求交。

**语法**：

```vbs
SSFunc.CrossLines(strIdOrXY, strNearIds, xArray, yArray, idArray)
```


#### DealObjectHander 地物指针列表操作

**函数**：`DealObjectHander`

**功能**：地物指针列表操作。

**语法**：

```vbs
SSFunc.DealObjectHander(vArrayObj, nCount, strMethod, strParameters)
```


#### DeleteIsLand 岛面删除

**函数**：`DeleteIsLand`

**功能**：岛面删除。

**语法**：

```vbs
SSFunc.DeleteIsLand(hObj, bIsObjHander, strParameters)
```

**参数**：

- hObj 地物ID
- bIsObjHander
- strParameters

**说明**：

执行该函数，对指定地物进行岛面的删除

**示例**：

```vbs
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "AREA"
SSProcess.SelectFilter
objcount=SSProcess.GetSelGeoCount '获取选择集地物个数
If objcount=0 Then Exit Sub
For i=0 To objcount- 1
hObj =SSProcess.GetSelGeoValue (i , "SSObj_ID") ' 获取地物ID
bIsObjHander = 0
strParameters = 0
SSFunc.DeleteIsLand hObj ,bIsObjHander, strParameters ' 删除面地物的岛面
Next
End Sub
```


#### ExcelClose EXCEL 操作，关闭 EXCEL 文件

**函数**：`ExcelClose`

**功能**：EXCEL操作，关闭EXCEL文件。

**语法**：

```vbs
SSFunc.ExcelClose(strFile, strSheet)
```

**参数**：

- strFile Excel文件名（含完整路径）
- strSheet 工作表名称

**说明**：

执行该函数，关闭打开的EXCLE工作表及文件。

**示例**：

```vbs
Sub OnClick()
'读EXCLE表格
Dim arFile(20),nRecordCount,nxCount
aa = SSFunc.ExcelSelectFile(arFile, nFileCount)
If aa = 0 Then Exit Sub
For nFile=0 To nFileCount-1
strFile = arFile(nFile) : strSheet = "合同信息" 'Sheet工作表名
SSFunc.ExcelOpen strFile, strSheet
'获取EXCLE表的总行数
nTotal = SSFunc.ExcelGetTotalRow
Redim aryKJSJ(nTotal)
Redim aryBSM(nTotal)
Redim aryr(29)
'获取每一行需要的数据
SSFunc.ExcelGetRecord "农户代码,地块名称,地类,等级,长,宽", "^", aryKJSJ, nRecordCount
'把每一行按列拆分成字符串
For i=1 nRecordCount-1
SSFunc.ScanString aryKJSJ(i), "^", aryr, nxCount
aryBSM(i) = trim(aryr(28))
Next
'关闭EXCLE表格
SSFunc.ExcelClose strFile, strSheet
Next
End Sub
```


#### ExcelGetRecord EXCEL 操作，获取 EXCEL 记录

**函数**：`ExcelGetRecord`

**功能**：EXCEL操作，获取EXCEL记录。

**语法**：

```vbs
SSFunc.ExcelGetRecord(strFields, strSeparator, vArray, nRecordCount)
```

**参数**：

- strFields 需要获取的数据的列名，列名之间以西文逗号分隔
- strSeparator 分隔符，获得的对应列名的记录值是一个由分隔符分隔的字符串；用什么分隔符来分隔由这个参数约定
- vArray 获取的每一行的列名的记录值所组成的字符串数据存放在这个数组
- nRecordCount 数组的实际数据个数

**说明**：

执行该函数，获取EXCLE工作表的相应列名所对应的每一行记录值，存放在一个字符串数组。

**示例**：

```vbs
Sub OnClick()
'读EXCLE表格
Dim arFile(20),nRecordCount,nxCount
aa = SSFunc.ExcelSelectFile(arFile, nFileCount)
If aa = 0 Then Exit Sub
For nFile=0 To nFileCount-1
strFile = arFile(nFile) : strSheet = "合同信息" 'Sheet工作表名
SSFunc.ExcelOpen strFile, strSheet
'获取EXCLE表的总行数
nTotal = SSFunc.ExcelGetTotalRow
redim aryKJSJ(nTotal)
redim aryBSM(nTotal)
Redim aryr(29)
'获取每一行需要的数据
SSFunc.ExcelGetRecord "农户代码,地块名称,地类,等级,长,宽", "^", aryKJSJ, nRecordCount
'把每一行按列拆分成字符串
For i=1 To nRecordCount-1
SSFunc.ScanString aryKJSJ(i), "^", aryr, nxCount
aryBSM(i) = trim(aryr(28))
Next
'关闭EXCLE表格
SSFunc.ExcelClose strFile, strSheet
Next
End Sub
```


#### ExcelGetTotalRow EXCEL 操作，获取 EXCEL 总行数

**函数**：`ExcelGetTotalRow`

**功能**：EXCEL操作，获取EXCEL总行数。

**语法**：

```vbs
SSFunc.ExcelGetTotalRow()
```

**参数**：

- 返回值 EXCLE工作表的总行数

**说明**：

执行该函数，获取EXCLE表的总行数函数。

**示例**：

```vbs
Sub OnClick()
'读EXCLE表格
Dim arFile(20),nRecordCount,nxCount
aa = SSFunc.ExcelSelectFile(arFile, nFileCount)
If aa = 0 Then Exit Sub
For nFile=0 To nFileCount-1
strFile = arFile(nFile) : strSheet = "合同信息" 'Sheet工作表名
SSFunc.ExcelOpen strFile, strSheet
'获取EXCLE表的总行数
nTotal = SSFunc.ExcelGetTotalRow
Redim aryKJSJ(nTotal)
Redim aryBSM(nTotal)
Redim aryr(29)
'获取每一行需要的数据
SSFunc.ExcelGetRecord "农户代码,地块名称,地类,等级,长,宽", "^", aryKJSJ, nRecordCount
'把每一行按列拆分成字符串
For i=1 To nRecordCount-1
SSFunc.ScanString aryKJSJ(i), "^", aryr, nxCount
aryBSM(i) = trim(aryr(28))
Next
'关闭EXCLE表格
SSFunc.ExcelClose strFile, strSheet
Next
End Sub
```


#### ExcelOpen EXCEL 操作，打开 EXCEL 文件

**函数**：`ExcelOpen`

**功能**：EXCEL操作，打开EXCEL文件。

**语法**：

```vbs
SSFunc.ExcelOpen(strFile, strSheet)
```

**参数**：

- strFile Excel文件名（含完整路径）
- strSheet 工作表名称

**说明**：

执行该函数，打开指定的Excel文件中的工作表。

**示例**：

```vbs
Sub OnClick()
'读EXCLE表格
Dim arFile(20),nRecordCount,nxCount
aa = SSFunc.ExcelSelectFile(arFile, nFileCount)
If aa = 0 Then Exit Sub
For nFile=0 To nFileCount-1
strFile = arFile(nFile) : strSheet = "合同信息" 'Sheet工作表名
SSFunc.ExcelOpen strFile, strSheet
'获取EXCLE表的总行数
nTotal = SSFunc.ExcelGetTotalRow
Redim aryKJSJ(nTotal)
Redim aryBSM(nTotal)
Redim aryr(29)
'获取每一行需要的数据
SSFunc.ExcelGetRecord "农户代码,地块名称,地类,等级,长,宽", "^", aryKJSJ, nRecordCount
'把每一行按列拆分成字符串
For i=1 To nRecordCount-1
SSFunc.ScanString aryKJSJ(i), "^", aryr, nxCount
aryBSM(i) = trim(aryr(28))
Next
'关闭EXCLE表格
SSFunc.ExcelClose strFile, strSheet
Next
End Sub
```


#### ExcelSelectFile EXCEL 操作，选择 EXCEL 文件

**函数**：`ExcelSelectFile`

**功能**：EXCEL操作，选择EXCEL文件。

**语法**：

```vbs
SSFunc.ExcelSelectFile(strFile)
```

**参数**：

- arFile 文件数组，记录用户选择的多个Excel文件
- nFileCount 文件数，记录用户选择的文件数目
- 返回值 0 没有选取任何文件；1 选取了文件

**说明**：

执行该函数，可弹出“选择Excel文件”对话框，供用户选择Excel文件；选择文件名存放在arFile数组中, nFileCount记录文件数目。

**示例**：

```vbs
Sub OnClick()
'读EXCLE表格
Dim arFile(20),nRecordCount,nxCount
aa = SSFunc.ExcelSelectFile(arFile, nFileCount)
If aa = 0 Then Exit Sub
For nFile=0 To nFileCount-1
strFile = arFile(nFile) : strSheet = "合同信息" 'Sheet工作表名
SSFunc.ExcelOpen strFile, strSheet
'获取EXCLE表的总行数
nTotal = SSFunc.ExcelGetTotalRow
Redim aryKJSJ(nTotal)
Redim aryBSM(nTotal)
Redim aryr(29)
'获取每一行需要的数据
SSFunc.ExcelGetRecord "农户代码,地块名称,地类,等级,长,宽", "^", aryKJSJ, nRecordCount
'把每一行按列拆分成字符串
For i=1 To nRecordCount-1
SSFunc.ScanString aryKJSJ(i), "^", aryr, nxCount
aryBSM(i) = trim(aryr(28))
Next
'关闭EXCLE表格
SSFunc.ExcelClose strFile, strSheet
Next
End Sub
```


#### GetAverageWidth 点列中是否有指定标记的点

**函数**：`GetAverageWidth`

**功能**：点列中是否有指定标记的点。

**语法**：

```vbs
SSFunc.GetAverageWidth(hObj, bIsObjHander, dbReturnLength)
```

**参数**：

- hObj
- bIsObjHander
- dbReturnLength


#### GetBinaryBitValue 获取二进制位的值

**函数**：`SetBinaryBitValue`

**功能**：设置二进制位的值。

**语法**：

```vbs
SSFunc.SetBinaryBitValue(nData, nBit)
```


#### GetChineseDigit 获取汉字数字

**函数**：`GetChineseDigit`

**功能**：获取汉字数字。

**语法**：

```vbs
SSFunc.GetChineseDigit(nDigit)
```

**参数**：

- nDigit 要获取汉字的数字 (取值范围为0-999)

**说明**：

使用此函数获取数字的汉字。

**示例**：

```vbs
使用此函数获取数字的汉字，比如获取 100的汉字一百。
Sub OnClick()
nDigit = 100
A = SSFunc.GetChineseDigit(nDigit )
Msgbox A
End Sub
```


#### GetCodeByObjectName 通过 ObjectName 获取编码

**函数**：`GetCodeByObjectName`

**功能**：通过ObjectName获取编码。

**语法**：

```vbs
SSFunc.GetCodeByObjectName(strObjectName, strLayerName)
```

**参数**：

- strObjectName 对象名称
- strLayerName 对象层名

**说明**：

此函数功能为通过ObjectName获取编码。

**示例**：

```vbs
获取ObjectName为“三角点”的地物个数及ID。
Sub OnClick()
Dim myIds(1000)
myObjName = " 三角点"
CheckObjByName myObjName ,myIds, myIdscnt
ids = ""
For i = 0 To myIdscnt - 1
If ids = "" Then
ids = myIds( i)
Else
ids = ids&","& myIds(i)
End If
Next
Msgbox "myIdscnt = "&myIdscnt&chr(13)& "ids = "&ids, 64," 消息"
End Sub
Function CheckObjByName(name ,byref ids(),byref idscnt)
code =SSFunc.GetCodeByObjectName( name, "")
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "=", code
SSProcess.SelectFilter
idscnt = SSProcess.GetSelGeoCount
For i=0 To idscnt-1
ids (i)= SSProcess.GetSelGeoValue ( i, "SSObj_ID" )
Next
End Function
```


#### GetDefaultAttrValue 获取扩展属性默认值

**函数**：`GetDefaultAttrValue`

**功能**：获取扩展属性默认值。

**语法**：

```vbs
SSFunc.GetDefaultAttrValue(nCode, vFields, nFieldCount, vDefaults, nDefaultCount)
```

**参数**：

- nCode 编码
- vFields 字段名
- nFieldCount 字段数
- vDefaults 默认值
- nDefaultCount 个数

**说明**：

此函数功能为获取扩展属性默认值。

**示例**：

```vbs
Sub OnOK()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "994000"
SSProcess.SelectFilter
Dim arField(1) : Dim arValue(15) : Dim arDefault(1) : arField (0) = "所属类型"
SSFunc.GetDefaultAttrValue 994000, arField, 1 , arDefault, nDefaultCount
strFields = "[面状属性],[第一树种],[棵数1],[第二树种],[棵数2],[第三树种],[棵数3],[胸径],[树高],[冠幅],[所属类型],[辅设材料],[所属街道]"
nUpper = SSProcess.GetSelGeoCount() - 1
For i=0 To nUpper
SSProcess.GetSelPolygonFocus i , x, y , z
str = SSProcess.GetSelGeoValue(i, strFields )
SSFunc.ScanString str , ",", arValue, nCount
if nDefaultCount = 1 Then
nIndex1 = instr( 1, arDefault( 0), arValue( 10) & ":" , 1)
if nIndex1 >= 1 Then
nIndex1 = nIndex1 + Len( arValue (10) & ":" )
nIndex2 = instr( nIndex1, arDefault (0), ",", 1)
if nIndex2 < 1 Then nIndex2 = Len ( arDefault( 0) ) + 1
arValue (10) = Mid(arDefault (0), nIndex1 , nIndex2-nIndex1)
End If
End If
CreateAreaAttrNote x , y, strFields , arValue
Next
SSProcess.SetMapStatus 0, 2
SSProcess.SaveBufferObjToDatabase
SSProcess.RefreshView
End Sub
```


#### GetNearstDistSquare 获取点到折线的最近距离平方

**函数**：`GetNearstDistSquare`

**功能**：获取点到折线的最近距离平方。

**语法**：

```vbs
SSFunc.GetNearstDistSquare(spx, spy, xPoints, yPoints, nCount, nNodeIndex, nSideIndex, nDirFlag, spxNearst, spyNearst)
```


#### GetObjectRect 获取地物列表的外接矩形

**函数**：`GetObjectRect`

**功能**：获取地物列表的外接矩形。

**语法**：

```vbs
SSFunc.GetObjectRect(vArrayObj, nCount, bIsObjHander, vArrayRect)
```

**参数**：

- vArrayObj 对象ID
- nCount 对象个数
- bIsObjHander 对象句柄
- vArrayRect 存储外接矩形字符串

**说明**：

此函数的功能为获取地物列表的外接矩形。

**示例**：

```vbs
Sub OnClick()
Dim idsss(100000),zbsss(4)
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT,LINE,AREA,NOTE"
SSProcess.SelectFilter
counttt=SSProcess.GetSelGeoCount
For i=0 To counttt- 1
idsss(i)=SSProcess.GetSelGeoValue (i, "SSObj_ID")
Next
If counttt=0 Then Exit Sub
SSFunc.GetObjectRect idsss , counttt, 0 , zbsss
x = ( cdbl(zbsss( 0))+cdbl (zbsss(2)))/2
y = ( cdbl(zbsss( 1))+cdbl (zbsss(3))) / 2
mapnumber = SSProcess.GetMapFrameNumber (x, y )
Msgbox mapnumber
End Sub
```


#### GetStripNumber 获取带号

**函数**：`GetStripNumber`

**功能**：获取带号。

**语法**：

```vbs
SSFunc.GetStripNumber()
```

**说明**：

执行该函数，将返回当前图幅带号。

**示例**：

```vbs
先对当前图幅代号进行设定为40，再对其进行获取。
Sub OnClick()
nStripNumber =40
SSFunc.SetStripNumber nStripNumber
a =SSFunc.GetStripNumber
Msgbox a
End Sub
```


#### HasPointMark 点列中是否有指定标记的点

**函数**：`HasPointMark`

**功能**：点列中是否有指定标记的点。

**语法**：

```vbs
SSFunc.HasPointMark(hObj, nBitValue, bIsObjHander)
```

**参数**：

- hObj
- nBitValue
- bIsObjHander


#### LineCrack 多义线（折线）分解

**函数**：`LineCrack`

**功能**：多义线（折线）分解。

**语法**：

```vbs
SSFunc.LineCrack()
```


#### OffsetObjList 平移地物

**函数**：`OffsetObjList`

**功能**：平移地物。

**语法**：

```vbs
SSFunc.OffsetObjList(vArrayObj, nCount, bIsObjHander, bSaveToMap, x, y, z)
```

**参数**：

- vArrayObj
- nCount
- bIsObjHander
- bSaveToMap
- x
- y
- z

**说明**：

执行此函数平移地物。


#### OpenLogFile 打开 LOG 文件

**函数**：`OpenLogFile`

**功能**：打开LOG文件。

**语法**：

```vbs
SSFunc.OpenLogFile(strFile, strPostfix)
```

**参数**：

- strFile 设置LOG文件的路径
- strPostfix 设置LOG文件的后缀名

**说明**：

执行该函数，将创建并打开LOG日志文件。

**示例**：

```vbs
Sub OnClick()
loglines="地名属性信息表（EXCEL）导入 "&"^" &chr(13)&chr( 10)&"^" &"********"_ +"^" &"地名属性 _边贸区口岸_边贸区口岸点信息表 .xlsx"&"^" &"********"_+"^" &chr(13)&"^" &"-----错误记录-----------" &"^"& "错误内容···"_ +"^" &"-------------"
Dim lineArray(10000)
edbName =SSProcess.GetProjectFileName
SSFunc.OpenLogFile GetParent (edbName)& "error.txt", ".log"
SSFunc.ScanString loglines , "^", lineArray, lineCount
For i = 0 To lineCount-1
SSFunc.WriteLogFile lineArray (i)
Next
SSFunc.CloseLogFile GetParent (edbName)& "error.txt", ".log", 1
End Sub
'-------------------------- 获取文件的父目录 -----------------------------------
Function GetParent(strFile )
Dim fileArr(20),PfileArr()
SSFunc.ScanString strFile , "\", fileArr, fileCount
Redim PfileArr(fileCount-1)
For i = 0 To fileCount-2
PfileArr (i)=fileArr (i)
Next
GetParent=Join(PfileArr ,"\" )
End Function
```


#### OutputGraphToBmp 输出指定范围内的图形到 BMP 图片

**函数**：`OutputGraphToBmp`

**功能**：输出指定范围内的图形到BMP图片。

**语法**：

```vbs
SSFunc.OutputGraphToBmp(minX, minY, maxX, maxY, nScale, strBmpFile, dpi)
```

**参数**：

- minX 纵坐标最小值
- minY 横坐标最小值
- maxX 纵坐标最大值
- maxY 横坐标最大值
- nScale 比例尺
- strBmpFile 输出路径
- dpi 输出像素

**说明**：

执行该函数，将把指定范围内创建的地物以bmp格式输出。

**示例**：

```vbs
Sub OnClick()
SSProcess.CreateNewObj 2
SSProcess.SetNewObjValue "SSObj_Code","2"
SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
SSProcess.AddNewObjPoint 0.78,0.15,  0, 0, ""
SSProcess.AddNewObjPoint 6.78,0.15, 0, 0, ""
SSProcess.AddNewObjPoint 6.78,3.15, 0, 0, ""
SSProcess.AddNewObjPoint 0.78,3.15, 0, 0, ""
SSProcess.AddNewObjPoint 0.78,0.15,  0, 0, ""
SSProcess.AddNewObjToSaveObjList
SSProcess.SaveBufferObjToDatabase
minX=0:minY=0:maxX=7.5:maxY=3.5
strBmpFile="C:\1&%^.bmp"
SSFunc.OutputGraphToBmp minX, minY, maxX, maxY, 25,strBmpFile,300
End Sub
```


#### ScanString 分解字符串，返回分解后的数组及数组元素个数

**函数**：`ScanString`

**功能**：分解字符串，返回分解后的数组及数组元素个数。

**语法**：

```vbs
SSFunc.ScanString(str, strSeparator, vArray, nCount)
```

**参数**：

- str 要分解的字符串
- strSeparator 分解符
- vArray 分解后的数组
- nCount 分解后的个数

**说明**：

使用此函数将指定的字符串进行分解。

**示例**：

```vbs
按照指定的分解符对字符串进行分解，并返回数组及个数。
Sub Onclick()
Dim arID (1000),idCount
ids="1;a;ab;1fd"
SSFunc.ScanString ids, ";", arID, idCount
End Sub
```


#### SelectGradeAttr 选择分级属性

**函数**：`SelectGradeAttr`

**功能**：选择分级属性。

**语法**：

```vbs
SSFunc.SelectGradeAttr(strTable, strFields, strTitle)
```

**参数**：

- strTable 属性表名称
- strFields 字段名称
- strTitle 视图标题

**说明**：

执行该函数，把基于strFields字段的属性值生成树视图。
返回值: 1 选择了某项或几项；0没有选择。

**示例**：

```vbs
'获取基于“NHDM”字段的属性分级视图
Sub OnClick()
b = SSFunc.SelectGradeAttr ("RLCM_承包经营权地块面属性表", "NHDM", "选择待输出权属人"
If b=0 Then Exit Sub
End Sub
```


#### SetBinaryBitValue 设置二进制位的值

**函数**：`SetBinaryBitValue`

**功能**：设置二进制位的值。

**语法**：

```vbs
SSFunc.SetBinaryBitValue(nData, nBit)
```


#### SetClipboardData 设置文本数据到剪贴板

**函数**：`SetClipboardData`

**功能**：设置文本数据到剪切板。

**语法**：

```vbs
SSFunc.SetClipboardData(strData)
```


#### SetStripNumber 设置带号

**函数**：`SetStripNumber`

**功能**：设置带号。

**语法**：

```vbs
SSFunc.SetStripNumber(nStripNumber)
```

**参数**：

- nStripNumber 带号

**说明**：

执行该函数，将按指定带号，重置当前工程的投影带号。

**示例**：

```vbs
Sub OnClick()
nStripNumber =38
SSFunc.SetStripNumber nStripNumber
End Sub
执行前效果：39带
执行后效果：38带
```


#### SortArray 排序

**函数**：`SortArray`

**功能**：排序。

**语法**：

```vbs
SSFunc.SortArray(vArray, nCount, nType)
```

**参数**：

- vArray 数组名称
- nCount 个数
- nType 排序类型nType=0时按数字与非数字组合排序，例如结果如：1,2,11,a2,a11,aa1,b3,b12nType=1时按ascii码值进行排序

**说明**：

使用此函数将对指定数组进行排序。

**示例**：

```vbs
将对数组进行排序。
Sub OnClick()
Dim vArray(3)
nCount=3
vArray(0)="b"
vArray(1)="a"
vArray(2)=132
nType=0
SSFunc.SortArray vArray, nCount, nType
End Sub
```


#### SortArrayByValue 对数组按值排序

**函数**：`SortArrayByValue`

**功能**：对数组按值排序。

**语法**：

```vbs
SSFunc.SortArrayByValue(arValue, arKey, nCount, nAscending, nType)
```

**参数**：

- arValue 排序参数原序号（ID）数组
- arKey 参与排序的关键字数组
- nCount 参与排序的数值个数
- nAscending 排序方式参数，1升序，0降序
- nType 排序数组字符类型（0带数字的字符/1原字符/2整数/3浮点数）

**说明**：

执行该函数可以对字符数组进行排序，并且把对应的ID也进行排序。
从选择集中拿到某个字段，对该字段的值进行排序，该值对应的ID也可以跟着进行排序。
Dim FWID(100), FWCs(100)
Sub OnClick()
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "==", "3103013"
SSProcess.SetSelectCondition "[story], 0
SSProcess.SelectFilter
FWCount = SSProcess.GetSelGeoCount()
Msgbox FWCount
For k = 0 To FWCount -1
FWID(k) = SSProcess.GetSelGeoValue (k, "SSObj_ID" )
FWCs(k) = SSProcess.GetSelGeoValue (k, "[story]" )
Next
'排序函数 排序参数原序号（ID）数组，排序参数数组，个数，1升序/0降序，排序字符类型（0带数字的字符/1原字符/2整数/3浮点数）
SSFunc.SortArrayByValue FWID,FWCs,4,1,0 '地块编码从小到大进行排序
End Sub


#### SortByBranch 对地物按分支排序

**函数**：`SortByBranch`

**功能**：对地物按分支排序。

**语法**：

```vbs
SSFunc.SortByBranch(vArrayObj, nCount, bIsObjHander, strParameter)
```


#### TransHZ2PY 汉字转拼音

**函数**：`TransHZ2PY`

**功能**：汉字转拼音。

**语法**：

```vbs
SSFunc.TransHZ2PY(strHZ, nMakeUpper)
```

**参数**：

- strHZ 输入的汉字
- nMakeUpper

**说明**：

使用此函数将汉字转换成拼音。

**示例**：

```vbs
将注记内容转换成拼音并储存至注记的扩展属性表备注字段中。
Sub OnClick()
'对注记类对象进行过滤
SSProcess.PushUndoMark
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "NOTE"
SSProcess.SelectFilter
'循环各个注记对象取其注记内容
NoteCount = SSProcess.GetSelNoteCount
IF NoteCount >0 Then
For i =0 To NoteCount -1
ID = SSProcess.GetSelNoteValue(i, "SSObj_ID")
fontString = SSProcess.GetSelNoteValue(i, "SSObj_FontString")
spellString = SSFunc.TransHZ2PY(fontString, nMakeUpper )
SSProcess.SetObjectAttr ID, "[Note]", spellString
Next
End If
Msgbox "输出完成！"
End Sub
```


#### ExcelSelectFile 写 LOG 文件

**函数**：`WriteLogFile`

**功能**：写LOG文件。

**语法**：

```vbs
SSFunc.WriteLogFile(strText)
```

**参数**：

- strText 需要写入LOG文件的字符串

**说明**：

执行该函数，将字符串写入LOG日志文件中。

**示例**：

```vbs
Sub OnClick()
loglines="地名属性信息表（EXCEL）导入 "&"^" &chr(13)&chr( 10)&"^" &"********"_ +"^" &"地名属性 _边贸区口岸_边贸区口岸点信息表 .xlsx"&"^" &"********"_+"^" &chr(13)&"^" &"-----错误记录-----------" &"^"& "错误内容···"_ +"^" &"-------------"
Dim lineArray(10000)
edbName =SSProcess.GetProjectFileName
SSFunc.OpenLogFile GetParent (edbName)& "error.txt", ".log"
SSFunc.ScanString loglines , "^", lineArray, lineCount
For i = 0 To lineCount-1
SSFunc.WriteLogFile lineArray (i)
Next
SSFunc.CloseLogFile GetParent (edbName)& "error.txt", ".log", 1
End Sub
'-------------------------- 获取文件的父目录 -----------------------------------
Function GetParent(strFile )
Dim fileArr(20),PfileArr()
SSFunc.ScanString strFile , "\", fileArr, fileCount
Redim PfileArr(fileCount-1)
For i = 0 To fileCount-2
PfileArr (i)=fileArr (i)
Next
GetParent=Join(PfileArr ,"\" )
End Function
```


---

## SSProject 对象参考

### 工程管理函数

#### CloseDataSource 关闭数据源

**函数**：`CloseDataSource`

**功能**：关闭数据源。

**语法**：

```vbs
SSProject.CloseDataSource(mapHandle, datasourceHandle)
```

**参数**：

- mapHandle 地图句柄
- datasourceHandle 数据源句柄

**说明**：

此函数功能为关闭数据源。

**示例**：

```vbs
Sub OnClick()
mapHandle = SSProject.GetActiveMap
datasourceHandle0 = SSProject.GetDataSource (mapHandle, 1)
datasourceHandle1 = SSProject.GetDataSource (mapHandle, 0)
SSProject.SetActiveDatasource mapHandle , datasourceHandle1
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Code", "<>", updateRegionCode
SSProcess.SelectFilter
SSProcess.DeleteSelectionObj
SSProject.SetActiveDatasource mapHandle , datasourceHandle10
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_LayerName", "<>", updateRegionLayer
SSProcess.SelectFilter
SSProcess.SelectionObjToClipBoard
SSProject.CloseDataSource mapHandle , datasourceHandle0
End Sub
```


#### CloseMap 关闭地图

**函数**：`CloseMap`

**功能**：关闭地图。

**语法**：

```vbs
SSProject.CloseMap(mapHandle)
```

**参数**：

- mapHandle 地图句柄

**示例**：

```vbs
Sub OnClick()
mapHandle = SSProject.GetActiveMap
SSProject. CloseMap mapHandle
End Sub
```


#### CreateDataSource 创建数据源

**函数**：`CreateDataSource`

**功能**：创建数据源。

**语法**：

```vbs
SSProject.CreateDataSource(datasourceHandle, datasourceName, mapScale, datasourceType)
```

**参数**：

- datasourceHandle 数据源句柄
- datasourceName 数据源名称
- mapScale 地图比例尺
- datasourceType 数据源类型


#### CreateLayer 创建图层

**函数**：`CreateLayer`

**功能**：创建图层。

**语法**：

```vbs
SSProject.CreateLayer(layerHandle, layerName)
```

**参数**：

- layerHandle 图层句柄
- layerName 层名


#### CreateMap 创建地图

**函数**：`CreateMap`

**功能**：创建地图。

**语法**：

```vbs
SSProject.CreateMap(templateName, mapName)
```

**参数**：

- templateName 模板名称
- mapName 地图名称


#### DeleteLayer 删除图层

**函数**：`DeleteLayer`

**功能**：删除图层。

**语法**：

```vbs
SSProject.SSProject.DeleteLayer( layerName)
```

**参数**：

- layerName 要删除的图层名称

**说明**：

执行该函数，将先删除图层内的所有对象，后删除图层。

**示例**：

```vbs
删除"居民地"图层。
Sub OnClick()
SSProcess.DeleteLayer "居民地"
End Sub
```


#### GetActiveDatasource 获取活动数据源句柄

**函数**：`GetActiveDatasource`

**功能**：获取活动数据源句柄。

**语法**：

```vbs
SSProject.GetActiveDatasource(mapHandle)
```

**参数**：

- mapHandle 数据源句柄

**说明**：

执行该函数返回活动数据源句柄。

**示例**：

```vbs
Sub OnClick()
Msgbox SSProject.GetActiveDatasource(mapHandle)
End Sub
```


#### GetActiveMap 获取活动地图句柄

**函数**：`GetActiveMap`

**功能**：获取活动地图句柄。

**语法**：

```vbs
SSProject.GetActiveMap()
```

**说明**：

执行该函数，将返回当前活动工程的句柄。

**示例**：

```vbs
Sub OnClick()
Msgbox SSProject.GetActiveMap()
End Sub
```


#### GetActiveProject 获取活动工程句柄

**函数**：`GetActiveProject`

**功能**：获取活动工程句柄。

**语法**：

```vbs
SSProject.GetActiveProject()
```

**说明**：

直接返回当前活动工程句柄。

**示例**：

```vbs
Sub OnClick()
Msgbox SSProject.GetActiveProject()
End Sub
```


#### GetDataSource 获取数据源句柄

**函数**：`GetDataSource`

**功能**：获取数据源句柄。

**语法**：

```vbs
SSProject.GetDataSource(mapHandle, index)
```

**参数**：

- mapHandle 地图句柄
- index 序号

**示例**：

```vbs
Sub OnClick()
keyName="mapScale"
mapHandle=SSProject.GetMap(0)
Msgbox SSProject.GetDataSource(mapHandle, 0)
End Sub
```


#### GetDataSourceByname 根据数据源名称获取数据源句柄

**函数**：`GetDataSourceByname`

**功能**：根据数据源名称获取数据源句柄。

**语法**：

```vbs
SSProject.GetDataSourceByname(mapHandle, datasourceName)
```

**参数**：

- mapHandle 地图句柄
- datasourceName 数据源名称

**说明**：

该函数根据数据源名称获取数据源句柄。


#### GetDataSourceInfo 获取数据源信息

**函数**：`GetDataSourceInfo`

**功能**：获取数据源信息。

**语法**：

```vbs
SSProject.GetDataSourceInfo(datasourceHandle, keyName)
```

**参数**：

- datasourceHandle 数据源句柄
- keyName 关键字


#### GetDataSourceLayer 获取数据源的图层句柄

**函数**：`GetDataSourceLayer`

**功能**：获取数据源的图层句柄。

**语法**：

```vbs
SSProject.GetDataSourceLayer(datasourceHandle, index)
```

**参数**：

- datasourceHandle 数据源句柄
- index  序号

**说明**：

执行该函数，将获取数据源的图层句柄。


#### GetDataSourceLayerByname 根据层名获取数据源的图层句柄

**函数**：`GetDataSourceLayerByname`

**功能**：根据层名获取数据源的图层句柄。

**语法**：

```vbs
SSProject.GetDataSourceLayerByname(datasourceHandle, layerName)
```

**参数**：

- datasourceHandle 数据源句柄
- layerName 层名

**说明**：

执行该函数，将根据层名获取数据源的图层句柄。

**示例**：

```vbs
Sub OnClick()
covName= “房屋面”
layerHandle= SSProject.GetDataSourceLayerByname(datasourceHandle ,covName)
End Sub
```


#### GetDataSourceLayerCount 获取数据源的图层数

**函数**：`GetDataSourceLayerCount`

**功能**：获取数据源的图层数。

**语法**：

```vbs
SSProject.GetDataSourceLayerCount(datasourceHandle)
```

**参数**：

- datasourceHandle 数据源句柄

**示例**：

```vbs
Sub OnClick()
mapHandle=SSProject.GetActiveMap
datasourceHandle = SSProject.GetActiveDatasource(mapHandle)
Msgbox SSProject.GetDataSourceLayerCount(datasourceHandle)
End Sub
```


#### GetLayerInfo 获取图层信息

**函数**：`GetLayerInfo`

**功能**：获取图层信息。

**语法**：

```vbs
SSProject.GetLayerInfo(layerHandle, keyName)
```

**参数**：

- layerHandle 图层句柄
- keyName 关键字

**示例**：

```vbs
Sub OnClick()
mapHandle=SSProject.GetActiveMap
datasourceHandle = SSProject.GetActiveDatasource(mapHandle)
layerHandle = SSProject.GetDataSourceLayer(datasourceHandle, 0)
Msgbox layerHandle
keyName="layerName"
Msgbox SSProject.GetLayerInfo(layerHandle, keyName)
End Sub
```


#### GetMap 获取地图句柄

**函数**：`GetMap`

**功能**：获取地图句柄。

**语法**：

```vbs
SSProject.GetMap(index)
```

**参数**：

- index 地图的索引号,从0开始

**说明**：

执行该函数，获取指定地图索引的句柄。

**示例**：

```vbs
Sub OnClick()
SSProject.GetMap(0)
End Sub
```


#### GetMapCount 获取地图个数

**函数**：`GetMapCount`

**功能**：获取地图个数。

**语法**：

```vbs
SSProject.GetMapCount()
```

**说明**：

执行该函数，可以获取当前工程地图的个数。

**示例**：

```vbs
Sub OnClick()
MapCount= SSProject.GetMapCount()‘获取当前工程地图个数
Msgbox MapCount
End Sub
```


#### GetMapDataSourceCount 获取地图数据源个数

**函数**：`GetMapDataSourceCount`

**功能**：获取地图数据源个数。

**语法**：

```vbs
SSProject.GetMapDataSourceCount(mapHandle)
```

**示例**：

```vbs
Sub Click()
mapHandle = SSProject.GetActiveMap
datasourceCount = SSProject.GetMapDataSourceCount (mapHandle)
If datasourceCount<>2 Then
Msgbox "请插入对比图！" : Exit Sub
End If
End Sub
```


#### GetMapInfo 获取地图信息

**函数**：`GetMapInfo`

**功能**：获取地图信息。

**语法**：

```vbs
SSProject.GetMapInfo(mapHandle, keyName)
```

**参数**：

- mapHandle 地图句柄
- keyName 需要获取的信息名称

**示例**：

```vbs
Sub OnClick()
keyName="mapScale"
mapHandle=SSProject.GetMap(0)
Msgbox SSProject.GetMapInfo(mapHandle, keyName)
End Sub
```


#### GetProject 获取工程句柄

**函数**：`GetProject`

**功能**：获取工程句柄。

**语法**：

```vbs
SSProject.GetProject(index)
```

**参数**：

- index 工程序号
- 函数直接返回打开的第index工程的工程句柄。

**说明**：

执行该函数，获取系统传回的所对应的特定值（日期型）。

**示例**：

```vbs
Sub OnClick()
prjCount = SSProject.GetProjectCount  '获取工程个数
curProjectHandle = SSProject.GetActiveProject '获取当前活动的工程句柄
For i=0 To prjCount -1
projectHandle = SSProject.GetProject(i)  '获取第i 个工程句柄
If projectHandle <> curProjectHandle Then
SSProject.SetActiveProject projectHandle '设置活动工程句柄
SSProcess.UpdateCurMap projectHandle '更新当前地图
End If
Next
End Sub
```


#### GetProjectCount 获取工程个数

**函数**：`GetProjectCount`

**功能**：获取工程个数。

**语法**：

```vbs
SSProject.GetProjectCount()
```

**示例**：

```vbs
Sub OnClick()
prjCount = SSProject.GetProjectCount  '获取工程个数
curProjectHandle = SSProject.GetActiveProject '获取当前活动的工程句柄
For i=0 To prjCount- 1
projectHandle= SSProject.GetProject(i)  '获取第i 个工程句柄
If projectHandle <> curProjectHandle Then
SSProject.SetActiveProject projectHandle '设置活动工程句柄
SSProcess.UpdateCurMap projectHandle '更新当前地图
End If
Next
End Sub
```


#### OpenDataSource 打开数据源

**函数**：`OpenDataSource`

**功能**：打开数据源。

**语法**：

```vbs
SSProject.OpenDataSource(datasourceHandle, datasourceName, datasourceType)
```

**参数**：

- datasourceHandle 数据源句柄
- datasourceName 数据源名称
- datasourceType 数据源类型


#### OpenMap 打开地图

**函数**：`OpenMap`

**功能**：打开地图。

**语法**：

```vbs
SSProject.OpenMap(mapName)
```

**参数**：

- mapName 需要打开的地图名称

**说明**：

一个工程文件一般情况只有一个地图，对于特殊业务比如房产图、地籍宗地图，就会在一个工程存在多张地图，此函数是用来打开那些有多张地图的EDB数据。


#### RefreshWorkspace 刷新工作空间窗口

**函数**：`RefreshWorkspace`

**功能**：刷新工作空间窗口。

**语法**：

```vbs
SSProject.RefreshWorkspace()
```

**说明**：

执行该函数可以刷新工作空间窗口，该函数不单独使用，EPS功能模块中已经添加了该函数。（例如执行画线功能时，画完线就会自动刷新工作空间窗口）。


#### SetActiveDatasource 设置活动数据源

**函数**：`SetActiveDatasource`

**功能**：设置活动数据源。

**语法**：

```vbs
SSProject.SetActiveDatasource(mapHandle, datasourceHandle)
```

**参数**：

- mapHandle 地图句柄，
- datasourceHandle 数据源句柄

**说明**：

执行该函数，将datasourceName设为活动数据源。

**示例**：

```vbs
Sub OnClick()
'切换活动工程
mapHandle=SSProject.GetActiveMap
datasourceHandle=SSProject.GetDataSourceByname(mapHandle, edbFile )
'设置活动数据源
SSProject.SetActiveDatasource mapHandle, datasourceHandle
End Sub
```


#### SetActiveDataSourceByname 根据数据源名称设置活动数据源

**函数**：`SetActiveDataSourceByname`

**功能**：根据数据源名称设置活动数据源。

**语法**：

```vbs
SSProject.SetActiveDataSourceByname(mapHandle, datasourceName)
```

**参数**：

- mapHandle 地图句柄
- datasourceName 数据源名称

**说明**：

执行该函数，将指定名称的数据源设为活动数据源。

**示例**：

```vbs
Sub Click()
'切换活动工程
mapHandle=SSProject.GetActiveMap
datasourceHandle=SSProject.GetDataSourceByname(mapHandle, edbFile)
'设置活动数据源
SSProject.SetActiveDataSourceByname mapHandle,datasourceName
End Sub
```


#### SetActiveMap 设置活动地图

**函数**：`SetActiveMap`

**功能**：设置活动地图。

**语法**：

```vbs
SSProject.SSProject.SetActiveMap(mapHandle)
```

**参数**：

- mapHandle 地图句柄

**说明**：

执行该函数，将把指定的地图设置为活动地图。


#### SetActiveProject 设置活动工程

**函数**：`SetActiveProject`

**功能**：设置活动工程。

**语法**：

```vbs
SSProject.SetActiveProject(projectHandle)
```

**参数**：

- projectHandle 工程句柄

**说明**：

执行该函数，获取系统传回的所对应的特定值（日期型）。

**示例**：

```vbs
切换当前工程为另外打开的工程：
Sub OnClick()
prjCount = SSProject.GetProjectCount '获取工程个数
curProjectHandle = SSProject.GetActiveProject '获取当前活动的工程句柄
For i=0 To prjCount- 1
projectHandle = SSProject.GetProject(i) '获取第i 个工程句柄
If projectHandle <> curProjectHandle Then
SSProject.SetActiveProject projectHandle '设置活动工程句柄
SSProcess.UpdateCurMap projectHandle '更新当前地图
End If
Next
End Sub
```


#### SetDataSourceInfo 设置数据源信息

**函数**：`SetDataSourceInfo`

**功能**：设置数据源信息。

**语法**：

```vbs
SSProject.SetDataSourceInfo(datasourceHandle, keyName, keyValue)
```

**参数**：

- datasourceHandle 数据源句柄
- keyName 关键字
- keyValue 关键值


#### SetLayerInfo 设置图层信息

**函数**：`SetLayerInfo`

**功能**：设置图层信息。

**语法**：

```vbs
SSProject.SetLayerInfo(layerHandle, keyName, keyValue)
```

**参数**：

- layerHandle 图层句柄
- keyName 关键字
- keyValue 关键值


#### SetMapInfo 设置地图信息

**函数**：`SetMapInfo`

**功能**：设置地图信息。

**语法**：

```vbs
SSProject.SetMapInfo(mapHandle, keyName, keyValue)
```

**参数**：

- mapHandle 地图句柄
- keyName 关键字
- keyValue 值


---

## 常用脚本函数 对象参考

### 获取指定目录下指定扩展名的所有文件

#### GetAllFiles 获取指定目录下指定扩展名的所有文件

**函数**：`GetAllFiles`

**功能**：获取指定目录下指定扩展名的所有文件。

**参数**：

- pathname 要获取文件的主目录
- fileExt 文件扩展名
- filecount 返回的文件数
- filenames 返回的文件名数组

**说明**：

执行该函数,将返回指定目录下指定扩展名的所有文件名，同时搜索目录下的所有子目录中的文件。
在调用函数前，必须给数组filenames定义空间。

**示例**：

```vbs
下面示例获取c:\test\目录下所有扩展名为MDB的文件。
Sub OnClick()
Dim filenames(1000)
Dim filecount
filecount = 0
GetAllFiles "c:\test\", "Mdb",filecount, filenames
End Sub
```


---
*（内容由AI生成，仅供参考）*
