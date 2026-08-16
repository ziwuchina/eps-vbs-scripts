# EPS VBS 脚本库

EPS2020 顺德基础地理测绘专版（云舟P）的 VBS 脚本集合，用于测绘业务提效。

本仓库沉淀了两类资产：
1. **可直接使用的脚本**（如 `导出全部属性.vbs`）
2. **EPS VBS 脚本编写指南**（踩坑经验 + 标准写法，见下文「EPS VBS 脚本编写指南」）

> 环境：EPS2020 顺德基础地理测绘专版（云舟P），脚本目录 `D:\EPS2020顺德基础地理测绘专版云舟P\DeskTop\顺德基础测绘\玮哥专用脚本`

---

## 目录

- [脚本列表](#脚本列表)
- [导出全部属性.vbs](#导出全部属性vbs)
- [EPS VBS 脚本编写指南](#eps-vbs-脚本编写指南)
  - [一、必须遵守（硬性要求）](#一必须遵守硬性要求)
  - [二、必须规避（坑）](#二必须规避坑)
  - [三、调试与验证](#三调试与验证)
  - [四、API 速查表](#四api-速查表)

---

## 脚本列表

| 脚本 | 功能 | 状态 |
|------|------|------|
| [导出全部属性.vbs](导出全部属性.vbs) | 导出所选要素的全部属性（基本属性 + 全量扩展属性）为 CSV | 可用 |

---

## 导出全部属性.vbs

### 功能

将 EPS 中选中（或按类型筛选）的要素全部属性导出为 CSV 文件，包括：

- **基本属性**：ID、编码、图层、类型、颜色、面积、长度、注记文字
- **全量扩展属性**：150+ 个字段（FeatureGUID、ZDGUID、LJZH、CH、SJC、KZMJ、JZMJ、GHJZMJ、JRMJ、YeWBH 等），覆盖 EPS 属性表「扩展属性」组全部字段，并自动合并选中要素中清单之外的额外属性

### 使用说明

1. 在 EPS 中**框选**要导出的要素（地物 / 注记均可）
2. 运行脚本 `导出全部属性.vbs`
3. 参数对话框：
   - **要素类型**：POINT / LINE / AREA / NOTE / 全部（默认「全部」）
   - **导出范围**：选择集 / 全部（默认「选择集」，直接导出当前选中的要素）
   - **输出文件**：CSV 输出路径（默认 `C:\全部属性导出.csv`）
4. 点击确定，脚本自动导出并提示完成

### 输出说明

- CSV 表头 = 基本属性列 + 全量扩展属性清单 + 动态发现的额外属性
- 每个要素一行；要素没有的字段留空
- 含逗号 / 引号 / 换行的字段自动做 CSV 转义

### 实现要点

- **全量属性清单硬编码**：来自 EPS 属性表截图（90+ 字段）+ 脚本目录其他脚本用到的扩展属性，共 150+ 列，保证无论选中什么要素都导出完整列结构
- **MemoData 动态枚举**：遍历选中要素的 `SSObj_MemoData`，自动合并清单之外的属性名，防止遗漏
- **双通道取值**：每个字段先解析 MemoData（属性名不带括号），取不到再用 `GetSelGeoValue(i, "[字段名]")`（带括号）兜底，确保 GUID 类等不在 MemoData 里的字段也能导出
- **选择集模式不清空选择**：直接 `UpdateSysSelection 0` 读取当前系统选择，避免 `ClearSelection` 清掉用户已选要素

---

## EPS VBS 脚本编写指南

> 以下内容基于玮哥专用脚本目录 50+ 个脚本的反复验证，是编写 EPS VBS 脚本的**经验基线**。

### 一、必须遵守（硬性要求）

#### 1. 文件编码必须 GBK / ANSI

**这是最容易踩、也最致命的坑。** 含中文注释的 UTF-8 文件会被 EPS 脚本宿主按 ANSI/GBK 解析，中文注释读成乱码，触发「未结束的字符串常量」语法错误，**脚本静默加载失败（点了没反应，无任何报错）**。

- 保存脚本必须用 GBK(936) / ANSI 编码，无 BOM
- 若用工具生成的是 UTF-8，落盘后必须转码：

```powershell
[System.IO.File]::WriteAllText($p, [System.IO.File]::ReadAllText($p,[System.Text.Encoding]::UTF8), [System.Text.Encoding]::GetEncoding("GBK"))
```

#### 2. 语法基线以目录现有脚本为准

EPS 脚本宿主只支持**最基础的 VBScript 语法**。编写前先对照玮哥目录内已有正常脚本的写法基线，**只使用目录中反复出现过的语法**。

#### 3. 参数对话框标准写法

```vbs
SSProcess.ClearInputParameter
SSProcess.AddInputParameter "参数名", "默认值", 0, "选项1,选项2", ""
res = SSProcess.ShowInputParameterDlg("对话框标题")
If res = 0 Then Exit Sub
SSProcess.UpdateScriptDlgParameter 1
值 = SSProcess.GetInputParameter("参数名")
```

- `AddInputParameter` 第 4 个参数是下拉选项（逗号分隔），第 5 个参数留空
- 用 `ReadEpsIni` / `WriteEpsIni` 记忆上次输入，提升体验

#### 4. 选择集标准写法

```vbs
' 方式一：按条件筛选
SSProcess.ClearSelection
SSProcess.ClearSelectCondition
SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"   ' 支持 POINT,LINE,AREA,NOTE
SSProcess.SelectFilter

' 方式二：直接使用当前系统选择（用户已框选）
SSProcess.UpdateSysSelection 0

' 读取数量与取值
geoCount = SSProcess.GetSelGeoCount()
noteCount = SSProcess.GetSelNoteCount()
值 = SSProcess.GetSelGeoValue(i, "SSObj_Code")
```

#### 5. 扩展属性取值标准写法

- **基本属性**：`GetSelGeoValue(i, "SSObj_前缀")`，如 `SSObj_ID` / `SSObj_Code` / `SSObj_LayerName` / `SSObj_Area` / `SSObj_Length` / `SSObj_Color` / `SSObj_Type` / `SSObj_FontString` / `SSObj_MemoData`
- **扩展属性**：`GetSelGeoValue(i, "[属性名]")`，**属性名必须带方括号**，如 `[ZL]`、`[DJH]`、`[ZDGUID]`
- **MemoData**：`SSObj_MemoData` 返回全部扩展属性的原始串，格式为 `属性名1|值1|属性名2|值2|...`（键值交替、`|` 分隔，**属性名不带方括号**）
- 注记用 `GetSelNoteValue(i, "SSObj_FontString")` / `GetSelNoteCount()`

#### 6. 写属性标准写法

```vbs
SSProcess.SetObjectAttr CLng(geoID), "[字段名]", 值
```

- ID 必须用 `CLng()` 转换
- 字段名带方括号

#### 7. 注记创建标准写法

```vbs
SSProcess.CreateNewObjByClass "3990202"
' 配合 SSObj_FontString（文字）、SSObj_FontHeight（字高）等
```

#### 8. 图层显隐标准写法

```vbs
SSProcess.GetLayerCount
SSProcess.GetLayerName(i)
SSProcess.SetLayerStatus 图层名, 0或1, 1
SSProcess.RefreshView
```

### 二、必须规避（坑）

以下语法在 EPS 脚本宿主中**很可能不被支持**，会导致脚本静默无反应（连对话框都不弹）。玮哥目录 50+ 脚本中**从未出现**过这些写法：

| 规避项 | 说明 | 替代方案 |
|--------|------|----------|
| `Scripting.Dictionary` | COM 字典对象，宿主不支持 | 用字符串拼接 + `InStr` 去重 |
| `Call` 语句 | 宿主不支持 | 直接调用，如 `SSProcess.SelectFilter` |
| `On Error` / `Err` | 错误处理不支持 | 不用错误处理，靠逻辑保证 |
| `vbCrLf` 常量 | 宿主不支持 | 用 `Chr(13)` |
| `ByRef` 参数 | 宿主不支持 | 用 Function 返回值 |
| 其他 COM 对象 | 除 `Scripting.FileSystemObject`（写文件）外尽量不用 | 用内置函数 |

> 经验：v1/v2 版脚本因使用 `Scripting.Dictionary`（及 Call / On Error / vbCrLf / ByRef）导致 EPS 中连对话框都不弹；v3 彻底移除这些语法后恢复正常。

### 三、调试与验证

EPS 脚本「点了没反应」排查两步法：

1. **cscript 语法检查**：`cscript //nologo 脚本.vbs`（OnClick 不调用，仅解析语法；语法错误会立即报错）
2. **桩对象模拟验证**：用 `Class` 实现 `FakeSSProcess` 桩对象（`GetSelGeoCount` / `GetSelGeoValue` 按索引返回预设数据），复制核心逻辑用 cscript 跑通，脱离 EPS 宿主验证算法正确性

> 注意：cscript 运行含中文的 UTF-8 脚本同样会报语法错误，验证前先转 GBK。

### 四、API 速查表

| 用途 | API | 说明 |
|------|-----|------|
| 选择集 | `ClearSelection` | 清空选择 |
| 选择集 | `ClearSelectCondition` | 清空选择条件 |
| 选择集 | `SetSelectCondition "SSObj_Type", "=", "AREA"` | 设置选择条件（支持 POINT,LINE,AREA,NOTE） |
| 选择集 | `SelectFilter` | 执行选择 |
| 选择集 | `UpdateSysSelection 0` | 使用当前系统选择（不清空） |
| 选择集 | `GetSelGeoCount()` | 选中地物数量 |
| 选择集 | `GetSelNoteCount()` | 选中注记数量 |
| 取值 | `GetSelGeoValue(i, "SSObj_Code")` | 基本属性（SSObj_ 前缀） |
| 取值 | `GetSelGeoValue(i, "[字段]")` | 扩展属性（**带方括号**） |
| 取值 | `GetSelGeoValue(i, "SSObj_MemoData")` | 扩展属性原始串（`名|值|名|值`） |
| 取值 | `GetSelNoteValue(i, "SSObj_FontString")` | 注记属性 |
| 写值 | `SetObjectAttr CLng(id), "[字段]", 值` | 写扩展属性（ID 需 CLng） |
| 参数 | `ClearInputParameter` | 清空参数 |
| 参数 | `AddInputParameter "名", 默认值, 0, "选项", ""` | 添加参数 |
| 参数 | `ShowInputParameterDlg("标题")` | 显示参数对话框 |
| 参数 | `GetInputParameter("名")` | 读取参数值 |
| 参数 | `UpdateScriptDlgParameter 1` | 更新参数 |
| 注记 | `CreateNewObjByClass "3990202"` | 创建注记 |
| 注记 | `SSObj_FontString` / `SSObj_FontHeight` | 注记文字 / 字高 |
| 图层 | `GetLayerCount` / `GetLayerName` | 图层列表 |
| 图层 | `SetLayerStatus 名, 0或1, 1` | 图层显隐 |
| 图层 | `RefreshView` | 刷新视图 |
| 配置 | `ReadEpsIni` / `WriteEpsIni` | 记忆上次输入 |

---

## 常见问题

**Q：脚本点了没反应？**
A：先查编码（必须 GBK），再查是否用了规避语法（Dictionary / Call / On Error / vbCrLf / ByRef）。

**Q：导出的属性列太少？**
A：确认用的是全属性版（v4）。旧版只动态枚举选中要素实际存在的字段，新版硬编码全量清单 + 动态合并。

**Q：选择集模式导不出东西？**
A：确认在 EPS 中已框选要素；选择集模式不清空选择，直接读取当前系统选择。
