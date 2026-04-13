# 计划：标题/子标题动态映射到 Word 模板

## 现状问题分析

当前代码使用"文本字符串查找替换"方式映射标题：

```
模板中的原文 "1. 断裂保险轴" → 在文档中搜索这个字符串 → 替换为用户修改后的标题
```

**存在的风险：**

1. **依赖硬编码默认值**：`DEFAULT_SECTION_TITLES` 和 `IMAGE_SLOTS` 中的默认文本必须与 Word 模板中的文字**完全一致**（含空格、标点），否则替换静默失败
2. **仅首次生效**：用户生成报告后模板中的文字已被替换，若不重新从原始模板复制，第二次生成时找不到原文
3. **子标题重复问题**：`{{image_1.1}}` 占位符被替换为三层布局时，描述层直接用了当前 label（正确）；但同时 `title_mapping` 也会尝试在文档其他位置替换原始 label 文本——两套逻辑可能冲突
4. **用户新增章节的标题不在模板中**：新章节走动态插入逻辑（847-856行），标题直接写入，没有占位符问题；但如果模板结构变化，锚点定位可能错位

## 方案：改用模板占位符

将 Word 模板中的章节标题和子标题改为 `{{...}}` 占位符，与图片占位符 `{{image_1.1}}` 保持一致的风格。

### 具体改动

#### Phase 1: 修改 Word 模板

在 `模版.docx` 中将原有的硬编码标题文本替换为占位符：

| 原始文本 | 替换为 |
|----------|--------|
| `1. 断裂保险轴` | `{{sec_title_1}}` |
| `2. 宏观断口形貌` | `{{sec_title_2}}` |
| `3. 微观断口形貌` | `{{sec_title_3}}` |
| `4. 显微组织分析` | `{{sec_title_4}}` |
| `5. 化学分析` | `{{sec_title_5}}` |
| `6. 性能验证` | `{{sec_title_6}}` |
| 子标题如 `1.1 保险轴断裂部位` | `{{slot_label_1.1}}` |
| ... | ... |

**优点：** 占位符是稳定的唯一标识，不随用户编辑而变化。

#### Phase 2: 修改后端 `report_tool.py`

**2a. 删除硬编码的默认文本映射逻辑（752-789行）**

删除以下逻辑：
- `DEFAULT_SECTION_TITLES` 字典中的标题字符串（仅保留 sec_id 列表用于 VLM prompt 选择）
- 遍历 `IMAGE_SLOTS` 做 label 映射的代码块（772-789行）
- 遍历 `DEFAULT_SECTION_TITLES` 做 title 映射的代码块（752-770行）

**2b. 新增：构建占位符映射**

从 `image_config` 直接构建映射：

```python
# 章节标题映射
title_mapping = {}
for section in image_config:
    sec_id = section["id"]
    # 从 sec_id 提取编号，如 "sec_1" -> "1"
    num = sec_id.split("_")[-1] if sec_id.startswith("sec_") else ""
    if num:
        title_mapping[f"{{{{sec_title_{num}}}}}"] = section["title"]

# 子标题映射
for section in image_config:
    for slot in section.get("slots", []):
        title_mapping[f"{{{{slot_label_{slot['id']}}}}}"] = slot["label"]
```

**2c. `IMAGE_SLOTS` 简化**

`IMAGE_SLOTS` 目前用于：
- 提供默认 label（用于旧映射逻辑）→ **删除，改用 image_config 中的 label**
- 判断是否 fixed_size → **保留 fixed_size 判断逻辑，但改为从 image_config 传入**
- 模板占位符 `{{image_1.1}}` → **不变，slot_id 不变则占位符不变**

`IMAGE_SLOTS` 可以简化为只保留 fixed_size 的判断规则（1-4 章节固定尺寸，5-6 原样），或者直接在前端 image_config 中携带 fixed_size 字段。

#### Phase 3: 前端无需改动

前端 `imageConfig` 的结构不变：
- `section.title` 用户可编辑 → 通过 `{{sec_title_N}}` 占位符映射
- `slot.label` 用户可编辑 → 通过 `{{slot_label_X.Y}}` 占位符映射
- `slot.id` 不可变 → 图片占位符 `{{image_X.Y}}` 保持稳定

#### Phase 4: 处理边界情况

- **用户删除的章节**：对应 `{{sec_title_N}}` 映射为空字符串
- **用户删除的槽位**：对应 `{{slot_label_X.Y}}` 映射为空字符串，`{{image_X.Y}}` 清空
- **新增章节**：不走占位符替换，走现有的动态插入逻辑（847-856行），无需改动
- **新增槽位**：`slot_XXXXXX` 格式的 ID，模板中无对应占位符，走溢出插入逻辑（862-866行）

### 不改动的部分

- `{{项目名称}}`、`{{背景概述}}` 等文本字段占位符 → 不变
- `{{基本图片分析}}` 等 VLM 分析结果占位符 → 不变
- `{{分析与讨论}}`、`{{结论}}` 等 LLM 生成占位符 → 不变
- `{{image_X.Y}}` 图片占位符 → 不变
- 新增章节/槽位的动态插入逻辑 → 不变

### 实施顺序

1. 修改 Word 模板 `模版.docx`（手工或脚本替换）
2. 修改 `report_tool.py`：替换映射逻辑
3. 简化 `IMAGE_SLOTS` 常量
4. 验证完整流程
