# Excel Eval 特性覆盖策略

> 本文档分析 AI 生成 Excel 场景下，评估框架应优先覆盖哪些 Excel 特性。
> 核心原则：**eval 工具的优先级跟着"AI 生成 Excel 的能力"走，而不是跟着"Excel 本身的市场前景"走。**

---

## 1. 当前覆盖现状

项目目前有 **8 个评估维度**：

| 维度 | 评估内容 | 覆盖深度 |
|------|---------|---------|
| DataAccuracy | 数值正确性、聚合校验 | ✅ 较完善 |
| Completeness | 要求的 sheet/表格/指标是否齐全 | ✅ 较完善 |
| FormulaLogic | 公式存在性、错误检测、硬编码 | ⚠️ 基础公式，缺高级公式 |
| ChartAppropriateness | 图表类型、标题、轴标签 | ⚠️ 基础元数据，缺高级图表 |
| ProfessionalFormatting | 字体/颜色/边框/条件格式/冻结窗格 | ⚠️ 条件格式仅检查存在性 |
| TableStructure | 表头、数据类型、合并单元格 | ⚠️ 缺 Excel Table 对象评估 |
| SheetOrganization | Sheet 命名/顺序/跨 Sheet 引用 | ✅ 较完善 |
| Relevance | 内容相关性 | ✅ 较完善 |

**核心差距：** 解析器能提取的信息有限，很多 Excel 原生特性（透视表、数据验证、保护等）既没有解析也没有评估。

---

## 2. 优先级决策框架：双维象限

用两个维度交叉判断每个特性是否值得投入：

- **X 轴：AI 生成该特性的可行性** — 当前主流 AI（GPT/Claude + openpyxl/xlsxwriter）能否正确生成这个特性？
- **Y 轴：用户场景中的使用频率** — 在 AI 生成 Excel 的典型场景（报表、模板、分析）中，这个特性出现的频率有多高？

```
                     AI 生成可行性
                低 ◄──────────────► 高

            ┌──────┤───────────────────│
            │      │                   │
     高     │  ⚪  │     🔴 必须做     │
     |      │观察区 │                   │
     |      │      │                   │
  使用频率   ├──────┤───────────────────│
     |      │      │                   │
     |      │  ❌  │     🟡 可以做     │
     低     │ 不做  │                   │
            │      │                   │
            └──────┴───────────────────┘
```

---

## 3. 特性分类详表

### 🔴 必须做（高频 + AI 能生成）

这些特性在真实 Excel 中高频出现，且 AI 通过 openpyxl 等库可以正确生成。
当前 eval 不覆盖或覆盖不足，是最大的评估盲区。

| 特性 | 当前状态 | 需要做什么 | 理由 |
|------|---------|-----------|------|
| **数据透视表** | ❌ 未覆盖 | 新增解析 + 新增评估维度或扩展现有维度 | 企业报表核心功能，openpyxl 可读取 PivotTable 定义 |
| **数据验证（下拉列表等）** | ❌ 未覆盖 | 新增解析 + 评估验证规则正确性 | 模板/表单类文件必备，AI 可通过 openpyxl 生成 |
| **条件格式规则正确性** | ⚠️ 仅检查存在性 | 深化评估：检查规则逻辑、范围、优先级 | 当前只知道"有条件格式"，不知道"对不对" |
| **Excel Table (ListObject)** | ⚠️ 提及但未深入 | 解析表对象、评估结构化引用使用 | 微软推荐的数据组织方式，AI 生成的表格应优先用 Table |
| **命名范围正确性** | ⚠️ 仅检测存在 | 验证命名范围定义是否正确、是否被公式使用 | 公式可读性和模型质量的关键指标 |

### 🟡 应该做（高频或中频 + AI 能生成，但优先级次于红色）

这些特性有明确价值，但紧迫度低于第一批。

| 特性 | 当前状态 | 需要做什么 | 理由 |
|------|---------|-----------|------|
| **工作表保护 & 单元格锁定** | ❌ 未覆盖 | 解析保护设置，评估锁定策略是否合理 | 模板类文件应保护公式区、开放输入区 |
| **打印布局** | ❌ 未覆盖 | 解析打印区域/标题行/页面方向 | 交付物需打印时必须正确，openpyxl 可设置 |
| **组合图表 & 双轴** | ⚠️ 基础图表覆盖 | 扩展图表评估：识别组合类型、双轴合理性 | Dashboard 类报表常用 |
| **迷你图 (Sparklines)** | ❌ 未覆盖 | 解析迷你图存在性和数据范围 | Dashboard 场景有价值，openpyxl 支持读取 |
| **数据标签 & 图表细节** | ⚠️ 基础覆盖 | 评估数据标签、趋势线等是否恰当 | 专业图表的完成度指标 |

### ⚪ 观察区（高频但 AI 当前难以正确生成）

这些特性在真实 Excel 中常见，但 AI 当前技术很难正确生成。
暂时不投入评估开发，但保持关注，等 AI 能力提升后再加入。

| 特性 | 为什么 AI 难生成 | 什么时候该加入 |
|------|-----------------|---------------|
| **VBA / 宏** | 需要 .xlsm 格式 + 嵌入可执行代码，安全风险大，大多数 AI 生成管线不支持 | 当主流 AI 工具开始生成 .xlsm 时 |
| **复杂事件处理** | 依赖 VBA 事件模型 | 同上 |
| **按钮 / 表单控件** | 需要 ActiveX 或 Form Control 对象模型 | 同上 |

### ❌ 不做（低频 + AI 难生成）

投入产出比最低，不建议覆盖。

| 特性 | 为什么不做 |
|------|-----------|
| **Power Query / Power Pivot** | 二进制格式存储，openpyxl 无法读写，AI 也无法生成 |
| **Solver / Goal Seek** | 设置保存在工作簿内部二进制区域，几乎不可能通过代码生成 |
| **外部数据连接** | 涉及服务器地址/凭据，AI 不应该生成 |
| **ActiveX 控件** | 过时技术，安全隐患，不值得投入 |
| **嵌入 OLE 对象** | 技术复杂度高，使用频率低 |

---

## 4. 实施策略：扩展现有维度，而非新增维度

### 4.1 设计决策

**结论：不新增评估维度，通过"扩展解析器 + 深化 prompt"来覆盖缺失特性。**

原因：

1. **每个维度 = 1 次独立 LLM 调用**。新增维度意味着每个 case 多一次调用，成本 +12.5%。
2. **权重标定成本高**。当前 7 个场景 × 8 个维度 = 56 个权重值，基于 62 个 GDPVal case 标定。新增维度需要全部重新标定。
3. **当前 8 个维度已覆盖正确的质量属性**。缺的不是"评估角度"，而是解析深度和 prompt 精度。

```
当前问题不是：                     而是：
"缺少评估数据验证的维度"           "TableStructure 维度看不到数据验证信息"
"缺少评估条件格式的维度"           "ProfessionalFormatting 的 prompt 只检查存在性"
"缺少评估透视表的维度"             "解析器根本没提取透视表数据"
```

### 4.2 特性 → 现有维度映射

| 缺失特性 | 归入维度 | 改动层 |
|---------|---------|-------|
| **数据透视表** | `TableStructure` | 解析器 + prompt |
| **数据验证（下拉列表等）** | `TableStructure` | 解析器 + prompt |
| **条件格式规则正确性** | `ProfessionalFormatting` | prompt（解析器已有基础数据） |
| **Excel Table (ListObject)** | `TableStructure` | 解析器 + prompt |
| **命名范围正确性** | `FormulaLogic` | prompt（解析器已有基础数据） |
| **工作表保护 & 单元格锁定** | `SheetOrganization` | 解析器 + prompt |
| **打印布局** | `ProfessionalFormatting` | 解析器 + prompt |
| **组合图表 / 双轴** | `ChartAppropriateness` | 解析器 + prompt |
| **迷你图 (Sparklines)** | `ChartAppropriateness` | 解析器 + prompt |
| **数据标签 / 趋势线** | `ChartAppropriateness` | prompt |

### 4.3 改动分两层

```
Layer 1: 解析器扩展（excel_parser.py + models.py）
│
│  扩展 PreparedData，让 LLM 能"看到"这些特性：
│  ├── PivotTableInfo:  source_range, rows, columns, values, filters
│  ├── ValidationInfo:  cell_range, type, formula, dropdown_values
│  ├── TableObjectInfo: name, range, has_structured_refs, total_row
│  ├── ProtectionInfo:  sheet_protected, locked_cells, unlocked_cells
│  └── PrintLayoutInfo: print_area, title_rows, orientation, scaling
│
▼
Layer 2: Prompt 深化（prompts/*.md）
│
│  在现有维度的 prompt 中增加评估要点：
│  ├── table_structure.md:     + 透视表合理性 + 数据验证 + Table 对象
│  ├── professional_formatting.md: + 条件格式规则正确性 + 打印布局
│  ├── formula_logic.md:       + 命名范围使用质量 + 结构化引用
│  ├── chart_appropriateness.md:   + 组合图表 + 迷你图
│  └── sheet_organization.md:  + 保护策略合理性
│
▼
不需要动的：
  ├── DimensionName 枚举 — 不变
  ├── SCENARIO_WEIGHTS — 不变
  ├── Pipeline 逻辑 — 不变
  └── 报告格式 — 不变
```

### 4.4 何时评估新特性：三层判断逻辑

新增的特性评估点不是每个 case 都适用。判断是否评估需要三层逻辑：

```
Layer 1: 维度级别 — 该维度是否适用？
│  现有机制：ScenarioDetector 输出 applicable_dimensions
│  例：纯数据清洗任务 → ChartAppropriateness = N/A
│
▼
Layer 2: 特性级别 — 维度内的某个评估点是否适用？
│  由 LLM 在 prompt 内根据 user_prompt + 文件内容自行判断
│  例：TableStructure 维度内，"透视表"评估点是否适用？
│
▼
Layer 3: 评判逻辑 — "有了评好坏"还是"没有要扣分"？
   ├── 用户要求了 + 没生成 → 扣分
   ├── 用户没要求 + 没生成 → 不扣分也不加分
   ├── 用户没要求 + 生成了且合理 → 加分
   └── 场景明显适合但没用 → 轻微建议（不重扣）
```

**核心设计：不在代码层做硬规则，而是在 prompt 中给 LLM 判断指引。**

以透视表为例，prompt 中应写明：

```
评估透视表使用的合理性：
1. 用户明确要求了透视表/汇总分析：
   - 生成了且合理 → 正面评价
   - 没生成，用 SUMIFS 硬拼 → 扣分
   - 没生成，用了其他合理方式 → 中性
2. 用户没明确要求：
   - 场景适合但没用（大量明细需要分组汇总） → 轻微建议
   - 场景不需要 → 跳过，不影响评分
3. 文件中有透视表 → 评估数据源、字段配置、汇总方式是否合理
```

解析器层面只需**如实提取**，不做判断：有透视表 → 提取详情；没有 → 空列表。LLM 结合 user_prompt 和文件内容自行决定"没有是否合理"。

这与现有维度的工作方式一致——`ChartAppropriateness` 也不是"没图表就扣分"，而是"该有的时候没有才扣分"。

### 4.5 何时拆分维度

如果实践中发现某个维度塞了太多评估点，导致 **prompt 过长或 LLM 评分不稳定**，再考虑拆分。
例如：从 `TableStructure` 中拆出 `DataInteractivity`（透视表 + 数据验证 + 保护）。
这应该是遇到实际问题后的决策，不预设。

---

## 5. 实施路线

```
Phase 1 — 解析器扩展 + 🔴 核心 prompt 深化
│
├── 1a. excel_parser.py: 新增透视表、数据验证、Table 对象解析
├── 1b. models.py: 新增 PivotTableInfo, ValidationInfo, TableObjectInfo
├── 1c. table_structure.md: prompt 增加透视表/验证/Table 评估点
├── 1d. professional_formatting.md: prompt 深化条件格式规则评估
├── 1e. formula_logic.md: prompt 深化命名范围评估
│
▼
Phase 2 — 🟡 次优先特性
│
├── 2a. excel_parser.py: 新增保护状态、打印布局解析
├── 2b. models.py: 新增 ProtectionInfo, PrintLayoutInfo
├── 2c. sheet_organization.md: prompt 增加保护策略评估
├── 2d. professional_formatting.md: prompt 增加打印布局评估
├── 2e. chart_appropriateness.md: prompt 扩展组合图表/迷你图
│
▼
Phase 3 — 持续观察
│
└── 跟踪 AI 生成 .xlsm / VBA 的能力进展，适时加入
```

---

## 6. 判断标准参考

当评估是否要加入某个新特性时，用以下 checklist：

- [ ] **AI 能生成吗？** — 主流 Python 库（openpyxl/xlsxwriter）能否创建该特性？
- [ ] **解析器能提取吗？** — openpyxl 或其他库能否从 .xlsx 中读取该特性的结构化数据？
- [ ] **LLM 能评判吗？** — 该特性的"好坏"能否通过文本/截图让 LLM 评分？
- [ ] **用户场景需要吗？** — 在项目的目标场景（报表、模板、分析）中是否常见？

四项都满足 → 🔴 必须做
三项满足 → 🟡 应该做
两项以下 → ⚪ 观察 或 ❌ 不做
