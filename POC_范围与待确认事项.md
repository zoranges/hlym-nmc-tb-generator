# NMC TB 自动化 POC — 范围与待确认事项清单

| 项目 | 内容 |
|---|---|
| **项目代号** | 雅马哈二期 / NMC TB Issuance Automation POC |
| **业务部门** | HLYM — New Model Control Department (NMC) |
| **文档版本** | v0.1 (Draft) |
| **创建日期** | 2026-05-12 |
| **状态** | 待需求方对齐 |
| **基于资料** | `NMC Project/` 下 4 份文件（PDF + 3×xlsx） |

---

## 1. 项目背景

HLYM NMC 部门在新车型导入与变更过程中，需向工厂（F1–F5）、SCM、QA、Finance、Metal Sys 等内部相关方发行 **Technical Bulletin (TB)**，作为零件清单、颜色、图案、废弃件等信息的正式权威单据。

当前流程为**人工**：工程师在 PCL / YPL / CIT / QCPS / LN BOM / 邮件 等多张 Excel 与系统中查找、过滤、跨表拼装数据，按 13 种模板手工填入 TB Excel 模板。流程冗长、易错、强依赖个人经验。

POC 目标：用自动化工具替代上述人工拼装动作，给定源数据可自动生成符合规范的 TB 输出。

---

## 2. POC 范围建议

> 在需求方确认前，下述为**建议默认范围**，最终以「待确认事项」对齐后的版本为准。

### 2.1 In Scope（建议纳入）

- **TB 类型**：模板 **No.1–5（Master Part List For Factory 1–5）**，以示例 `TB-25-137-F3` 为基准产物
- **覆盖章节**：示例 TB 中的章节 **A、B、C、D**
  - A. PARTS WHICH SUPPLY FROM YMC TO F3
  - B. PARTS WHICH SUPPLY FROM YMTT, YMCN & IYM
  - C. PARTS WHICH SUPPLY FROM YMAC (YIMM, TYM…)
  - D. PARTS WHICH SUPPLY FROM LCP TO F3
- **零件来源处理**：
  - **YMC** 零件：按 `TB for YMC` SOP 7 步 + 3 个 SITUATION
  - **YMAC** 零件：按 `TB for YMAC` SOP + 3 个 SITUATION（含 PART SET NO `-**` 与颜色拼接）
  - **LCP** 零件：按 `TB for LCP` SOP（含 VMI `-S1` 子装件与颜色拼接）
- **核心业务规则**：
  - SEC/LVL1–4/PART NO./PART NAME/QTY 从 PCL 抽取
  - SUPPLIER 由 `CU_FA CODE` 表映射
  - EX/NEW 由 `LN BOM` 存在性判定
  - USE IN 由 `QCPS - Sub Assy / ML(R) / ML(L)` 工序站位判定（External Vendor / Sub Assy / F3）
  - PART SET NO `-**` 完整化（通过 YPL 列 AX/AY/AZ + CIT/CIT Matrix）
  - VMI 子装件零件号生成（末尾加 `-S1`，结合颜色码如 `-P0S1`）
- **输入**：用户上传一套完整源数据（PCL / YPL / CIT / QCPS / LN BOM / CU_FA CODE / Local supplier email）
- **输出**：与 `TB-25-137-F3` 同版式的 `.xlsx`（保留抬头、章节分组、签字栏占位）

### 2.2 Out of Scope（建议本期不做）

- TB 模板 **6（Supplier Sub-Part）、7（Steering/Wheel Module）、8（Colouring）、9（Graphic & Emblem）、10（Obsolete）、11（Frame/Engine No.）、12（Comparison）、13（Running Change）**
- 示例 TB 中的章节 **E–K**（F1↔F2↔F3↔Koike↔YIMM↔YMVN 互供，缺规则）
- **TB 编号自动生成**（如 `TB-25-137`）与版本管理（REV. 0/1/2…）
- **审批流**（Finance / SCM / QA / Metal Sys 多角色签字与流转）
- **数据源在线集成**（SharePoint / LN System 直连）
- **PDF 出版与分发**（邮件、公告板）
- 多语言、移动端
- **图案 & 颜色相关数据**（C&G GRAPHIC / C&G COLOUR 表本期为空）

---

## 3. 输入与输出（建议默认）

### 3.1 输入

| 文件/表 | 来源 sheet | 用途 |
|---|---|---|
| PCL | `1. PCL` (158 列) | 主零件清单 → 抽 SEC/LVL/PART_NO/PART_NAME/QTY/CU_FA/PART_SET_NO |
| YPL | `2. YPL` | YMAC 零件 → PART SET NO 完整化、颜色码（列 AY/AZ） |
| QCPS-Sub Assy / ML(R) / ML(L) | `3. / 4. / 5.` | USE IN 判定（External Vendor / Sub Assy / F3） |
| CU_FA CODE | `6. CU_FA CODE` | 工厂代码 → SUPPLIER 名称映射 |
| LN BOM | `7. LN BOM` (~23k 行) | EX / NEW 判定 |
| CIT Type A/B/C | `8. CIT - TYPE A/B/C` | 颜色码 → 颜色名/Mark |
| CIT Matrix | `11. CIT Matrix` | 多色型横向对照 |
| EMAIL - Local supplier info | `EMAIL - Local supplier info` | 外协供应商查表（来自 SCM PIC 邮件） |

### 3.2 输出

- 一份 `.xlsx`，结构与 `TB-25-137-F3` 一致：
  - 抬头：`HLYM TECHNICAL BULLETIN`、TB NO.、REF. NO.、NATURE、PURPOSE、WITH EFFECTIVE FROM、CONTENTS 目录
  - 章节 A–D（每章节含字段：`No. | SEC | LVL1–LVL4 | PART NO. | PART NAME | QTY/BIKE | SUPPLIER | USE IN | EX/NEW | REMARKS`）
  - NOTES 区
  - REV. 历史表（POC 默认 `[▲0] NEW ISSUANCE`）
  - 签字栏占位（Finance / SCM / QA / Metal Sys，仅显示文字、不签电子签）

---

## 4. 待确认事项

> **优先级说明**：P0 = 阻塞性（不解决无法开工）；P1 = 重要（影响 POC 完成度）；P2 = 一般（可在交付后讨论）

### 4.1 P0 — 必须在开工前确认

| # | 事项 | 影响 | 备选方向 |
|---|---|---|---|
| **Q1** | POC 是否限定 **TB 模板 1–5 / 章节 A–D**？还是要覆盖更多模板/章节？ | 决定开发量（从 1 周量级到 1 个月+） | a) 限 1–5 / A–D；b) 含 6–11；c) 全 13 类 |
| **Q2** | TB 章节 **E–K**（F1/F2/F3/Koike/YIMM/YMVN 互供）的判定规则在哪里？现有资料无 SOP | 若纳入范围，需 Yamaha 补 SOP 或样例 | a) 本期不做；b) 补 SOP 后纳入 |
| **Q3** | **数据输入方式**：每次用户手动上传 7–8 个 Excel？还是要接 SharePoint / LN System？ | 决定架构（前端 + 本地处理 vs. 后端 + 系统集成） | a) Excel 上传 POC；b) SharePoint 拉取；c) LN 系统直连 |
| **Q4** | **输出形态**：仅 `.xlsx`？还是要同时出 PDF？版式是否必须与示例 1:1（合并单元格、列宽、字体）？ | 决定渲染复杂度 | a) `.xlsx` 数据准确即可；b) `.xlsx` + 1:1 版式；c) `.xlsx` + PDF |
| **Q5** | **POC 平台形态**：命令行脚本、桌面工具（如本地 Web）、还是企业内网 Web 应用？ | 决定技术栈与部署 | a) Python CLI；b) 本地 Web（FastAPI/Streamlit）；c) 内网 Web |

### 4.2 P1 — 影响 POC 完整度

| # | 事项 | 当前现状 | 期望产出 |
|---|---|---|---|
| **Q6** | **空表数据**：`9. C&G - GRAPHIC`、`10. C&G - COLOUR`、`EMAIL - Import supplier info` 当前为空 | 影响图案/进口供应商场景演示 | Yamaha 补样例数据 |
| **Q7** | **业务规则文档化**：现有规则全部嵌在 SOP 文字里（如 PART SET NO 的"第 6–9 位 vs 第 5–8 位"匹配），需反推 | 易遗漏/误解 | 整理 Markdown 规则清单后请 Yamaha 评审 |
| **Q8** | **字段映射的边界**：PCL 共 158 列，YPL 共 88 列，是否所有列都可能用到？或仅 SOP 提到的 J/N/O/P/Q/W/Z/AC/AD/AE/AF/EI 等？ | 决定数据预处理范围 | 确认列清单 + 各列业务含义 |
| **Q9** | **EX/NEW 判定**：SOP 提到"若 LN BOM 不存在还需要 LN System Where-Used 复核"，POC 是否做这一步？ | 影响判定准确率 | a) 仅查 LN BOM 离线表；b) 集成 LN System（需接口） |
| **Q10** | **TB 编号生成规则** `TB-25-137`：是否在 POC 内自动分配？还是用户手填？ | 涉及编号台账管理 | 默认用户手填，自动方案需 Yamaha 提供编号台账 |
| **Q11** | **REV. 历史与差异比对**：POC 是否支持基于上一版 TB 自动生成 `REV.1 / REV.2` 与差异说明？ | 关系到模板 12（Comparison）能力 | 本期建议 Out of Scope |
| **Q12** | **签字栏**：仅显示固定文字（KCL / VICTOR / YMK 等姓名），还是要做电子签 / DocuSign 集成？ | 决定是否触发审批流模块 | 建议本期固定文字 |
| **Q13** | **多色型展开行为**：YMAC/LCP 中一个 PART SET NO 会展开为多个具体色号行（如 `-PN` / `-P7`），POC 输出是否要保留这种"一对多"行结构？ | 已在 SOP 中明确，建议保留 | 默认保留，按 CIT/CIT Matrix 全展开 |

### 4.3 P2 — 可在 POC 验收后讨论

| # | 事项 |
|---|---|
| **Q14** | 是否需要校验/审计日志（谁在何时生成了哪份 TB） |
| **Q15** | 多用户协作 / 权限分级 |
| **Q16** | 与 SCM 系统对接，自动从 LOO 拉取供应商更新 |
| **Q17** | 国际化（中/英/日切换） |
| **Q18** | 性能要求（PCL 千行/万行规模下的生成耗时上限） |
| **Q19** | 浏览器/操作系统兼容矩阵 |
| **Q20** | 与企业 SSO / AD 集成 |

---

## 5. 资料缺口清单

| 缺口 | 影响 | 需 Yamaha 提供 |
|---|---|---|
| 空表 `9. C&G - GRAPHIC` / `10. C&G - COLOUR` | 模板 8/9 无法做 | C&G GRAPHIC、C&G COLOUR 样例数据 |
| 空表 `EMAIL - Import supplier info` | 进口供应商无法测试 | 至少 1 个进口供应商邮件样例 |
| 章节 E–K（工厂互供）规则 | 章节 E–K 无法生成 | 类似 YMC/YMAC/LCP 的 SOP 文档 |
| 模板 6–13 的 SOP 与样例 | 仅模板 1–5 可做 | 每种模板 1 份样例 + SOP |
| `New Model Flow.xlsx` 中链接的参考文档（PCL/YPL/CIT/QCPS/VIN/Metal Sys/Progress Meeting）真实文件 | 无法验证全列字段语义 | SharePoint 拉取或文件下载 |
| TB 编号台账 | 自动编号无法做 | 历史 TB 编号、命名规则、年度复位规则 |
| 审批流程图 | 不能模拟流转 | Finance/SCM/QA/Metal Sys 审批顺序与停留条件 |
| 异常/边界规则 | 易出错 | 例：找不到供应商、CU_FA 不在 mapping、PART SET NO `-**` 无法在 YPL 找到时如何处理 |

---

## 6. 假设条件（在需求方反驳前默认成立）

- **A1**：源数据 Excel 的 sheet 名称、列结构与样例 `[POC] TB Structure Explanation.xlsx` 一致（PCL 158 列、YPL 88 列、CIT 22 列…）
- **A2**：用户每次会一次性提供**全部**源数据文件，不存在"部分缺失再补"的流程
- **A3**：单次生成只针对**一个新车型**（如 DA21 / T155-A,MYS,'26），不跨型号
- **A4**：用户具备阅读 SOP 的业务背景，POC 不承担"零业务知识"用户的教育职能
- **A5**：色号、零件号编码规则在 POC 期间保持稳定（不会出现新格式）
- **A6**：LN BOM 离线导出可代表系统当前状态，POC 不要求实时同步
- **A7**：所有数据为脱敏/公开数据，POC 阶段无信息安全合规要求

---

## 7. 建议里程碑

> 假设确认 P0 后采纳建议范围（模板 1–5，章节 A–D，Excel 上传 + Excel 输出）

| 阶段 | 周期 | 关键交付物 |
|---|---|---|
| **M0** 需求对齐 | 0.5 周 | 本文档 v1.0 (已签认)；规则清单 v1.0 |
| **M1** 数据层 | 1 周 | 7 类源数据 Excel 解析器 + 数据字典；单元测试 |
| **M2** 规则引擎 | 1.5 周 | YMC / YMAC / LCP 三套生成规则；端到端跑通示例 TB-25-137 |
| **M3** 输出渲染 | 1 周 | 与示例 1:1 版式的 Excel 输出 |
| **M4** UI + 联调 | 1 周 | Web/CLI 界面、上传 → 下载链路、对比示例验收 |
| **M5** 验收与文档 | 0.5 周 | 用户手册、Yamaha UAT、缺陷修复 |

**合计**：~5.5 周（不含 M0 前的需求等待时间）

---

## 8. 风险

| 风险 | 等级 | 缓解 |
|---|---|---|
| Yamaha 反馈周期长，P0 问题无法快速确认 | 高 | 单独安排 1h kick-off 会议集中过完 Q1–Q5 |
| 业务规则反推有误 | 中 | M1 末交付规则清单文档供 Yamaha 评审；M2 用示例 TB diff 校验 |
| 1:1 版式（合并单元格、签字栏布局）耗时超预期 | 中 | 使用 `openpyxl` + 模板克隆策略；不重新画版式，只填数据 |
| 源数据格式与样例不一致（实际生产数据列错位） | 中 | 加入"列名 + 列位置 + 列类型"三重校验 |
| LN BOM 23k 行性能瓶颈 | 低 | 用 pandas / dict 索引，预期 < 1s |
| 章节 E–K 后续被追加进 POC | 中 | 在合约/范围文档中明确"超出 A–D 视为变更" |

---

## 9. 下一步建议

1. **本周内**：把本文档发给 NMC 项目负责人，约一次 60 分钟会议过完 Q1–Q5（P0）
2. **会后 1 天内**：根据反馈出 v1.0 范围确认版
3. **同步进行**：让 Yamaha 补齐资料缺口（第 5 节清单）
4. **M0 完成**：即可启动 M1 数据层开发

---

*本文档由 Cascade 基于 `NMC Project/` 现有资料整理，所有"建议范围"为开工前的默认值，最终以 Yamaha 需求方书面确认为准。*
