# shared/contracts —— 子 skill 之间的交接契约

包根 `SKILL.md` §4 定义了 Issue / PlanItem 两个核心结构。这里是它们的**机器可读 schema**，
以及为什么必须存在。

## 为什么需要契约

`workflow/audit_all.py` 要把 7 个子 skill 各自的报告**合成一份**。
如果每个子 skill 自定输出格式，合成就要写 7 套解析逻辑 —— 那正是本项目要避免的。
所以：**子 skill 只产出符合 Issue schema 的列表，不产出自己的报告文件**；
报告由 `shared/lib/report.py` 统一渲染。

同理，`workflow/plan.py` 把 Issue 翻译成 PlanItem，`workflow/apply.py` 只认 PlanItem。
子 skill 的"改"半部实现某几个 `action`，不自己决定"要不要改"—— 要不要改由用户在 plan.yaml 里定。

## 三个文件

| 文件 | 作用 |
|---|---|
| `issue.schema.json` | 一条问题。所有审计型子 skill 的返回值元素 |
| `plan.schema.json` | 一条待执行操作 + 整份 plan.yaml 的结构 |
| `README.md` | 本文件 |

## 一条铁律

**Issue 的 `standard_source` 字段必须填**（`template` / `config` / `default`）。
它回答"你凭什么说这里不合规"。缺了它，报告就无法复核，也无法区分
"学校规定的" 和 "工具建议的" —— 这是本包最容易变成噪声的地方。

## 校验

```bash
python shared/contracts/validate.py _output/audit.json
```

（`validate.py` 待实现；在那之前 `workflow/audit_all.py` 会做基本的字段完整性检查。）
