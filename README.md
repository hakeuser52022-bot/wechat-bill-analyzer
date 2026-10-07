# 微信账单分析器

读取微信支付 Excel 账单，输出分类净支出、分段统计和餐费预算。开发使用了 AI 辅助；个人参与情况请在报名简介中据实填写。

## 快速运行

需要 Python 3.10 或更新版本。在项目目录执行：

```sh
python -m pip install -r requirements.txt
python tools/make_demo_data.py
python tools/run_demo.py
python tests/test_core.py
python tools/preflight.py
```

演示使用合成数据，不需要真实账单。完整结果见 [演示输出](docs/demo.md)。
准备报名时另运行 `python tools/preflight.py --submission`，未填完个人材料会返回失败。

## 功能与入口

| 入口 | 功能 |
|---|---|
| src/analyze.py | 分类、净支出、退款和收入，支持 `--out` 输出 Excel |
| src/timeseg.py | 按 JSON 中的日期区间分段，支持 `--out` |
| src/meals.py | 餐饮统计和预算，支持 `--days`、`--json` |
| src/budget.py | 经常性支出与节省方案，支持 `--month-days`、`--entertainment`、`--json` |

例如：

```sh
python src/analyze.py examples/data/demo_bill.xlsx --out report.xlsx
python src/analyze.py examples/data/demo_bill.xlsx --rules examples/rules.example.json --no-report
python src/meals.py examples/data/demo_bill.xlsx --from 2024-04-16 --to 2024-05-01 --days 31
```

## 统计口径

- 所有区间包含开始日期、不包含结束日期。指定完整区间时，日均包含没有消费的日期；未指定的边界从账单记录推断，可能不代表完整月份。
- 部分退款按已退金额扣除。例如原付 39 元、退回 20 元，净支出为 19 元。无法确定退款金额时停止并报告错误。
- `income` 排除交易类型含“退款”的回流；`balance` 为该收入减净支出。`cash_balance` 为本期所有收入流水减支出流水。跨期退款会使两种数值不同，二者都不是账户余额。
- 原交易金额保留在明细中；分类与预算使用扣除退款后的金额副本。
- 外卖正餐预算明确假设替代午餐，先扣除对应食堂午餐，再计外卖。奶茶和加餐是额外消费。
- 历史月化采用 30 天；计划使用指定天数，餐费现状对照也按计划天数展示。
- 某餐或可选项目没有样本时，不提供该档完整预算。JSON 中 `available=false`，total 和 daily 为 null；known_subtotal 仅表示已知项目小计。
- 节省额按餐饮、非餐饮分开，非餐饮预算不扣餐饮节省额。娱乐是在历史支出之外的额外额度，已有娱乐消费需人工核对，避免重复预留。

## 限制

仅支持包含必要字段的 xlsx，支持 Excel 日期和 ISO 文本日期。交易金额必须非负。分类是关键词匹配，超市、综合商户与院校付款仍可能需要人工校正。预付充值仅提示，不自动匹配后续刷卡消费。

退款按导出时原订单状态回溯到消费日；本工具不做跨账单交易匹配。退款回流依赖交易类型标记，数据源字段语义不同必须先核对。餐费锚点是每笔交易均值，多笔付款组成一餐时可能低估。订阅、教材等是否经常发生需按个人情况调整，预算不是必然发生的金额。

## 材料

- [项目讲解](docs/presentation.md)
- [修订与验证记录](docs/dev-log.md)
- [数据口径](docs/data-format.md)
- [个人简介](docs/intro.md)
- [提交清单](docs/submission-checklist.md)
- [GitHub 指南](docs/github-guide.md)

公开前人工检查完整文件内容。自动扫描不能证明不存在隐私信息；报名简介与公开仓库应分开整理。
依赖为 openpyxl、XlsxWriter 和 python-docx；项目许可证为 MIT，作者姓名仍待填写。

## 演示数据更新

本版演示账单由独立编写的生成器构造，使用 2024 年 4 月的虚构日期、示例商户及 DEMO 前缀交易号。没有读取个人账单；不再沿用原演示中的学校、商店、旅行或入学日程。Excel、JSON 和演示文档均已重新生成。
