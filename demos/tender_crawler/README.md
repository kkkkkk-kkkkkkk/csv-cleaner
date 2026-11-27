# 招投标公开信息采集器（演示版）

针对"采集公开招标/中标/预公示信息 → 清洗去重 → 结构化入库 → **导出甲方可直接打开用的 xlsx**"这类需求的可运行原型。

## 跑起来

```bash
python -m unittest test_crawler -v     # 9 个用例：链接抽取/字段清洗/金额换算/去重/入库/重试/失败入账/xlsx
python run_demo.py                     # 离线跑通全链路
python run_demo.py --retries 4 --xlsx out/tenders.xlsx   # 指定重试次数与导出路径
```

产物：`out/tenders.jsonl`（全量留档）、`out/tenders.sqlite3`（幂等主表）、`out/tenders.xlsx`（三页交付表）、`out/run.log`（运行日志）。

## 组成

- `crawler.py`
  - `extract_links()`：列表页抽详情链接，绝对化 + 页内去重，可按 URL 规则过滤（`/tender/`、`/notice/`、`/bid/`）
  - `parse_detail()`：字段规则表 `FIELD_RULES` 一处改全站生效（标题/采购人/金额/截止时间/地区/公告类型/发布时间）
  - 金额统一换算为元：`128.50万元 → 1285000`、`460,000元 → 460000`
  - `fingerprint()`：标题+采购人+预算标准化后 SHA1 前 16 位；同标的的"更正公告"会被判重（演示里 4 条详情只入库 3 条）
  - `fetch_with_retry()`：**指数退避重试**（默认 3 次），失败不抛异常，返回 `(html, ok, error)`
  - `crawl()`：列表→详情→清洗→去重→入库，返回 `(records, stats, failures)`；`stats` 含 `retries/errors/failed` 与 **`fill_rate`**（每个字段的填充率）
  - `store_jsonl()` / `store_sqlite()`：JSONL 全量留档 + SQLite 主键 `fp` 幂等写入
  - `store_xlsx()`：**三页交付表** —— `tenders`（主表，冻结首行 + 筛选 + 千分位 + 缺字段黄标）、`failed`（没抓到的链接 + 原因）、`fill_rate`（字段填充率，判断这站值不值得长期抓）
- `fixtures/`：4 个详情页 + 1 个列表页（含重复链接、无预算页、更正公告，专门验证去重与缺失计数）
- `run_demo.py`：离线跑，日志同时写 `out/run.log` 与控制台；支持 `--fixtures/--out/--retries/--xlsx`

## 实测结果（2026-11-27 重跑）

```
Ran 9 tests in 1.2s
OK
抓取统计: {'pages': 1, 'details': 4, 'dup': 1, 'missing': 1, 'retries': 0, 'errors': 0, 'failed': 0, 'elapsed': 0.01}
入库条数: 3 | 失败链接: 0 | xlsx: out/tenders.xlsx
字段填充率: title 1.0 / buyer 1.0 / deadline 1.0 / region 1.0 / notice_type 1.0 / publish_time 0.667 / budget_yuan 0.667
  - 公开招标   某市政务云平台运维服务项目     1285000  2026-12-08 09:30
  - 公开招标   县域医共体信息化建设项目       460000   2026-12-02 14:00
  - 竞争性磋商 某学校食堂食材配送项目         None     2026-11-30 17:00
```

**为什么值得看**：第 3 条预算为空时，旧版只在 jsonl 里留一个 `null`，甲方分不清"抓漏了"还是"本来没有"；新版把它标成 `6/7` 完整度并在黄格提示，同时 `fill_rate` 页给出全站字段填充率 —— **少什么，说清楚**。

## 接真实站点时怎么改

1. 把 `FETCH` 换成带 Cookie/UA/代理的下载函数（`requests` 或 `urllib`）；重试与 429/403 退避已内置，只需保证下载函数抛异常而不是返回空串；
2. 列表页若为动态渲染，用 Playwright 取 `page.content()` 再喂给 `parse_detail()`；
3. `FIELD_RULES` 按目标站点的字段名调整，金额单位写进 `BUDGET_RULE` 的分组；
4. 增量跑：SQLite 已有 `fp` 直接 `INSERT OR REPLACE`，重复跑不会产生脏数据；
5. `store_xlsx()` 的列可以按甲方模板改，行高/列宽/筛选都已预置。

合规边界：只采公开公示页面，不登录、不碰个人隐私字段，单站点 QPS ≤ 1。
