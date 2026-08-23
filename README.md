# CET 四六级成绩查询

面向福清华侨中学的 CET 四、六级成绩查询站点。项目通过教育部考试院的官方查询接口批量获取成绩，并将已脱敏的结果发布为静态网页。

网站地址：[nuo-luo.github.io/cet-query](https://nuo-luo.github.io/cet-query/)

## 使用方式

学生可在网站中按姓名搜索，并按班级或成绩状态筛选。成绩数据更新后，GitHub Pages 会自动发布最新结果。

## 隐私说明

- 学生名单 Excel 和完整身份证号不会提交到公开仓库。
- `scores.json` 仅保存脱敏后的证件号（前 4 位与后 4 位）。
- 名单以两个 GitHub Actions Secret 保存：`ROSTER_XLSX_B64_PART_1` 和 `ROSTER_XLSX_B64_PART_2`。

## 自动查询

工作流位于 [`.github/workflows/scrape.yml`](.github/workflows/scrape.yml)，保留手动触发，并在北京时间 8 月 24 日执行三次：

- 06:00
- 06:01
- 06:05

同一时间只会运行一个查询任务；首次已完成查询后，排队任务会自动跳过，避免重复请求。

## 本地运行

1. 在 `account/` 目录放入 `福清华侨中学学生名单.xlsx`（该目录已被 Git 忽略）。
2. 安装依赖：`pip install requests openpyxl`
3. 运行：`python scraper.py`

查询完成后会生成 `scores.json`。请在提交前确认其中不存在完整身份证号。

## 目录说明

| 文件 | 说明 |
| --- | --- |
| `index.html` | 静态查询页面 |
| `scraper.py` | 批量查询与结果生成脚本 |
| `scores.json` | 已脱敏的公开查询结果 |
| `.github/workflows/scrape.yml` | GitHub Actions 定时查询工作流 |

