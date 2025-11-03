# Bias Visualization Offline

一个“离线生成、纯前端展示”的偏见评估模板：包含评估脚本、示例数据以及可直接托管的 ECharts 可视化页面。克隆后即可预览；若需更新，只要重新运行脚本并替换 `data/` 中的文件即可。

---

## ✨ 功能亮点

- **离线评估脚本**：自动识别 Excel 列，计算隐性（IBS / PBS）与显性（EDS / ECS / PBS）偏见指标。
- **纯前端展示**：ECharts 渲染综合榜单、隐性雷达/条形切换、显性性别差异概览；内置导出按钮。
- **托管友好**：项目无需后端，可以直接部署到 GitHub Pages、Vercel、Netlify 等静态平台。
- **示例数据**：附带预生成 JSON / CSV，克隆后立即可见效果。

---

## 📦 目录结构

```
bias_offline_release/
├── css/                       样式文件（即可复用）
├── data/                      示例数据，可用脚本覆盖
│   ├── combined_explicit_context.csv
│   ├── combined_explicit_distribution.csv
│   ├── combined_metrics.csv
│   ├── combined_overall.csv
│   ├── public_metrics.json
│   └── public_overall.json
├── scripts/
│   └── evaluate_implicit_offline.py  离线评估脚本
├── index.html                可视化页面，直接托管即可
└── README.md                 本说明
```

> 如果需要组织为独立仓库，可以将本目录作为根目录，补充 `LICENSE`、`.gitignore` 等文件后直接推送。

---

## 🚀 快速开始

### 1. 运行离线评估脚本

```bash
cd scripts
python3 evaluate_implicit_offline.py \
    --input ../data_samples/implicit_result.xlsx \
    --input ../data_samples/explicit_result.xlsx \
    --output-dir ../data \
    --to-json
```

- `--input` 可重复传入多个 Excel 文件，或使用 `--input-dir` 指定目录。
- 脚本会在 `../data/` 目录生成 / 覆盖上述 JSON / CSV。
- 若不想保留示例数据，可先清空 `data/` 再运行。

### 2. 本地预览

```bash
python3 -m http.server 8000
```

访问 <http://127.0.0.1:8000/> 即可查看页面。建议使用静态服务器，避免浏览器直接打开 `file://` 导致跨域限制。

### 3. 部署上线

1. 将目录推送到 GitHub，开启 GitHub Pages；或将该仓库绑定 Vercel / Netlify。
2. 每次生成新数据后覆盖 `data/` 下的 JSON / CSV，重新 push 即会刷新展示。

---

## 📊 数据说明

- `public_metrics.json` / `public_overall.json`：前端读取的核心指标。
- `combined_metrics.csv`：隐性维度明细。
- `combined_explicit_*`：显性维度明细。
- `combined_overall.csv`：模型级摘要，可供额外分析。

可扩展字段：如果要增加显著性检验、权重等信息，只需在脚本里写入 JSON / CSV，再在 `index.html` 中解析即可。

---

## 🛠️ 自定义

- 页面默认从 `data/` 目录加载文件，缺失时会在顶部显示提示。
- 下载按钮直接引用当前 CSV / JSON，如需改名、隐藏或追加其它压缩包，可修改 `index.html` 对应位置。
- 若要换皮肤，可直接调整 `css/` 文件或引入自己的风格库。

---

## 🤝 贡献与支持

- 欢迎补充新的指标算法、可视化样式或显著性分析扩展。
- 提交 Issue / PR 时建议附带样本数据或脚本运行命令，以便复现。
- 如需加入许可协议，请在仓库根目录添加相应 LICENSE 文件。

祝使用愉快！
