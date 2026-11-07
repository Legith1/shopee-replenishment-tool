# Shopee 官方仓全流程智能补货与装箱系统 (Shopee Full-Stage Replenishment System)

专为 **Shopee 跨境官方仓（FBS/海外仓）** 卖家量身打造的全链路智能补货、申报与装箱自动化工具。系统覆盖从销量测算、审批申报到装箱打包的全业务闭环，支持多站点差异化规则（马来西亚 MY、泰国 TH、菲律宾 PH），具备均分连箱与同色拼箱自动识别能力。

---

## 🌟 核心业务阶段与功能特性

### 🚀 阶段一：补货数据智能测算与总表生成
- **多源数据自动对齐**：按 SKU 匹配马帮 30 天销量，按 仓库SKU ID 提取 Shopee 可售/在途库存及 RPA 销售速度。
- **站点差异化安全上限**：
  - 🇲🇾 **马来 (MY) / 🇹🇭 泰国 (TH)**：补货量上限限制为 **75 天销量**。
  - 🇵🇭 **菲律宾 (PH)**：补货量上限严格限制为 **60 天销量**（自动移除 75 天列，N 列公式设为 `60*M-G`）。
- **智能箱规算法与拦截规则**：优先整箱试算；超出上限时平滑测算最贴合值；遇停产/下架/超长备注自动拦截留空；未维护箱规严格留空。
- **100% 保留原生公式与条件格式**：红绿预警颜色与重复值高亮无缝生效。

### 📥 阶段二：Shopee 官方仓入库申请表自动生成
- 导入审批完毕的总表，自动通过 `Warehouse SKU ID` 映射。
- 将【采购批准】填入【* 申请入库数量】，将【出口货值】填入【* 单件出口申报货值】，**未批准商品严格留空**，可直接上传 Shopee 后台。

### 📦 阶段三：打包装箱分配与尺寸重量自动计算
- **连箱均分解析**：序号 `3-7` 箱数 `5` 自动拆解并均分至 `箱子 #3` ~ `箱子 #7` 各列。
- **同色拼箱混装识别**：同一序号下多行商品（或同色行）自动合并装入同一箱（如 `箱子 #19` 混装多款鱼线）。
- **箱规尺寸重量匹配**：单品箱自动从箱规库提取长/宽/高/重量并填入【尺寸和重量 Sheet】；拼箱/混装项自动留空等待手动复核。

### 🛡️ 阶段四：3 重全自动数据防呆验证网关
1. **SKU 字典对齐校验**：校验商品是否在底表/箱规库中存在。
2. **装箱闭环平衡校验**：自动核验每行所有分配箱子之和是否严格等于采购批准数量。
3. **单箱尺寸重量安全校验**：单箱重量 >30kg 或单边 >100cm 时自动预警。

---

## 📁 项目目录结构

```text
shopee_replenishment_tool/
├── core.py               # 核心多阶段业务计算与 Excel 处理引擎
├── gui.py                # 现代化桌面图形界面客户端 (Tkinter)
├── cli.py                # 命令行自动化流水线接口
├── run.bat               # Windows 一键双击启动脚本
├── requirements.txt      # Python 依赖清单 (openpyxl)
└── README.md             # 项目说明与 GitHub 发布文档
```

---

## 🚀 快速上手使用

### 1. 安装环境依赖
```bash
pip install -r requirements.txt
```

### 2. 启动图形化工作台 (推荐)
- **Windows 用户**：直接双击根目录下的 **`run.bat`**。
- **命令行启动**：
  ```bash
  python gui.py
  ```

### 3. 命令行流水线调用 (CLI)
```bash
# 执行阶段一：生成补货总表
python cli.py --stage 1 --site my --date 11.7

# 执行阶段二：生成入库申请表
python cli.py --stage 2 --site my --date 11.7

# 执行阶段三：生成打包与尺寸重量表
python cli.py --stage 3 --site my --date 11.7

# 一键执行全部阶段
python cli.py --stage all --site my --date 11.7
```

---

## 🛠️ GitHub 部署指南

```bash
cd /d "D:\Accio\work3\shopee_replenishment_tool"
git init
git add .
git commit -m "feat: shopee official warehouse full-stage replenishment system"
git branch -M main
git remote add origin https://github.com/<您的GitHub用户名>/<仓库名>.git
git push -u origin main
```
