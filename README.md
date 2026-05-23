# 山西财经大学本科学年论文格式自动化工具

基于《山西财经大学本科学年论文内容与格式规范》，自动格式化 .docx 学年论文的 Claude Code Skill。

## 功能

- 自动设置 A4 页面、页边距（上3cm/下2.5cm/左2.5cm/右2cm）
- 自动添加页眉（"山西财经大学20XX级本科生学年论文"）
- 自动格式化正文（小四号宋体、1.25倍行距、首行缩进2字符）
- 自动识别并格式化各级标题（一级三号黑体居中、二级四号宋体加粗、三级小四号宋体加粗）
- 格式化参考文献（小四号宋体、悬挂缩进）

## 安装

将本仓库克隆到 `~/.claude/skills/sxcd-paper-formatter/`：

```bash
git clone https://github.com/<your-username>/sxcd-paper-formatter.git ~/.claude/skills/sxcd-paper-formatter/
```

## 使用

在 Claude Code 中说：

> "帮我把这篇论文格式化成山西财经大学的规范格式"

或直接运行 Python 脚本：

```bash
python format_paper.py 论文主体.docx 2024
```

## 文件结构

```
sxcd-paper-formatter/
├── SKILL.md                    # 技能主入口
├── README.md                   # 项目说明
├── format_paper.py             # 格式化Python脚本
└── references/
    └── format-rules.md         # 格式规则详解
```

## 规范依据

- 附件1：山西财经大学本科学年论文内容与格式规范
- 附件4：山西财经大学本科学年论文模板

## 许可证

MIT License
