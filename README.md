# 立丰咨询业务智能体

基于 [GPT Academic](https://github.com/binary-husky/gpt_academic) 二次开发的咨询业务垂直领域智能体。项目使用 Gradio 提供 Web 界面，在通用大模型对话、文档解析和插件系统之上，增加了面向咨询项目的数据上传、历史数据对比、动因分析、成本标准汇总和报告生成等功能。

## 核心功能

- 多模型对话与统一 Gradio 操作界面
- PDF、Word、Excel、Markdown 等文档解析与总结
- 咨询业务数据上传和历史数据对比
- 多 Sheet 动因分析、成本标准与图表生成
- 知识库入口和 RAG 相关插件
- 可扩展的 `crazy_functions` 插件机制

## 目录

```text
main.py              Gradio 应用入口
config.py            公开默认配置（不应写入真实密钥）
core_functional.py   基础快捷功能
crazy_functional.py  插件注册与分组
crazy_functions/     通用及业务插件
request_llms/        模型接入层
shared_utils/        配置、日志、上传和 Web 工具
themes/              页面主题与前端资源
tests/               测试与验证脚本
```

## 安装

建议使用 Python 3.9–3.11。

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/macOS: source .venv/bin/activate
python -m pip install -r requirements.txt
```

项目对 Gradio 版本有明确要求，请使用 `requirements.txt` 中指定的依赖，不要直接升级 Gradio。

## 安全配置

不要在 `config.py`、`docker-compose.yml` 或任何已跟踪文件中写入真实 API Key。推荐在项目根目录新建已被 Git 忽略的 `config_private.py`：

```python
API_KEY = "your-api-key"
LLM_MODEL = "gpt-4o-mini"
WEB_PORT = 12345
```

配置优先级为：

```text
环境变量 > config_private.py > config.py
```

## 启动

```bash
python main.py
```

程序会输出实际监听地址。默认页面标题为“立丰咨询业务垂直领域智能体”。如需知识库按钮指向的外部服务，需另行在对应端口部署知识库应用。

Docker 启动：

```bash
docker compose up -d
```

使用 Docker 前，请使用 `.env` 或部署平台的密钥注入机制配置 API Key，并复核端口、挂载目录与访问控制。

## 数据与隐私

- 用户上传文件、日志、中间表格和生成报告不应提交到 Git。
- 输入数据可能包含客户或业务敏感信息，部署前应配置访问认证、数据保留期限和备份策略。
- 调用外部大模型时，需确认数据传输符合所在组织的合规要求。

## 上游与许可

本仓库源自 GPT Academic，并包含针对立丰咨询业务的定制修改。原项目、第三方组件及模型分别受各自许可条款约束；使用和再分发前请阅读本仓库 `LICENSE` 及相关组件的许可说明。
