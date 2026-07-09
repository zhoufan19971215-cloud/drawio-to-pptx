# Drawio to PPTX

在线将 .drawio 文件转换为可编辑的 PowerPoint 演示文稿。

## 功能

- 上传 .drawio / .dio / .xml 文件
- 实时 Canvas 预览
- 一键导出为 .pptx（使用 python-pptx 生成原生可编辑 PPT）

## 技术栈

- **前端**: 原生 HTML/CSS/JS + Canvas 预览
- **后端**: Flask + python-pptx
- **部署**: Render

## 本地运行

```bash
pip install -r requirements.txt
python server.py
```

访问 http://127.0.0.1:8765

## 部署到 Render

1. 将项目推送到 GitHub
2. 登录 https://dashboard.render.com
3. 点击 **New** -> **Web Service**
4. 连接你的 GitHub 仓库
5. Render 会自动识别 `render.yaml` 配置并部署
6. 部署完成后获得公网地址，无需代理即可访问
