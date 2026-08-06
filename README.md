# 📦 面单识别 → Tracking 回填工具

上传 UPS/FedEx 面单图片，AI 自动识别 Tracking Number 和收件人，匹配到 ParcelOutbound Excel 并回填。

## 功能

- 📸 批量上传面单图片（PNG/JPG）
- 🤖 Claude Vision AI 自动识别 Tracking # 和收件人
- 🔗 按收件人姓名模糊匹配 Excel 记录
- ✅ 人工确认匹配结果后一键回填
- 📥 下载更新后的 Excel 文件

## 本地运行

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动应用
streamlit run app.py
```

浏览器会自动打开 `http://localhost:8501`

## 部署到 Streamlit Cloud（免费，推荐）

1. 把这个文件夹推到 GitHub 仓库
2. 打开 [share.streamlit.io](https://share.streamlit.io)
3. 点 "New app" → 选你的 repo → Main file 填 `app.py`
4. 部署完成后会得到一个链接，同事直接访问即可

> API Key 每个用户自己在页面左侧输入，不会存储在服务器上。

## 部署到其他平台

### Railway / Render

```bash
# Procfile
web: streamlit run app.py --server.port=$PORT --server.address=0.0.0.0
```

### Docker

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY app.py .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0"]
```

## 使用方法

1. 在左侧栏输入 Anthropic API Key（从 console.anthropic.com 获取）
2. 上传面单图片 → 点击「开始识别」
3. 上传 ParcelOutbound Excel
4. 检查匹配结果，勾选/取消需要回填的条目
5. 点击下载更新后的 Excel
