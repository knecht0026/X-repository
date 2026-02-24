# X 推文导出到 Google Docs（Chrome 插件）

这个插件可以把当前 X 页面中可见的推文导出到 **同一个** Google Docs 文档中，并按以下格式写入：

- 小标题：`yy/mm-【推文主题-由AI生成】`（样式为 Heading 2，便于目录管理）
- 正文：推文原文（样式为 Normal Text）

## 已满足你的 3 个需求

1. **单一文档保存**：首次导出时创建文档并保存 `docId`，后续继续写入同一文档。
2. **小标题格式固定**：每条推文自动生成 `yy/mm-【主题-由AI生成】`。
3. **正文非标题格式**：正文统一写为普通正文样式。

## 安装

1. 打开 Chrome 扩展页：`chrome://extensions`
2. 启用「开发者模式」
3. 选择「加载已解压的扩展程序」并选择本目录

## 必须配置（Google OAuth）

在 `manifest.json` 中把：

```json
"client_id": "YOUR_GOOGLE_OAUTH_CLIENT_ID.apps.googleusercontent.com"
```

替换为你自己的 OAuth Client ID（Google Cloud Console）。

并确保该 OAuth 客户端已启用 Google Docs API，对应 Scope：

- `https://www.googleapis.com/auth/documents`

## 使用步骤

1. 打开任意 X 页面（时间线、用户页或线程页）
2. 点击插件图标
3. 首次可填写：
   - 文档标题（仅首次创建生效）
   - 文档 ID（可留空自动创建）
   - AI 配置（可选）
4. 点击「导出当前页面推文」

## AI 主题生成说明

- 如果配置了 `API Endpoint + Model + API Key`，插件会调用你的 AI 接口生成简短主题。
- 如果没配置或调用失败，会自动使用本地规则生成主题兜底。

## 说明与限制

- 目前导出的是**当前页面可见**的推文（先滚动加载更多可导出更多）。
- 推文去重依据：`datetime + text`。
