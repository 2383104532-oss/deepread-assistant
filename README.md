# DeepRead Assistant · 组会论文精读助手

一个运行在 **DSH 工作台（dsh-worktable）** 里的「组会论文精读助手」自定义窗口。纯前端 HTML + JS，不依赖任何后端；论文摘要由你选择的 LLM（GLM / DeepSeek / GPT 或任意 OpenAI 兼容模型）生成。

## 功能

- **页面一 · 精读**
  - 加载项目里任意论文的 PDF（输入文件夹名 + PDF 文件名，或本地上传并自动保存进项目同名文件夹）。
  - 自动生成「讲故事」版本的结构化摘要（旧认知 → 新痛点 → 新贡献 + 对应原文与解析），用 marked 渲染成 Markdown，可复制。
  - 支持多种摘要模型：智谱 GLM（多模态看图）、DeepSeek（文本）、GPT/OpenAI（可多模态看图），或任意 OpenAI 兼容端点（自定义 Base URL + Key + 模型名）。
- **页面二 · 询问 + 生笔记**
  - 第一部曲·澄清：基于摘要由 AI 一次一问、逐轮追问，直到理解你的阅读目标。
  - 第二部曲·目标锁定：生成「目标复述 + 要点清单」，锁定后一键生成正式精读笔记（论文画像 / PSR / 贡献拆解 / 实验怎么做 / 关键证据 / 我的思考 / 与我研究关联 / 质量评估）。
- **页面三 · 预演**
  - 加载该论文的 PPT，自动转换（通过 DSH 工作台终端调用本机 PowerPoint/python-pptx）成视图 + 每页备注发言稿，左右并排、16:9 可缩放、上下页/滚轮切换。
  - 保留计时器。

## 技术要点

- 纯前端 `deepread.html`，使用 dsh-worktable 提供的原生皮肤（`/api/worktable/template/dshell.css`）。
- 依赖工作台接口：`/api/worktable/site`、`/api/worktable/file`、`/api/worktable/write`、`/api/worktable/mkdir`、`/api/worktable/term`。
- LLM 调用走 **OpenAI 兼容的 `{base}/chat/completions`**（DeepSeek / 智谱 / OpenAI / 本地 Ollama 等通用）。
- API Key 由你在页面里填写，**只存在浏览器 localStorage**，不写入任何文件、不进本仓库。
- `make-ppt-view.ps1`：把 `.pptx` 转成 `<PPT名>.pdf`（PowerPoint COM）并提取 `<PPT名>.md`（python-pptx）。
- `save-upload.ps1`：把上传的 PDF（base64）解码成真实 PDF 存进项目同名文件夹。

## 使用

1. 在 DSH 工作台的某个项目里，把 `deepread.html` 与两个 `.ps1` 放到项目根目录。
2. 在工作台该项目的 `widget-result.json` 写入挂载清单（本仓库附了一份示例）：
   ```json
   {"window":"窗口1","path":"deepread.html","kind":"html"}
   ```
3. 打开该窗口，在「① 精读」页选择摘要模型并填入对应 API Key（仅存浏览器），点「保存配置」。
4. 加载论文 PDF（或上传），即可自动生成摘要；随后走页面二澄清→生笔记、页面三 PPT 预演。

> 注意：`make-ppt-view.ps1` / `save-upload.ps1` 依赖本机安装 **Microsoft PowerPoint** 与 **python + python-pptx**；运行「上传自动保存」与「PPT 自动转换」需要 DSH 工作台的终端（`/api/worktable/term`）可用。

## 隐私

- API Key 仅存本地浏览器 localStorage，不上传、不入库。
- 论文 PDF / 备注 / 摘要 / 阅读状态等属于你的个人研究数据，默认不会被本仓库包含（见 `.gitignore`）。

## License

MIT
