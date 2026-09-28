# 江西科技学院本科毕业论文 Typst 模板

GitHub 仓库：https://github.com/xgcr66/jxit-thesis-typst （仓库拼写为 jxit；本机目录为 jxut-thesis-typst）。

依据用户提供的《本科生毕业论文（设计）撰写基本规范（2022）》制作的**非官方模板**。使用一份结构化 Typst 内容，导出 PDF 和可编辑 Word。封面采用用户提供的江西科技学院校徽。

> 本仓库是排版模板，示例正文为填写提示。提交前以所在学院最新通知为准。请在自己的本地副本填写个人信息，公开分享时移除私人资料。

## 下载样稿

- [PDF 样稿](examples/v1_本科毕业论文模板.pdf)
- [可编辑 Word 样稿](examples/v1_本科毕业论文模板.docx)
- [格式依据与处理说明](docs/格式依据.md)
- [双格式支持范围](docs/导出说明.md)

## 本机使用

1. 解压或克隆到任意目录，例如 `D:\jxut-thesis-typst`。
2. 用 VS Code 或文本编辑器打开 `thesis.typ`，修改封面信息、摘要、关键词和正文。
3. 首次使用需安装 **Python 3.10+**，然后在项目目录的 PowerShell 中运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup.ps1
```

安装脚本在本目录创建 `.venv`，安装 `python-docx==1.2.0`，下载并校验 Typst 0.15.1 Windows x64 编译器。已经安装的 Typst 也可通过 `TYPST_BIN` 指定。

4. 双击 **build.cmd**，即可生成 PDF 与 Word；安装了 Microsoft Word 时会在后台更新目录，随后打开 `build` 文件夹。

```powershell
# 同时生成两种格式；使用 Microsoft Word 更新目录
.\build.ps1 -UpdateFields
# 只生成 PDF
.\build.ps1 -Format pdf
# 生成带 v2 前缀的新版本
.\build.ps1 -Version 2 -UpdateFields
```

输出文件：

```text
build/v1_本科毕业论文模板.pdf
build/v1_本科毕业论文模板.docx
```

没有 Microsoft Word 也可以生成 DOCX；打开后在目录区域右键选择“更新域 → 更新整个目录”。Word 与 Typst 是不同排版引擎，分页可能不同，各自目录对应各自页码。

本次交付的 D 盘目录已配置可用的本机工具，直接双击 `build.cmd` 即可；`local-tools.json` 属于本机配置，已忽略，不上传。

## 如何写内容

```typst
#import "lib/model.typ": *

// thesis.typ 中 chapters 的一项
chapter("系统设计", (
  h2("总体设计"),
  h3("模块划分"),
  p("在这里写正文。普通字符串同时用于 PDF 和 Word。"),
  rich(("引用文献", cite(1), "后的论述。")),
  fig("系统架构图", "assets/my-figure.png", width: 14.0),
  tab("数据表示例", ("字段名", "数据类型", "含义", "备注"), (
    ("id", "INTEGER", "编号", "主键"),
  )),
))
```

- 引用只使用 `references` 中已核验的条目，按正文首次出现顺序手动维护文献列表。
- 图和表在每章内各自自动连续编号，调整内容后重新导出即可统一更新编号。
- 图片采用 PNG/JPEG，以兼容两种格式；矢量源文件可另存。请在段落中先引用图表。
- 添加章节、标题、表格、图片后，重新运行导出；Word 中手工修改的内容不会回写 Typst。
- `appendix: ()`、`acknowledgements: ()` 可隐藏对应可选部分。
- 每个段落使用 `p("……")`；请勿直接粘贴任意 Typst 宏或数学公式作为字符串并期待自动转换。
- PDF 也可以直接运行 `.\.tools\typst.exe compile main.typ build\preview.pdf`，但统一脚本会额外校验图片、引用和字体。

## 格式规则

A4 纵向；横向页边距 3 cm；正文宋体小四，西文 Times New Roman；20 pt 行距；首行缩进 2 个汉字。章标题黑体三号居中，二级黑体小四，三级宋体小四。页眉、页脚宋体五号。摘要、目录、正文分别重新计页，正文后参考文献、附录与致谢连续计页。

**页眉页脚冲突的处理已由使用者确认：**规范文字要求页眉、页脚距离各 2.8 cm，而原文件实际设置约为 1.5/1.75 cm。此模板采用文字要求 2.8 cm，并额外预留 0.7 cm 避让区，因此正文有效上下边界为 3.5 cm；基本页边距参数仍记录 2.54 cm。该处理并非对学校未明确规则的官方解释，细节见格式依据。

Windows 字体需具备 SimSun（宋体）、SimHei（黑体）、FangSong（仿宋）、Times New Roman。脚本检测到缺失会停止并给出提示，不悄悄替换字体。其他系统需自行合法安装这些字体；若改用其他字体，须重新检查版式。

## 文件结构

```text
main.typ                   PDF 入口与内容导出元数据
thesis.typ                 唯一正文来源、封面与摘要信息
layout.json                字体、页边距、页眉页脚配置
lib/model.typ              可导出的段落、标题、图表语义块
lib/layout.typ             Typst 排版规则
assets/                    校徽及模板自有示例图
scripts/build.py            校验并导出两种格式
scripts/export_docx.py      可编辑 Word 生成器
scripts/update-word-fields.ps1  Word 目录更新
setup.ps1 / build.ps1 / build.cmd
examples/                  已生成并检查的示例 PDF / DOCX
tests/                     格式结构与错误分支测试
```

## 测试

先导出一次，再执行：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

测试检查源数据、标题层级、图片与引用验证、Word 可编辑表格、目录域、分节页码及规范样式。实际渲染检查记录见 [验收记录](docs/验收记录.md)。

## 使用边界与授权

这不是任意 `.typ` 文件到 Word 的通用转换器。Word 由经过校验的同源内容构建，标题、正文、表格、图题及引用为可编辑文本；图片作为图片嵌入。自定义布局、复杂公式、脚注等暂未提供双格式语义块，扩展时须同时完善两种渲染器。

代码采用 MIT License。校徽是用户提供的学校标识，相关权利归权利人所有，不纳入代码的 MIT 授权；仅用于本模板的学校论文封面。仓库未附带学校规范原文件、商用字体、个人论文或签名。声明文本沿用用户提供的规范，签名与日期由作者本人填写。
