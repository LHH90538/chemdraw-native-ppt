# ChemDraw Native PPT Skill

将化学文献图片拆解并重建为可编辑 PowerPoint：分子为 **ChemDraw 原生嵌入对象**，文字、箭头和简单图形为可编辑 PPT 对象。

A Codex skill for evidence-grounded reconstruction of chemical figures into editable PowerPoint, using native ChemDraw OLE for molecules.

## 能做什么

- 指导 Codex 读取原图、逐个重建分子并核对连接关系、粗键/楔形键、字体和上下标。
- 将准备好的 CDXML/CDX 导入多页 PPT，支持不同数量的分子。
- 输出独立项目及 `start.cmd`，显示分子逐个放置的过程。
- 检查保存后的 ChemDraw 对象数量并导出预览供人工或代理检查。

**这不是端到端 OCSR 模型。** 脚本不会直接识别图片；Codex 需要重建并核对分子。复杂或模糊化学结构不能保证零误差。照片/晶体插图缺少重建数据时须注明位图例外。

![Self-authored example](assets/synthetic-example/preview.png)

## 环境

- Windows，已安装并可用的 Microsoft PowerPoint 和 ChemDraw。
- Windows PowerShell 5.1，允许执行项目脚本；程序不修改执行策略或注册表。
- Python 3.10+（仅初始化项目，使用标准库）。
- 不内置 ChemDraw/Office，不提供其许可证；不需要新增识别 API 密钥。

## 安装个人 Codex Skill

将本仓库克隆到你的个人技能目录，以 `chemdraw-native-ppt` 为文件夹名。PowerShell 示例：

```powershell
$skillRoot = if ($env:CODEX_HOME) { Join-Path $env:CODEX_HOME 'skills' } else { Join-Path $env:USERPROFILE '.codex/skills' }
git clone https://github.com/LHH90538/chemdraw-native-ppt.git (Join-Path $skillRoot 'chemdraw-native-ppt')
```

已有同名目录时请先检查和备份，不要覆盖个人改动。安装后在新任务中附图调用：

> 用 $chemdraw-native-ppt 处理这张化学文献图，分子使用 ChemDraw 原生对象，核对结构、特殊键和排版，交付 PPT 与一键项目。

## 运行自制示例

在仓库目录下：

```powershell
python scripts/new_project.py --mode example --destination ../chemdraw-demo
powershell.exe -NoProfile -STA -File ../chemdraw-demo/build.ps1
```

也可双击新项目里的 `start.cmd`。默认将 PPT 保存到当前 Windows 用户的 Downloads；项目的 `work` 子目录保存日志和预览。建立新图项目用 `--mode empty`，然后由 Codex 完成源分子和 `project.json`。

## 文件和检查

- [SKILL.md](SKILL.md)：代理工作流程。
- [项目格式](references/project-format.md)：配置、坐标、缩放及输出说明。
- [重建规范](references/reconstruction.md)、[绘图样式](references/cdxml-style.md)、[排错](references/troubleshooting.md)。

本地测试覆盖：多分子与多页导入、中文项目路径、原生文字上下标与图形、错误特殊键检查和空项目拒绝。尚未跨 ChemDraw/PowerPoint 版本系统测试；对象结构检查不等同于化学正确性或交互编辑验证。每次交付仍需看预览并核对原图。

## 素材、隐私与许可证

代码、文档和仓库内自制乙醇/苯示例采用 [MIT License](LICENSE)。公开包不含原有文献截图、晶体图片、个人路径、账户凭据或工作日志。代码借助 AI 编写，经本机测试整理。

用户后续导入的文献图片和第三方材料不因本仓库许可证而获得授权。软件商标及第三方软件许可不包含在 MIT 授权中。本项目与 ChemDraw/Revvity、Microsoft、OpenAI 无官方关联。

反馈问题时请提供软件版本和去除个人路径、机密结构及凭据后的错误日志；不要上传秘密或未获授权的材料。
