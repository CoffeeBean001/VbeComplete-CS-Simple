# VbeComplete

VBA 编辑器（VBE）的代码补全与编辑增强工具。常驻托盘，靠全局键盘钩子 + COM 后期绑定
跟 VBE 打交道 —— 不安装插件、不往工程里加引用、不改你的 VBA 工程，双击一个 exe 就能用。

```
┌─ VBE 代码窗 ────────────────────────────┐
│ Sub 生成工资表()                        │
│     Dim ruleSht As Worksheet, w As W    │  ← 候选窗跟着光标走
│         ┌──────────────────────────┐    │
│         │ Worksheets               │    │
│         │ w                        │    │
│         └──────────────────────────┘    │
└─────────────────────────────────────────┘
```

## 补什么

候选池里**只有你自己写的东西**：过程名、变量名、参数名、常量名、枚举成员。

VBA 关键字、内建函数、内建类型、模块名 / 表名（`Sheet1`、`ThisWorkbook`）、窗体名、
类名、自定义类型名，一律不弹 —— 少掉这几百条噪音，剩下的每一条都能直接按 Tab 用掉。

作用域按 VBE 的口径算：当前过程的局部变量排最前，其次是模块级，然后是工程级；
别人过程里的 `Private` 变量不会跑到你眼前。`Option Explicit` 模块里没声明过的名字不会
被当成隐式变量收进来。

## 快捷键

| 按键 | 作用 |
| --- | --- |
| 打字 | 自动出候选窗；`↑`/`↓` 选择，`Tab` 确认，`Esc` 收起 |
| `Ctrl+Space` | 主动唤出候选（注释 / 字符串里不弹） |
| `Enter` | 回车自动缩进、自动补 `End Sub` / `End If` / `Next` 等收尾 |
| `Ctrl+/` | 批量注释 / 取消注释（支持多行选区） |
| `Ctrl+D` / `Ctrl+K` | 复制当前行到下方 / 删除当前行（有选区 = 整个选区） |
| `Ctrl+Shift+Enter` | 在当前行上方插入空行 |
| `Ctrl+↑` / `Ctrl+↓` | 跳到本过程最后一个 `Dim` 行 / 最后一条代码 |
| `Ctrl+Q` | 代码片段库（`Ctrl+Enter` 插入到光标处） |
| `Backspace` | 一次删掉整行空白、成对的空括号 / 空引号 |

鼠标单击、粘贴、拖动滚动条不会误触发候选窗：只有真的在敲字才弹。

## 构建

需要 .NET 10 SDK（Windows x64）。

```bat
build.bat
```

等价于：

```bat
cd src
dotnet publish -c Release -r win-x64 --self-contained false ^
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o ..\dist
```

产物是 `dist\VbeComplete.exe`（单文件、约 590 KB）。它只依赖机器上已装的 .NET 运行时，
不打包运行时本体。

## 运行

双击 `dist\VbeComplete.exe`。**没有主窗口** —— 图标在右下角托盘里：

- 双击托盘图标：把图标提示出来 / 再弹一次提示气泡
- 右键托盘图标：开关「开机自动启动」、退出

`run-cs.bat` / `stop-cs.bat` 用于开发时启停；`build.bat` 编译完会自动跑一遍自检。

## 配置

`settings.example.ini` 是模板，复制成 `settings.ini` 再改（改完重启生效）。
优先级：环境变量（`VBECOMPLETE_` 前缀）> `settings.ini` > 代码里的默认值。
每一项都在模板里写清楚了作用，认不出的值会被跳过，不会把功能写坏。

## 自检

```bat
check-cs.bat
```

九步，从纯逻辑到真机：

| 开关 | 内容 | 需要 VBE 吗 |
| --- | --- | --- |
| `--selftest` | 候选池、模糊匹配、作用域、编辑判据、配置、片段库（460+ 条断言） | 否 |
| `--snipetest` | 片段库增删改查（临时库，不碰你的数据） | 否 |
| `--uitest` | 候选窗 / 片段窗离屏渲染成 PNG | 否 |
| `--startuptest` | 内嵌图标、开机自启项的读写与还原 | 否 |
| `--enginetest` | 用真实模块文本跑候选规则 | 是 |
| `--paretest` | 补全后括号 / 空格的落点 | 是 |
| `--carettest` | 补全后光标位置 | 是 |
| `--hktest` | 热键命令（注释 / 删除行 / 跳转） | 是 |
| `--enterprobe` | 回车收尾（会真的驱动一次键盘钩子） | 是 |

另有 `--dump`（把当前模块全文写成 `moddump.txt`）、`--pooldump <模块名>`（看候选池）、
`--probe`（当前连的是哪本工作簿、光标在哪）、`--yieldprobe`（VBE 成员列表是否可见）。

## 目录结构

```
src/
  App/        托盘、开机自启、图标、主流程与键盘钩子
  Cli/        命令行入口 + 各种自检 / 诊断命令
  Core/       候选引擎、作用域分析、VBA 解析、模糊匹配、编辑手感
  Interop/    Win32 封装、COM 后期绑定、VBE 读写
  Storage/    配置、日志、片段库（SQLite）、最近使用
  Ui/         候选窗、片段窗
  Tests/      自检用例
  Resources/  图标与清单文件
```

## 实现要点

- **不碰 VBE 内部**：全部走 `VBE7.DLL` 的 COM 后期绑定（`IDispatch`），VBE 没装也能启动。
- **钩子线程一次 COM 都不调**：低级键盘钩子里只读内存状态，真动作投递给 UI 线程 ——
  钩子回调一旦被阻塞超过 `LowLevelHooksTimeout`，系统会把钩子摘掉。
- **候选池是推出来的，不是猜的**：模块文本 + 工程作用域 → 声明表 → 池子；池子只缓存，
  文本一变（含等长改名）就重建。
- **SQLite 直连系统 `winsqlite3.dll`**：不引 NuGet、不带原生二进制，片段库仍是标准
  `.db` 文件，可以拿 DB Browser 直接打开。
- **发布只有一个 exe**：图标内嵌在程序集里，托盘、窗口、任务栏共用同一份。

## 环境要求

- Windows 10 / 11 x64
- .NET 10 桌面运行时
- 只对 VBE（Excel / Word 等 Office 应用里的 VBA 编辑器）生效
