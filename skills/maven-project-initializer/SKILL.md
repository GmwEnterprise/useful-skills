---
name: maven-project-initializer
description: 在 vscode+opencode 环境中初始化 Maven Java 工程配置。当用户需要初始化 Maven Java 项目、配置 Java 开发环境时触发。
---

# Maven Project Initializer

在 vscode + opencode 环境中初始化 Maven Java 工程所需的配置。

## 何时使用

- 新建 Maven Java 项目需要配置开发环境
- 现有 Maven Java 项目缺少 vscode/opencode 配置

## 前置要求

- 已安装 JDK
- 已安装 VSCode + Java Extension Pack
- 已安装 opencode

## 执行步骤

### 第 1 步：确认 Maven 工程结构

检查当前项目是否为 Maven Java 工程结构。

**若不满足**：告知用户当前项目不是标准 Maven Java 工程结构，列出缺失项，询问是否继续。

**若满足**：进入第 2 步。

### 第 2 步：确认 JDK 安装路径

先自动检测当前平台：

```bash
# 检测平台
uname -s 2>/dev/null || echo "Windows"
```

然后按平台向用户确认 JDK 来源。**必须使用询问 + 提供选项的方式**，格式如下：

> 当前平台为 **{platform}**，请确认 JDK 的安装来源：
>
> **A)** sdkman — 路径格式：`~/.sdkman/candidates/java/<version>`
> **B)** asdf — 路径格式：`~/.asdf/installs/java/<version>`
> **C)** 直接提供 JDK 路径

根据用户选择的来源，执行对应的探测命令获取具体路径：

| 来源 | 探测命令 | 适用平台 |
|------|---------|---------|
| sdkman | `sdk current java` 获取版本，拼接 `~/.sdkman/candidates/java/<version>` | Linux, macOS |
| asdf | `asdf current java` 获取版本，拼接 `~/.asdf/installs/java/<version>` | Linux, macOS |
| Scoop | `scoop list` 查找已安装的 JDK，拼接 `~/scoop/apps/<jdk-package>/current` | Windows |
| 直接提供路径 | 使用用户提供的路径 | 全平台 |

> **Windows 平台额外选项**：
>
> **A)** 手动安装 — 常见路径如 `C:\Program Files\Java\jdk-<version>`
> **B)** Chocolatey — 路径格式：`C:\Program Files\Java\jdk-<version>`
> **C)** Scoop — 路径格式：`~/scoop/apps/temurin<version>/current`（或其他 scoop 安装的 JDK 包名）
> **D)** 直接提供 JDK 路径

获取到候选路径后，**必须验证路径有效性**：

```bash
# Linux/macOS
ls -d "<candidate_path>" && "<candidate_path>/bin/java" -version

# Windows
if exist "<candidate_path>\bin\java.exe" ("<candidate_path>\bin\java.exe" -version)
```

若验证失败，告知用户并提供其他选项重新选择。

确认有效路径后，**将路径中的 `~` 展开为绝对路径**（如 `/home/user` 或 `C:\Users\user`），进入第 3 步。

### 第 3 步：创建 .opencode 插件

创建 **`.opencode/plugins/inject-env.js`**（`.opencode/package.json` 由 opencode 自动管理，无需手动创建）：

将各平台的 `{{JAVA_HOME_<PLATFORM>}}` 替换为第 2 步确认的 JDK 完整路径。只填写当前平台对应的路径，其他平台保留为占位符或省略该分支：

```js
export const InjectEnvPlugin = async () => {
  return {
    "shell.env": async (input, output) => {
      switch (process.platform) {
        case "darwin":
          output.env.PROCESS_PLATFORM = "darwin"
          output.env.JAVA_HOME = "{{JAVA_HOME_DARWIN}}"
          break
        case "linux":
          output.env.PROCESS_PLATFORM = "linux"
          output.env.JAVA_HOME = "{{JAVA_HOME_LINUX}}"
          break
        case "win32":
          output.env.PROCESS_PLATFORM = "win32"
          output.env.JAVA_HOME = "{{JAVA_HOME_WIN32}}"
          break
        default:
          output.env.PROCESS_PLATFORM = "unknown: " + process.platform
      }
    },
  }
}
```

**注意**：
- Windows 路径使用正斜杠（如 `C:/Program Files/Java/jdk-21`）或双反斜杠
- 只生成分支，无需生成所有三个平台

### 第 4 步：创建 .vscode/settings.json

将 `{{JAVA_HOME}}` 替换为第 2 步确认的 JDK 路径：

```json
{
    "java.compile.nullAnalysis.mode": "automatic",
    "java.dependency.packagePresentation": "hierarchical",
    "java.configuration.updateBuildConfiguration": "automatic",
    "java.jdt.ls.java.home": "{{JAVA_HOME}}",
    "maven.executable.options": "-DskipTests=true"
}
```

**Windows 注意**：路径中的反斜杠需转义为双反斜杠（如 `C:\\Program Files\\Java\\jdk-21`），或使用正斜杠。

若 `.vscode/settings.json` 已存在，则合并配置项，保留用户已有设置。

### 第 5 步：验证

- 确认 `.opencode/plugins/inject-env.js` 已创建，`JAVA_HOME` 路径与第 2 步一致
- 确认 `.vscode/settings.json` 已创建/更新，`java.jdt.ls.java.home` 指向有效的 JDK 路径
- 确认路径格式符合当前平台规范

## 检查清单

- [ ] 已确认项目为标准 Maven Java 工程结构（`pom.xml` + `src/main/java`）
- [ ] 已通过询问+选项方式确认 JDK 来源和路径
- [ ] JDK 路径已验证有效
- [ ] `.opencode/plugins/inject-env.js` 已创建，支持当前平台（darwin/linux/win32）
- [ ] `.vscode/settings.json` 已创建/更新，JDK 路径与插件一致且格式正确
