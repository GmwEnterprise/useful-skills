## 项目说明

本项目仓库为个人开发技能专用仓库，有关技能的创建、管理、维护均在目录 `skills/<skill-name>/` 中进行。
该项目主要使用 opencode 进行开发与维护，该项目所需要用到的各种 commands、skills 请按 opencode 的官方推荐形式创建及使用。

## 必须遵守

- 创建技能之前，请加载 skill-creator 技能，确认创建一个技能所需要注意的所有细节；如果用户未安装此技能，则停止操作，提示用户安装；
- 所有技能默认使用的用户环境已经具备 uv、nodejs；
- 技能创建优先基于内嵌项目：`skills/<skill-name>/packages/<embedded_project>/`；
- 内嵌项目技术栈优先使用 python(uv) 或 javascript(npm)，基于技能所用场景选择最合适的技术栈；
- 创建技能时，初始化内嵌项目优先使用 uv 或 npm 直接初始化；
- 内嵌项目调用一律要求使用 bash/powershell 脚本间接调用，脚本请存放到 `skills/<skill-name>/scripts/` 目录中；
- 不需要额外依赖的功能，则直接创建 `.py` 或 `.js` 即可；