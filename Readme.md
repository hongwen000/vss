# VSCode Window Launcher

VSCode Window Launcher是一个跨平台实用程序，帮助开发者快速定位和切换Windows系统中的窗口。尽管它是一个通用的窗口切换程序，但对VSCode窗口使用智能搜索算法进行优先搜索，并根据窗口标题和用户优先级找到最相关的VSCode窗口。

## 特点

- **非常快**：是否受尽Windows系统不稳定的任务栏的折磨？是否对着一堆窗口按下Alt+F4时感到不知所措？本应用能非常快地帮你找到并切换到你想要的窗口！
- **智能窗口搜索**：使用智能搜索算法高效定位和聚焦所需的VSCode窗口，该算法会考虑到窗口标题的完全匹配以及首字母匹配。
- **优先级排序结果**：搜索结果根据完全标题匹配和首字母缩写匹配进行排序，用户输入的缩写匹配具有更高的权重。

## 安装

要开始使用VSCode Window Launcher，请按照以下步骤进行操作：

1. 克隆存储库：

```
git clone https://github.com/hongwen000/vss.git
```

2. 进入克隆的目录：

```
cd vss
```

3. 安装所需的依赖项：

```
pip install -r requirements.txt
```

4. 运行应用程序：

```
python main.py
```

5. 创建快捷方式：

```
python setup.py
```

这样，您就可以在桌面和开始菜单上创建VSCode Window Launcher的快捷方式。现在，您可以快速搜索和定位打开的VSCode窗口了。

## 全局快捷键

VSCode Window Launcher还提供了一个全局快捷键**Alt+Space**，用于快速激活应用程序并开始搜索窗口。按下**Alt+Space**组合键将打开VSCode Window Launcher界面，您可以立即开始输入搜索查询，然后按Enter键以快速切换到所选的VSCode窗口或其他窗口。

## 贡献

我们欢迎社区的贡献！如果您想为该项目做出贡献，请按照以下步骤进行操作：

1. 在GitHub上Fork此存储库。
2. 将您Fork的存储库克隆到本地机器。
3. 将您的更改提交到一个新的分支。
4. 将您的分支推送到您Fork的存储库。
5. 在此存储库的`main`分支中打开一个拉取请求。

请在所有拉取请求中包含对您更改的描述。这将帮助维护人员更有效地审查您的贡献。


## 支持

如果您有任何问题、问题或建议，请在[GitHub存储库](https://github.com/hongwen000/vss/issues)上提出问题，我们将乐意帮助.

# VSCode Window Launcher

VSCode Window Launcher is a cross-platform utility that helps developers quickly locate and switch between windows on the Windows operating system. While it is a general window switching program, it prioritizes searching for Visual Studio Code (VSCode) windows using an intelligent search algorithm based on window title and user priorities.

## Features

- **Lightning Fast**: Tired of struggling with the unpredictable Windows taskbar? Feeling overwhelmed when trying to navigate through a sea of open windows with Alt+F4? This application is lightning fast in helping you locate and switch to the window you desire!
- **Intelligent Window Search**: Efficiently locates and brings focus to desired VSCode windows using a smart search algorithm that considers both complete title matches and first letter matches of window titles.
- **Priority-based Sorting**: Search results are sorted based on full title matches and first letter abbreviation matches, with user-input abbreviation matches given higher priority.

## Installation

To get started with VSCode Window Launcher, follow these steps:

1. Clone the repository:

```
git clone https://github.com/hongwen000/vss.git
```

2. Navigate to the cloned directory:

```
cd vss
```

3. Install the required dependencies:

```
pip install -r requirements.txt
```

4. Run the application:

```
python main.py
```

5. Create shortcuts:

```
python setup.py
```

This will create shortcuts for VSCode Window Launcher on your desktop and start menu. Now you can quickly search and locate open VSCode windows.

## Global Shortcut Key

VSCode Window Launcher also provides a global shortcut key **Alt+Space** to quickly activate the application and start searching for windows. Pressing the **Alt+Space** key combination will open the VSCode Window Launcher interface, where you can immediately start typing your search query, and then press Enter to quickly switch to the selected VSCode window or other windows.

## Contributing

We welcome contributions from the community! If you would like to contribute to the project, please follow these steps:

1. Fork this repository on GitHub.
2. Clone your forked repository to your local machine.
3. Commit your changes to a new branch.
4. Push your branch to your forked repository.
5. Open a Pull Request against the `main` branch in this repository.

Please include a description of your changes with any Pull Requests. This will help maintainers review your contribution more effectively.

## Support

If you have any questions, issues, or suggestions, please open an issue on the [GitHub repository](https://github.com/hongwen000/vss/issues), and we will be happy to assist you.

Please boast more features and provide the English version after making the modifications.
