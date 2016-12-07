# <a name="set-up-your-sharepoint-client-side-web-part-development-environment"></a>设置 SharePoint 客户端 Web 部件开发环境

>**注意：**SharePoint Framework 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

可以使用 Visual Studio 或你自己的自定义开发环境生成 SharePoint 客户端 Web 部件。可以使用 Mac、PC 或 Linux。

>**注意：**执行本文中的步骤之前，请确保[设置 Office 365 租户](../../set-up-your-developer-tenant)。

## <a name="install-developer-tools"></a>安装开发人员工具

### <a name="nodejs"></a>NodeJS
安装 [NodeJS](https://nodejs.org/en/) 长期支持 (LTS) v4.x.x 版本。

* 如果已经安装 NodeJS，请务必使用 `node -v` 检查最新版本。它应返回最新的 [LTS 版本](https://nodejs.org/en/download/)。 
* 如果使用的是 Mac，建议使用 [homebrew](http://brew.sh/) 来安装和管理 NodeJS。 

安装节点后，请通过运行以下命令确保正在运行 npm 的 V3：
    
```
npm -g install npm@3
```

如果使用 Ubuntu Linux，前一个命令可能导致出现“权限被拒绝”消息，因此它应改为作为以下命令执行 

```
sudo npm -g install npm@3
```

### <a name="code-editors"></a>代码编辑器
安装代码编辑器。可以使用任何支持客户端开发的代码编辑器或 IDE 来构建 Web 部件，如：

* [Visual Studio Code](https://code.visualstudio.com/)
* [Sublime](https://www.sublimetext.com/)
* [Atom](https://atom.io)
* [Webstorm](https://www.jetbrains.com/webstorm) 

此文档中的步骤和示例使用 [Visual Studio Code](https://code.visualstudio.com/)，但你可以使用你选择的任何编辑器。 

### <a name="if-youre-using-a-pc"></a>如果你使用的是 PC

需要安装 *windows-build-tools*。windows-build-tools 将安装由 Microsoft 免费提供的 Visual C++ Build Tools 2015。这些工具需要编译热门本机模块。它还将安装 Python 2.7，适当地配置计算机和 npm。 

运行以下命令：
    
```
npm install --global --production windows-build-tools
```

#### <a name="if-you-are-using-visual-studio"></a>如果使用的是 Visual Studio 
如果要将 Visual Studio 作为开发环境使用，请安装以下所需的工具及更新：

* [Visual Studio 2015](https://go.microsoft.com/fwlink/?LinkId=691978&clcid=0x409)
* [Visual Studio Update 3](https://www.visualstudio.com/en-us/news/releasenotes/vs2015-update3-vs) 或更高版本
* [适用于 Visual Studio 的 Node.js 工具](https://aka.ms/getntvs)

### <a name="if-you-are-using-ubuntu"></a>如果使用的是 Ubuntu

需要使用以下命令安装编译器工具：
    
```
sudo apt-get install build-essential
```

### <a name="if-you-are-using-fedora"></a>如果使用的是 fedora

需要使用以下命令安装编译器工具：
    
```
sudo yum install make automake gcc gcc-c++ kernel-devel
```

## <a name="install-yeoman-and-gulp"></a>安装 Yeoman 和 gulp

[Yeoman](http://yeoman.io/) 可帮助启动新项目，并规定最佳做法和工具，以帮助你保持高效率。SharePoint 客户端开发工具包括用于创建新 Web 部件的 Yeoman 生成器。生成器为主机 Web 部件提供公共生成工具、公共样本代码和公共体育场网站以进行测试。

输入以下命令以安装 Yeoman 和 gulp：
    
```
npm i -g yo gulp
```

## <a name="install-yeoman-sharepoint-generator"></a>安装 Yeoman SharePoint 生成器

Yeoman SharePoint Web 部件生成器帮助你快速创建一个具有适当工具链和项目结构的 SharePoint 客户端解决方案项目。

输入以下命令以安装 Yeoman SharePoint 生成器：
    
```
npm i -g @microsoft/generator-sharepoint 
```

## <a name="optional-tools"></a>可选工具

下面同样是一些可能会用到的工具：

* [Fiddler](http://www.telerik.com/fiddler)
* [适用于 Chrome 的 Postman 插件](https://www.getpostman.com/docs/introduction)
* [Cmder for Windows](http://cmder.net/)
* [Oh My Zsh for Mac](http://ohmyz.sh/)
* [Git 源控件工具](https://git-scm.com/)

## <a name="next-steps"></a>后续步骤

现在可以[生成首个客户端 Web 部件](web-parts/get-started/build-a-hello-world-web-part)！
