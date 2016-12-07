# <a name="deploy-your-client-side-web-part-to-a-classic-sharepoint-page-hello-world-part-3"></a>将客户端 Web 部件部署到经典 SharePoint 页面（Hello world 第 3 部分）

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

在本文中，你将学习如何将客户端 Web 部件部署到 SharePoint 并查看它在经典 SharePoint 服务器端页面上的运行方式。本文继续介绍在以前文章 [将客户端 Web 部件连接到 SharePoint](./connect-to-sharepoint) 中构建的 hello world web 部件。

在开始之前，请确保你已完成以下文章中的步骤：

* [构建首个 SharePoint 客户端 Web 部件](./build-a-hello-world-web-part)
* [将客户端 Web 部件连接到 SharePoint](./connect-to-sharepoint)

也可以通过观看 [SharePoint PnP YouTube 频道](https://www.youtube.com/watch?v=G9JB1HuNs7Q&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq)上的视频，按照这些步骤进行操作。 

<a href="https://www.youtube.com/watch?v=G9JB1HuNs7Q&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq">
<img src="../../../../images/spfx-youtube-tutorial3.png" alt="Screenshot of the YouTube video player for this tutorial" />
</a>


## <a name="package-the-helloworld-web-part"></a>打包 HelloWorld Web 部件

在控制台窗口中，转到“[构建首个 SharePoint 客户端 Web 部件](./build-a-hello-world-web-part)”中创建的 Web 部件项目目录。

```
cd helloworld-webpart
```

如果 `gulp serve` 仍在运行，通过选择 `Ctrl+C` 停止运行

不同于工作台上，为了在经典 SharePoint 服务器端页面中使用客户端 Web 部件，你需要对 SharePoint 部署和注册 Web 部件。首先，你需要打包 Web 部件。

在 Visual Studio Code 或你的首选 IDE 中打开 **HelloWorldWebPart** Web 部件项目。

从 **config** 文件夹打开 **package-solution.json**。

**package-solution.json** 文件定义包元数据，如下面的代码所示：

```json
{
    "solution": {
    "name": "helloworld-webpart-client-side-solution",
    "id": "ed83e452-2286-4ea0-8f98-c79d257acea5",
    "version": "1.0.0.0"
    },
    "paths": {
    "zippedPackage": "helloworld-webpart.spapp"
    }
}
```

在控制台窗口中，输入以下命令以打包包含 Web 部件的客户端解决方案：

```
gulp package-solution
```

命令将在 `sharepoint` 文件夹中创建包：

```
helloworld-webpart.spapp
```

### <a name="package-contents"></a>包内容

包使用 SharePoint 功能来打包 Web 部件。默认情况下，gulp 任务创建以下内容：

* Web 部件的功能。
* Web 部件的 .webpart 文件，是一个描述 Web 部件的 XML 文件。

你可以在 **sharepoint** 文件夹中查看原始包内容。 

然后将内容打包到 **.spapp** 文件。包格式非常类似于 SharePoint 外接程序包并使用 Microsoft 开放式打包约定来打包解决方案。 

JavaScript 文件、CSS 和其他资产未打包，必须将它们部署到 CDN 等外部位置。为了在开发过程中测试 Web 部件，你可以从本地计算机加载所有资产。 

## <a name="deploy-the-helloworld-package-to-app-catalog"></a>将 HelloWorld 包部署到应用目录

接下来，你需要将所生成的包部署到应用程序目录。

转到网站的应用程序目录。

上载 **helloworld-webpart.spapp** 或将其拖放到应用目录。

![将解决方案上载到应用目录](../../../../images/upload-solution-app-catalog.png) 

这将部署客户端解决方案包。由于这是一个完全信任的客户端解决方案，SharePoint 将显示一个对话框并要求你信任要部署的客户端解决方案。

![信任客户端解决方案部署](../../../../images/sp-app-deploy-trust.png) 
    
选择“**部署**”

>**注意：**如果包部署失败，则你可能正在使用普通租户。请确保[设置 Office 365 租户](../../set-up-your-developer-tenant)以启用“第一版”选项，或者使用 Office 365 开发人员租户。 

## <a name="install-the-client-side-solution-on-your-site"></a>在网站上安装客户端解决方案

转到开发人员网站集。

选择右侧顶部导航栏上的齿轮图标并选择“**添加应用**”以转到“应用”页面。

在“**搜索**”框中，输入 **helloworld**，然后选择“**Enter**”筛选你的应用。
    
![将应用添加到网站](../../../../images/install-app-your-site.png) 
    
选择 **helloworld-webpart-client-side-solution** 应用以在网站上安装该应用。
    
![信任应用](../../../../images/app-installed-your-site.png) 

在开发人员网站上安装客户端解决方案和 Web 部件。

“**网站内容**”页将显示客户端解决方案的安装状态。请确保在转到下一步之前，安装已完成。

## <a name="preview-the-web-part-in-a-classic-sharepoint-page"></a>在经典 SharePoint 页面中预览 Web 部件

现在，你已经部署并安装客户端解决方案，将 Web 部件添加到经典 SharePoint 页面。请记住，JavaScripts 和 CSS 等资源可以从本地计算机获取。

从 **\dist** 文件夹打开 **<你的 Web 部件 guid>.manifest.json**。
    
请注意，**loaderConfig** 项中的 **internalModuleBaseUrls** 属性仍引用你的本地计算机：

```json
"internalModuleBaseUrls": [
    "http://`your-local-machine-name`:4321/"
]
```

将 Web 部件添加到 SharePoint 服务器端页面之前，先运行本地服务器。
    
在包含 **helloworld-webpart** 项目目录的控制台窗口中，运行 gulp 任务以从 localhost 开始提供服务：
    
```
gulp serve --nobrowser
```

>**注意：** `--nobrowser`不会自动启动 Web 部件工作台。

## <a name="add-the-helloworld-web-part-to-classic-page"></a>将 HelloWorld Web 部件添加到经典页面

在浏览器中转到网站集。
    
在随后的步骤中，创建一个经典页面，然后转到网站中的 **SitePages** 库。
    
在右侧顶部导航栏中选择齿轮图标，并选择“**网站内容**”。
    
选择 **SitePages** 库图标，以转至 **SitePages** 库。
    
选择“**新建**”创建一个经典 SharePoint 页面。
    
输入 **HelloWorld** 作为页面名称。
    
选择“**创建**”按钮以创建 Web 部件页面。SharePoint 将创建你的页面。
    
在功能区中，选择“**插入-> Web 部件**”以打开 Web 部件库。
    
在 Web 部件库中，选择类别“**自定义**”。
    
>**注意：**在预览时，可在 Web 部件库的“**自定义**”类别下获取客户端 Web 部件。 

你应看到 Hello World Web 部件。

![使用自定义类别打开的 Web 部件库](../../../../images/webpart-gallery-helloworld.png)
    
选择 Hello World Web 部件，然后选择“**添加**”以将其添加到页面。
    
将从本地环境加载 Web 部件资产。为了加载本地计算机上托管的脚本，你需要启用浏览器以加载不安全脚本。这取决于使用的浏览器，请确保启用加载此会话的不安全脚本。
    
你应该会看到之前文章中构建的 **HelloWorld** Web 部件（从当前网站检索列表）。 

![经典页面中的 Hello World Web 部件](../../../../images/sp-wp-classic-page.png)

## <a name="edit-web-part-properties"></a>编辑 Web 部件属性

选择 Web 部件编辑菜单，选择“**编辑 Web 部件**”以打开 Web 部件的属性窗格。
    
![编辑 Web 部件](../../../../images/edit-webpart-classic-page.png)

属性窗格作为服务器端 Web 部件属性窗格打开。但是，你可以选择为客户端 Web 部件配置属性。

![配置 Web 部件 - 属性窗格选项](../../../../images/webpart-configure-property-pane.png)
    
选择“**配置**”按钮，以显示客户端 Web 部件的新客户端属性窗格。
    
这与在工作台中构建和预览的属性窗格相同。
    
编辑“**描述**”属性，输入“**Client-side web parts are awesome!**”
    
![经典页面中的 Hello World Web 部件](../../../../images/sp-wp-classic-page-pp.png)

注意，在键入时更新 Web 部件的位置，仍然有相同的行为（如反应窗格）。
    
选择 **x** 图标以关闭客户端属性窗格。
    
>**注意：**将需要选择 **x** 图标数次以关闭属性窗格。这是一个已知问题。
    
选择服务器端属性窗格中的“**确定**”按钮以保存并关闭 Web 部件属性窗格。
    
因为 Web 部件在经典 SharePoint 页面中运行，选择“**确定**”或“**应用**”按钮将保存 Web 部件属性。
    
在功能区中，选择“**保存**”来保存该页面。

## <a name="next-steps"></a>后续步骤

恭喜你！你已将客户端 Web 部件部署到经典 SharePoint 页面你可以在下一主题“[将客户端 Web 部件源部署到 CDN](./deploy-web-part-to-cdn)”中继续构建出 Hello World Web 部件，你将从中学习如何从 CDN（而非 localhost）部署和加载 Web 部件资产。
