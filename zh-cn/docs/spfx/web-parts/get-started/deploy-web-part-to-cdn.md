# <a name="deploy-your-sharepoint-client-side-web-part-to-a-cdn"></a>将 SharePoint 客户端 Web 部件部署到 CDN

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

在本文中要将 **HelloWorld** 资产部署到远程 CDN，而不使用本地环境。要使用一个集成 CDN 的 Azure 存储帐户来部署资产。SharePoint Framework 生成工具可提供对部署到 Azure 存储帐户的全新支持；但是也可将文件手动上载到自己喜爱的 CDN 提供程序或上载到 SharePoint。

也可以通过观看 [SharePoint PnP YouTube 频道](https://www.youtube.com/watch?v=xMQMNtsHDhk&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq)上的视频按照这些步骤进行操作。 

<a href="https://www.youtube.com/watch?v=xMQMNtsHDhk&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq">
<img src="../../../../images/spfx-youtube-tutorial3.png" alt="Screenshot of the YouTube video player for this tutorial" />
</a>

## <a name="prerequisites"></a>先决条件

开始之前，请确保已完成以下任务：

* [生成首个客户端 Web 部件](./build-a-hello-world-web-part)
* [将客户端 Web 部件连接到 SharePoint](./connect-to-sharepoint)
* [将客户端 Web 部件部署到经典 SharePoint 页面](./serve-your-web-part-in-a-sharepoint-page)

## <a name="configure-azure-storage-account"></a>配置 Azure 存储帐户

配置 Azure 存储帐户并将其与 CDN 集成。

可以按照文章[将存储帐户与 CDN 集成](https://azure.microsoft.com/en-us/documentation/articles/cdn-create-a-storage-account-with-cdn/)中的说明和本文中的详细步骤创建 Azure 存储帐户并将其与 CDN 集成。将需要以下信息。

### <a name="storage-account-name"></a>存储帐户名称

这是用于创建存储帐户的名称，方法如[步骤 1：创建存储帐户](https://azure.microsoft.com/en-us/documentation/articles/cdn-create-a-storage-account-with-cdn/#step-1-create-a-storage-account)中所述。

例如，在以下屏幕截图中，**spfxsamples** 是存储帐户名称。

![显示“新建存储帐户”页面的屏幕截图](../../../../images/deploy-create-storage-account.png)

此操作会新建一个存储帐户终结点 **spfxsamples.blob.core.windows.net**。 

>**注意：**需要为 SharePoint Framework 项目创建唯一的存储名称。


### <a name="blob-container-name"></a>BLOB 容器名称

创建新的 Blob 服务容器。此操作可在存储帐户仪表板中进行。

选择“**+ 容器**”，然后使用以下信息创建新的容器：

* 名称：**helloworld-webpart**
* 访问类型：容器

![显示创建 blob 容器的选项的图像](../../../../images/deploy-option-blob-container.png)

### <a name="storage-account-access-key"></a>存储帐户访问密钥

在存储帐户仪表板中，选择仪表板中的“**访问密钥**”并复制其中的一个访问密钥。

![显示存储帐户访问密钥的图像](../../../../images/deploy-storage-account-accesskey.png)

### <a name="cdn-profile-and-endpoint"></a>CDN 配置文件和终结点

创建新的 CDN 配置文件并将 CDN 终结点与该 BLOB 容器关联。

创建新的 CDN 配置文件，方法如[步骤 2：创建新的 CDN 配置文件](https://azure.microsoft.com/en-us/documentation/articles/cdn-create-a-storage-account-with-cdn/#step-2-create-a-new-cdn-profile)所述。

例如，在以下屏幕截图中，**spfxwebparts** 是 CDN 配置文件名称。

![创建新的 CDN 配置文件的屏幕截图](../../../../images/deploy-create-cdn-profile.png)

创建 CDN 终结点，方法如[步骤 3：创建新的 CDN 终结点](https://azure.microsoft.com/en-us/documentation/articles/cdn-create-a-storage-account-with-cdn/#step-3-create-a-new-cdn-endpoint)所述。

例如，在以下屏幕截图中，**spfxsamples** 是终结点名称，**Storage** 是原始类型，**spfxsamples.blob.core.windows.net** 是存储帐户。

![创建 CDN 终结点的屏幕截图](../../../../images/deploy-create-cdn-endpoint.png)

将使用以下 URL：http://spfxsamples.azureedge.net 创建 CDN 终结点

因为已将 CDN 终结点与存储帐户关联，因此也可通过以下 URL：http://spfxsamples.azureedge.net/helloworld-webpart/ 访问 BLOB 容器

但是请注意，现在尚未部署这些文件。

## <a name="project-directory"></a>项目目录

切换到控制台，并确保仍然位于安装 Web 部件项目所使用的项目目录中。

选择 **Ctrl+C** 结束 **gulp serve** 任务，并转至项目目录：

```
cd helloworld-webpart
```

## <a name="configure-azure-storage-account-details"></a>配置 Azure 存储帐户详细信息

切换到 Visual Studio Code，然后转至 **HelloWorld** Web 部件项目。

在 **config** 文件夹中打开 **deploy-azure-storage.json**。

这是包含 Azure 存储帐户详细信息的文件。

```json
{
  "workingDir": "./temp/deploy/",
  "account": "<!-- STORAGE ACCOUNT NAME -->",
  "container": "helloworld-webpart",
  "accessKey": "<!-- ACCESS KEY -->"
}
```

使用自己的存储帐户名称、BLOB 容器和存储帐户访问密钥分别替换 **account**、**container**、**accessKey**。

**workingDir** 是 Web 部件资产将要位于的目录。

在此示例中使用前面步骤中创建的存储帐户，该文件将如下所示：

```json
{
  "workingDir": "./temp/deploy/",
  "account": "spfxsamples",
  "container": "helloworld-webpart",
  "accessKey": "q1UsGWocj+CnlLuv9ZpriOCj46ikgBvDBCaQ0FfE8+qKVbDTVSbRGj41avlG73rynbvKizZpIKK9XpnpA=="
}
```

保存该文件。

## <a name="prepare-web-part-assets-to-deploy"></a>准备要部署的 Web 部件资产

需要生成资产，然后才能将其上载到 CDN。

切换到控制台并执行以下 `gulp` 任务：

```
gulp --ship
```

此操作将生成上载到 CDN 提供程序所需的缩小资产。`--ship` 指示要生成分发的生成工具。还应注意生成工具的输出指示“生成目标”为“SHIP”。

```
Build target: SHIP
[21:23:01] Using gulpfile ~/apps/helloworld-webpart/gulpfile.js
[21:23:01] Starting gulp
[21:23:01] Starting 'default'...
```

可在 `temp\deploy` 目录下找到缩小的资产。

## <a name="deploy-assets-to-azure-storage"></a>将资产部署到 Azure 存储

切换到 **HelloWorld** 项目目录的控制台。

输入 gulp 任务以将资产部署到存储帐户：

```
gulp deploy-azure-storage
```

此操作会将 Web 部件捆绑包及 JavaScript 和 CSS 文件等其他资产部署到 CDN。

### <a name="configuring-web-part-to-load-from-cdn"></a>将 Web 部件配置为从 CDN 加载

要使 Web 部件从 CDN 加载，需要告诉它  CDN 路径。

切换到 Visual Studio Code，然后从 **config** 文件夹打开 **write-manifests.json**。

在 **cdnBasePath** 属性中输入 CDN 基路径。

```json
{
  "cdnBasePath": "<!-- PATH TO CDN -->"
}
```

在此示例中使用前面步骤中创建的 CDN 配置文件，该文件将如下所示：

```json
{
  "cdnBasePath": "https://spfxsamples.azureedge.net/helloworld-webpart/"
}
```

>**注意：**CDN 基路径是已关联 BLOB 容器的 CDN 终结点。

保存该文件。

## <a name="deploy-the-updated-package"></a>部署已更新的包

### <a name="package-the-solution"></a>打包解决方案

因为已更改 Web 部件捆绑包，因此需要将此包重新部署到应用目录中。前面使用 **--ship** 生成了用于分发的缩小资产。

切换到 **HelloWorld** 项目目录的控制台。

输入 gulp 任务以打包客户端解决方案，这次使用设置的 `--ship` 标记。此操作将强制任务挑选在前面步骤中配置的 CDN 基路径：

```
gulp bundle --ship
gulp package-solution --ship
```

> **注意：**“gulp bundle --ship”是开发者预览版所需的临时修复方法，用于确保重新生成文件正确进行以便进行打包。

此操作将在 **sharepoint\solution** 文件夹中创建更新的客户端解决方案包。

### <a name="upload-to-your-app-catalog"></a>上载到应用目录

将客户端解决方案包上载或拖放到应用目录中。

因为已部署过此包，因此会出现是否要替换已有包的提示。

![替换客户端解决方案包提示的屏幕截图](../../../../images/sp-app-replace-pkg.png)

选择“**替换它**”。

应用目录现在将具有最新的客户端解决方案包，在此包中 Web 部件捆绑包从 CDN 中加载。

此操作将更新 SharePoint 中 **HelloWorld** Web 部件的所有实例以立即从 CDN 提取资源。

## <a name="test-the-helloworld-web-part"></a>测试 HelloWorld Web 部件

### <a name="classic-sharepoint-page"></a>经典 SharePoint 页面

转至已创建的 **HelloWorld** Web 部件页。**HelloWorld** Web 部件现在将从 CDN 加载 Web 部件捆绑包和其他资产。

注意将不再运行 **gulp serve**，因此无任何内容是通过 **localhost** 提供。

## <a name="deploying-to-other-cdns"></a>部署到其他 CDN

要将资产部署到喜爱的 CDN 提供程序，可以复制 **tmp\deploy** 文件夹中的文件。若要生成用于分发的资产，需要按照之前的方法使用 **--ship** 参数运行以下 gulp 命令：

```
gulp --ship
```

## <a name="next-steps"></a>后续步骤

可以加载 jQuery、jQuery UI，然后生成 jQuery Accordion Web 部件。若要继续操作，请参阅[将 jQueryUI Accordion 添加到客户端 Web 部件](./add-jqueryui-accordion-to-web-part)。
