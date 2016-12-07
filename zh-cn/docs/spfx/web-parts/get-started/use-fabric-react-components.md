# <a name="use-office-ui-fabric-react-components-in-your-sharepoint-client-side-web-part"></a>在 SharePoint 客户端 Web 部件中使用 Office UI Fabric React 部件。

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

本文介绍如何生成一个简单的 Web 部件，以使用 [Office UI Fabric](https://github.com/OfficeDev/office-ui-fabric-react) 的 DocumentCard 组件。Office UI Fabric React 是用于构建 Office 和 Office 365 体验的前端框架。Fabric React 包括一个快速响应、移动优先的组件集合，便于使用 Office 设计语言来创建 Web 体验。

>**注意：**Office UI Fabric React 组件处于 *v1 之前的状态*。有关 v1 版本和路线图的信息，请参阅 GitHub 中 Office UI Fabric React 存储库中的[路线图](https://github.com/OfficeDev/office-ui-fabric-react/blob/master/ghdocs/ROADMAP.md)。 

下图显示了一个使用 Office UI Fabric React 创建的 DocumentCard 组件。

![SharePoint 工作台中的 DocumentCard Fabric 组件的图像](../../../../images/fabric-components-doc-card-view-ex.png)

也可以通过观看 [SharePoint PnP YouTube 频道](https://www.youtube.com/watch?v=P8WmNhcSWHU&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq)上的视频，按照这些步骤进行操作。 

<a href="https://www.youtube.com/watch?v=P8WmNhcSWHU&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq">
<img src="../../../../images/spfx-youtube-tutorial3.png" alt="Screenshot of the YouTube video player for this tutorial" />
</a>


## <a name="create-a-new-web-part-project"></a>创建新的 Web 部件项目

在你最喜爱的位置创建新的项目目录：

```
md documentcardexample-webpart
```
    
获取项目目录：

```
cd documentcardexample-webpart
```

通过运行 Yeoman SharePoint 生成器创建新的 Web 部件：

```
yo @microsoft/sharepoint
```
    
出现提示时：

* 接受默认的 **documentcardexample-webpart** 作为你的解决方案名称，并选择“**输入**”。
* 选择“**使用当前文件夹**”作为放置文件的位置。
* 将 **DocumentCardExample** 用作 Web 部件名称，并选择“**Enter**”。
* 接受默认的 **DocumentCardExample 说明**，然后选择“**Enter**”。
* 选择“**React**”作为框架，然后选择“**Enter**”。

在这种情况下，Yeoman 将安装必需的依赖项并为解决方案文件提供基架。这可能需要几分钟的时间。Yeoman 将为项目提供基架，以同时包括 DocumentCardExample Web 部件。
    
构架完成后，在控制台中键入以下命令，以在 Visual Studio 代码中打开 Web 部件项目：

```
code .
```
    
现在你拥有带 React 框架的 Web 部件项目。

从 **src\webparts\documentCardExample** 文件夹打开 **DocumentCardExampleWebPart.ts**。 

正如你所看到的，`render` 方法将创建一个响应元素，并将其呈现在 Web 部件 DOM 中。

```ts
public render(mode: DisplayMode, data?: IWebPartData): void {
    const element: React.ReactElement<IDocumentCardProps> = React.createElement(DocumentCard, {
        description: this.properties.description
    });

ReactDom.render(element, this.domElement);
}
```
    
从 **src\webparts\documentCardExample\components** 文件夹打开 **DocumentCardExample.tsx**。 
    
这是 Yeoman 添加到项目的主要响应组件，它将呈现在 Web 部件 DOM 中。

```ts
export default class DocumentCardExample extends React.Component<IDocumentCardProps, {}> {
public render(): JSX.Element {
    return (
        <div className={styles.documentcard}>
            <div className={styles.container}>
                <div className={css('ms-Grid-row ms-bgColor-themeDark ms-fontColor-white', styles.row)}>
                    <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
                        <span className='ms-font-xl ms-fontColor-white'>
                            Welcome to SharePoint!
                        </span>
                        <p className='ms-font-l ms-fontColor-white'>
                            Customize SharePoint experiences using Web Parts.
                        </p>
                        <p className='ms-font-l ms-fontColor-white'>
                            {this.props.description}
                        </p>
                        <a
                            className={css('ms-Button', styles.button)}
                            href='https://github.com/SharePoint/sp-dev-docs/wiki'
                        >
                            <span className='ms-Button-label'>Learn more</span>
                        </a>
                    </div>
                </div>
            </div>
        </div>
        );
    }
}
```

## <a name="add-an-office-ui-fabric-component"></a>添加 Office UI Fabric 组件

若要使用 Office UI Fabric 组件，需要安装 npm 包。

在控制台窗口中，键入以下命令以安装 Office UI Fabric 组件 npm 包：

```
npm i office-ui-fabric-react --save
```

>**注意：**如果正在使用 Windows，则可能会在安装过程中收到以下异常信息：“错误：EPERM：不允许操作，...。”若要解决此问题，请以管理员身份重新打开控制台应用程序。打开上下文菜单（右键单击），并选择“**以管理员身份运行**”。 

## <a name="add-the-documentcard-component"></a>添加 DocumentCard 组件

Office UI Fabric React 组件安装后，可以将该组件添加到 web 部件。 

从 **src\webparts\components\documentCardExample** 文件夹打开 **DocumentCardExample.tsx**。 

将以下 `import` 语句添加到文件顶部，以导入要使用的结构响应组件。

```ts
import {
    DocumentCard,
    DocumentCardPreview,
    DocumentCardTitle,
    DocumentCardActivity,
    IDocumentCardPreviewProps
} from 'office-ui-fabric-react/lib/DocumentCard';
```

删除当前 `render` 方法并添加以下更新的 `render` 方法：

```ts
public render(): JSX.Element {
    const previewProps: IDocumentCardPreviewProps = {
    previewImages: [
        {
        previewImageSrc: require('document-preview.png'),
        iconSrc: require('icon-ppt.png'),
        width: 318,
        height: 196,
        accentColor: '#ce4b1f'
        }
    ],
    };

    return (
        <DocumentCard onClickHref='http://bing.com'>
        <DocumentCardPreview { ...previewProps } />
        <DocumentCardTitle title='Revenue stream proposal fiscal year 2016 version02.pptx'/>
        <DocumentCardActivity
            activity='Created Feb 23, 2016'
            people={
            [
                { name: 'Kat Larrson', profileImageSrc: require('avatar-kat.png') }
            ]
            }
        />
        </DocumentCard>
    );
}
```

保存该文件。

在此代码中，DocumentCard 组件包括一些额外部分：

* DocumentCardPreview
* DocumentCardTitle
* DocumentCardActivity

`previewProps` 属性包含 DocumentCardPreview 的一些属性。

请注意，将相对路径与 `require` 语句结合使用来加载图像。当前，需要使用 webpack 公共路径插件并将源文件或文件夹中的文件的相对路径输入到 `lib` 文件夹。这应该与当前工作源位置相同。
    
从 **src\webparts\documentCardExample** 文件夹打开 **DocumentCardExampleWebPart.ts**。 
    
在文件顶部添加以下代码，以请求 webpack 公共路径插件。
    
```ts
require('set-webpack-public-path!');
```
    
保存该文件。

## <a name="copy-the-image-assets"></a>复制图像资产

将以下图像复制到你的 **src\webparts\documentCardExample** 文件夹：

* [avatar-kat.png](https://github.com/SharePoint/sp-dev-docs/blob/master/assets/avatar-kat.png)
* [icon-ppt.png](https://github.com/SharePoint/sp-dev-docs/tree/master/assets/icon-ppt.png)
* [document-preview.png](https://github.com/SharePoint/sp-dev-docs/tree/master/assets/document-preview.png)

## <a name="preview-the-web-part-in-workbench"></a>预览工作台中的 Web 部件

在控制台中，键入以下命令以预览工作台中的 Web 部件：
    
```
gulp serve
```
    
在工具箱中，选择 `DocumentCardExample` Web 部件以添加以下项：
    
![SharePoint 工作台中的 DocumentCard Fabric 组件的图像](../../../../images/fabric-components-doc-card-view-ex.png)

