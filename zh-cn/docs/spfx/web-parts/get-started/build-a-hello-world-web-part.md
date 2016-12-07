# <a name="build-your-first-sharepoint-client-side-web-part-hello-world-part-1"></a>生成首个 SharePoint 客户端 Web 部件（Hello World 第 1 部分）。

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

客户端 Web 部件是运行在 SharePoint 页面的上下文中的客户端组件。可以将客户端 Web 部件部署到 SharePoint Online，还可以使用现代 JavaScript 工具和库生成它们。

客户端 Web 部件支持：

* 使用 HTML 和 JavaScript 生成。
* SharePoint Online 和本地环境。

>**注意：**执行本文中的步骤之前，请确保[设置开发环境](../../set-up-your-development-environment)。

也可以通过观看 [SharePoint PnP YouTube 频道](https://www.youtube.com/watch?v=ralspfOBgic&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq)上的视频，按照这些步骤进行操作。 

<a href="https://www.youtube.com/watch?v=ralspfOBgic&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq">
<img src="../../../../images/spfx-youtube-tutorial5.png" alt="Screenshot of the YouTube video player for this tutorial" />
</a>


## <a name="create-a-new-web-part-project"></a>创建新的 Web 部件项目
在最喜爱的位置创建新的项目目录：
    
```
md helloworld-webpart
```

转到项目目录：

```
cd helloworld-webpart
```

通过运行 Yeoman SharePoint 生成器创建新的 HelloWorld Web 部件：

```
yo @microsoft/sharepoint
```
    
出现提示时：

* 接受默认的 **helloworld-webpart** 作为解决方案名称，并选择“**Enter**”。
* 选择**使用当前文件夹**放置文件。

接下来的一组提示会要求提供有关 Web 部件的特定信息：

* 接受默认的 **helloworld-webpart** 作为 Web 部件名称，并选择“**Enter**”。
* 接受默认的 **HelloWorld description** 作为 Web 部件描述，并选择“**Enter**”。
* 接受默认的 **No javascript web framework** 作为要使用的框架，并选择“**Enter**”。

![Yeoman SharePoint 生成器提示创建 Web 部件客户端解决方案](../../../../images/yeoman-sp-prompts.png)

在这种情况下，Yeoman 将安装必需的依赖项并为解决方案文件和 **HelloWorld** Web 部件提供基架。这可能需要几分钟的时间。 

基架完成后，应该可以看到指示成功搭建基架的以下消息：

![SharePoint 客户端解决方案成功搭建基架](../../../../images/yeoman-sp-complete.png)

有关任何错误故障排除的信息，请参阅[已知问题](../basics/known-issues)。

### <a name="using-your-favorite-code-editor"></a>使用最喜爱的代码编辑器
因为 SharePoint 客户端解决方案基于 HTML/TypeScript，所以可使用任何支持客户端开发的代码编辑器或 IDE 来生成 Web 部件，例如：

* [Visual Studio Code](https://code.visualstudio.com/)
* [Sublime](https://www.sublimetext.com/)
* [Atom](https://atom.io)
* [Webstorm](https://www.jetbrains.com/webstorm)

>**注意：**本文在步骤和示例中使用 Visual Studio code。你可以使用你喜欢的任何代码编辑器。 
   
## <a name="preview-the-web-part"></a>预览 Web 部件
若要预览 Web 部件，请在本地 Web 服务器上生成并运行它。默认情况下，客户端工具链使用 HTTPS 端点。但是，由于没有为本地开发环境配置默认证书，因此浏览器将报告证书错误。SPFx 工具链带有开发者证书，可安装用于生成 Web 部件。

要安装开发者证书以在 SPFx 开发中使用，请切换到控制台，确保仍处于 **helloworld-webpart** 目录中，并输入以下命令：

```
gulp trust-dev-cert
```

现在，我们已安装开发者证书，在控制台中输入以下命令以生成并预览 Web 部件：

```
gulp serve
```

此命令将执行一系列 Gulp 任务，以在“localhost:4321”上创建基于节点的本地 HTTPS 服务器，并启动默认浏览器以从本地开发环境中预览 Web 部件。

![Gulp Serve Web 部件项目](../../../../images/helloworld-wp-gulp-serve.png)

SharePoint 客户端开发工具使用 [gulp](http://gulpjs.com/) 作为任务运行程序，以处理如下生成过程任务：

* 捆绑和缩小 JavaScript 与 CSS 文件。
* 在每个生成前运行工具以调用捆绑和缩小任务。
* 将 SASS 文件编译为 CSS。
* 将 TypeScript 文件编译为 JavaScript。

如果你是 gulp 新用户，则可以阅读 [Using Gulp](http://docs.asp.net/en/latest/client-side/using-gulp.html)（Gulp 使用），该文章介绍了如何使用 gulp 与 Visual Studio 联合生成 ASP.NET 5 项目。

Visual Studio Code 为 gulp 和其他任务运行程序提供内置支持。使用 Windows 上的 **Ctrl + Shift + B** 或 Mac 上的 **Cmd + Shift + B** 进行调试并预览 Web 部件。 

### <a name="sharepoint-workbench"></a>SharePoint 工作台
SharePoint 工作台是开发者设计界面，使你无需将 Web 部件部署在 SharePoint 中即可快速预览和测试。SharePoint 工作台包括客户端页面和客户端画布，你可在其中添加、删除和测试开发中的 Web 部件。

![在本地运行的 SharePoint 工作台](../../../../images/sp-workbench.png)

若要添加 HelloWorld Web 部件，请选择“**添加**”按钮。“添加”按钮将打开工具箱，其中可以看到要添加的可用 Web 部件列表。该列表将包括 **HelloWorld** Web 部件，以及开发环境中可用的其他本地 Web 部件。
   
![localhost 中的 SharePoint 工作台工具箱](../../../../images/sp-workbench-toolbox.png)
   
选择“**HelloWorld**”将 Web 部件添加到此页面：
   
![SharePoint 工作台中的 HelloWorld Web 部件](../../../../images/sp-workbench-helloworld-wp.png)

恭喜你！你刚刚将第一个客户端 Web 部件添加到了客户端页面。
   
现在，选择 Web 部件最左侧的铅笔图标，以显示 Web 部件属性窗格。
   
![HelloWorld Web 部件属性窗格](../../../../images/sp-workbench-helloworld-pp.png)

你可以在属性窗格中定义属性，以便自定义 Web 部件。属性窗格由客户端驱动，并在 SharePoint 中提供一致的设计。 
   
将 **Description** 文本框中的文本修改为 **Client-side web parts are awesome!**

请注意，Web 部件中的文本也会随着键入更改。 

属性窗格可用的新功能之一是配置其更新行为，该行为可设置为反应式或非反应式。默认的更新行为是反应式，并允许你在编辑属性时查看所做的更改。如果行为设置为反应式，所做的更改会立即保存。  

## <a name="web-part-project-structure"></a>Web 部件项目结构
可以使用 Visual Studio Code 来浏览 Web 部件项目结构。 

* 在控制台中，转到 **src\webparts\helloWorld** 目录。 
* 输入以下命令以在 Visual Studio Code 中打开 Web 部件项目（或使用最喜欢的编辑器）：

```
code .
```

![HelloWorld 项目结构](../../../../images/helloworld-wp-vscode-project-structure.png)

如果出现错误，则需要[安装 PATH 中的代码命令](https://code.visualstudio.com/docs/editor/setup)。

TypeScript 是生成 SharePoint 客户端 Web 部件的主要语言。TypeScript 是键入的 JavaScript 超集，它编译为纯 JavaScript。SharePoint 客户端开发工具使用 TypeScript 类、模块和接口生成，可帮助开发者生成可靠的客户端 Web 部件。 

以下是项目中的一些关键文件。

### <a name="web-part-class"></a>Web 部件类
**HelloWorldWebPart.ts** 定义 Web 部件的主入口点。Web 部件类 **HelloWorldWebPart** 扩展了 **BaseClientSideWebPart**。任何客户端 Web 部件均应扩展 **BaseClientSideWebPart** 类，才会被定义为一个有效的 Web 部件。

```ts
public constructor(context: IWebPartContext) {
    super(context);
}
```

**BaseClientSideWebPart** 实现生成 Web 部件所需的最少功能。此类还提供了许多参数，以验证并访问只读属性，例如 **displayMode**、Web 部件属性、Web 部件上下文、 web 部件 **instanceId**、Web 部件 **domElement** 等。

请注意，Web 部件类定义为接受 **IHelloWorldWebPartProps** 属性类型。

该属性类型在单独文件 **IHelloWorldWebPartProps.ts** 中被定义为一个接口。

```ts
export interface IHelloWorldWebPartProps {
    description: string;
}
```

此属性定义用于为 Web 部件定义自定义属性类型，稍后将在属性窗格部分进行介绍。 

#### <a name="web-part-render-method"></a>Web 部件呈现方法
呈现 Web 部件的 DOM 元素位于 **render** 方法中。此方法用于在该 DOM 元素内呈现 Web 部件。在 **HelloWorld** Web 部件中，DOM 元素设置为 DIV。方法参数包括显示模式（读取或编辑）和配置的 Web 部件属性（如果存在）： 

```ts
public render(): void {
  this.domElement.innerHTML = `
    <div class="${styles.helloWorld}">
      <div class="${styles.container}">
        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
          <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
            <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
            <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
            <p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
            <a href="https://github.com/SharePoint/sp-dev-docs/wiki" class="ms-Button ${styles.button}">
              <span class="ms-Button-label">Learn more</span>
            </a>
          </div>
        </div>
    </div>
  </div>`;
}
```

这种模型非常灵活，这样 Web 部件可内置到任意 JavaScript 框架并加载到 DOM 元素中。下面是如何加载 React 组件而不是纯 HTML 的示例。

```ts
render(): void {
    let e = React.createElement(TodoComponent, this.properties);
    ReactDom.render(e, this.domElement);
}
```

>**注意：**在项目中添加新 Web 部件时，Yeoman SharePoint 生成器允许你选择“**React**”作为所选框架。 

#### <a name="configure-the-web-part-property-pane"></a>配置 Web 部件属性窗格
属性窗格也在 **HelloWorldWebPart** 类中进行定义。需要在 **PropertyPaneSettings** 属性中定义属性窗格。

定义属性后，可以使用 **render** 方法中所示的 `this.properties.<property-value>` 在 Web 部件中访问它们：

```ts
<p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
```

阅读[将属性窗格与 Web 部件集成](../basics/integrate-with-property-pane)文章，了解如何使用属性窗格和属性窗格字段类型。

将 **propertyPaneSettings** 方法替换为显示如何添加 TextField 以外的属性类型的以下代码。 

```ts
protected get propertyPaneSettings(): IPropertyPaneSettings {
  return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
            PropertyPaneTextField('description', {
              label: 'Description'
            }),
            PropertyPaneTextField('test', {
              label: 'Multi-line Text Field',
              multiline: true
            }),
            PropertyPaneCheckbox('test1', {
              text: 'Checkbox'
            }),
            PropertyPaneDropdown('test2', {
              label: 'Dropdown',
              options: [
                { key: '1', text: 'One' },
                { key: '2', text: 'Two' },
                { key: '3', text: 'Three' },
                { key: '4', text: 'Four' }
              ]}),
            PropertyPaneToggle('test3', {
              label: 'Toggle',
              onText: 'On',
              offText: 'Off'
            })
          ]
          }
        ]
      }
    ]
  };
}
```

由于我们添加了新的属性字段，让我们从框架导入这些字段。

滚动到文件顶部，并从 `@microsoft/sp-client-preview` 将以下内容添加到导入部分：

```ts
PropertyPaneCheckbox,
PropertyPaneDropdown,
PropertyPaneToggle
```

完整的导入部分将如下所示：

```ts
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-client-preview';
```

保存该文件。

现在将这些属性添加到映射到我们刚刚添加的字段的 **IHelloWorldWebPartProps** 接口。

打开 **IHelloWorldWebPartProps.ts** 并将现有代码替换为以下代码：

```ts
export interface IHelloWorldWebPartProps {
    description: string;
    test: string;
    test1: boolean;
    test2: string;
    test3: boolean;
}
```

保存该文件。

切换回 **HelloWorldWebPart.ts** 文件。

将属性添加到 Web 部件属性后，你能以之前访问 **description** 属性的方式来访问此属性：

```ts
<p class="ms-font-l ms-fontColor-white">${this.properties.test2}</p>
```

若要设置该属性的默认值，则需要更新 Web 部件清单的 **properties** 属性包：

打开 `HelloWorldWebPart.manifest.json`，然后将 `properties` 修改为：

```ts
"properties": {
  "description": "HelloWorld",
  "test": "Multi-line text field",
  "test1": true,
  "test2": "2",
  "test3": true
}
```

### <a name="web-part-manifest"></a>Web 部件清单
**HelloWorldWebPart.manifest.json** 文件定义了 Web 部件元数据，如版本、ID、显示名称、图标和说明。每个 Web 部件均应包含此清单。

```ts
{
  "$schema": "../../../node_modules/@microsoft/sp-module-interfaces/lib/manifestSchemas/jsonSchemas/clientSideComponentManifestSchema.json",
  "id": "2f1ef3e2-5f37-4da8-b1ff-52c468a28ae2",
  "componentType": "WebPart",
  "version": "0.0.1",
  "manifestVersion": 2,
  "preconfiguredEntries": [
    {
      "groupId": "2f1ef3e2-5f37-4da8-b1ff-52c468a28ae2",
      "group": {
        "default": "Under Development"
      },
      "title": {
        "default": "HelloWorld"
      },
      "description": {
        "default": "HelloWorld description"
      },
      "officeFabricIconFontName": "Page",
      "properties": {
        "description": "HelloWorld",
        "test": "Multi-line text field",
        "test1": true,
        "test2": "2",
        "test3": true
      }
    }
  ]
}
```
### <a name="preview-the-web-part-in-sharepoint"></a>在 SharePoint 中预览 Web 部件

SharePoint 工作台还托管于 SharePoint 中，以预览和测试开发中的本地 Web 部件。主要优点是，在 SharePoint 上下文中运行的同时，还可以与 SharePoint 数据进行交互。

请转到以下 URL：”https://your-sharepoint-site/Shared%20Documents/workbench.aspx”

默认情况下，浏览器配置为不从 localhost 加载脚本。如果进行了此项配置，工作台会通知你。

![加载不安全脚本以从 localhost 运行脚本](../../../../images/sp-workbench-o365-unsface-scripts.png) 

要执行本地脚本，需要将浏览器配置为从未经身份验证的源加载脚本。这是因为在通过 HTTPS 连接到页面时，却通过 HTTP 加载脚本。根据所使用的浏览器，启用此功能的选项可能会有所不同。例如，在 Chrome 浏览器中，可以选择地址栏右侧的灰色盾牌来加载不安全脚本。 

![允许浏览器加载不安全脚本，以从 localhost 运行脚本](../../../../images/chrome-load-unsafe-scripts.png)

启用加载脚本后，会看到工作台加载。将 hello world Web 部件添加到画布：

![运行在 SharePoint Online 站点中的 SharePoint 工作台](../../../../images/sp-workbench-o365.png)

请注意，SharePoint 工作台现在具有 Office 365 套件导航栏。

选择画布上的**添加图标**以显示工具箱。工具箱现在显示托管 SharePoint 工作台与 **HelloWorldWebPart** 的站点上可用的 Web 部件。

![运行于 SharePoint Online 站点的 SharePoint 工作台中的工具箱](../../../../images/sp-workbench-o365-toolbox.png)

从工具箱添加 **HelloWorldWebPart**。现在，你即将在托管于 SharePoint 的页面中运行 Web 部件！

![在运行于 SharePoint Online 站点的 SharePoint 工作台中运行的 HelloWorld Web 部件](../../../../images/sp-workbench-o365-helloworld-wp.png)

由于 Web 部件仍处于开发和测试阶段，无需将 Web 部件打包并部署到 SharePoint。 

## <a name="next-steps"></a>后续步骤
祝贺你的第一个 Hello World Web 部件成功运行！既然 Web 部件已在运行，在下一主题[连接到 SharePoint](./connect-to-sharepoint) 中，你可以继续生成 Hello World Web 部件。你将使用相同的 Hello World Web 部件项目，并添加与 SharePoint 列表 REST API 进行交互的功能。请注意，`gulp serve` 命令仍在控制台窗口或 Visual Studio Code（如果使用 Visual Studio Code 编辑器）中运行。在你浏览下一篇文章时，可以继续让它运行。
