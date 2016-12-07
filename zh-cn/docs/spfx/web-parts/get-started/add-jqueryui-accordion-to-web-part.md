# <a name="add-jqueryui-accordion-to-your-sharepoint-client-side-web-part"></a>将 jQueryUI Accordion 添加到 SharePoint 客户端 Web 部件

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

本文介绍如何将 jQueryUI Accordion 添加到 Web 部件项目。这包括创建新的 Web 部件，如下图中所示。 

![包含 jQuery Accordion 的 Web 部件的屏幕截图](../../../../images/jquery-accordion-wb.png)

也可以通过观看 [SharePoint PnP YouTube 频道](https://www.youtube.com/watch?v=YcECe5JbAnA&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq)上的视频，按照这些步骤进行操作。 

<a href="https://www.youtube.com/watch?v=YcECe5JbAnA&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq">
<img src="../../../../images/spfx-youtube-tutorial5.png" alt="Screenshot of the YouTube video player for this tutorial" />
</a>

## <a name="prerequisites"></a>先决条件
开始之前，请完成下面的步骤：

* [构建第一个 Web 部件](build-a-hello-world-web-part.md)
* [连接到 SharePoint](connect-to-sharepoint.md) 

开发人员工具链使用 Webpack、SystemJS 和 CommonJS 捆绑 Web 部件。这包括加载任何外部依赖项（如 jQuery 或 jQueryUI）。要在高级别加载外部依赖项，将需要：

* 获取外部库，方法是通过 npm 或从供应商处下载。
* 如果可用，请安装各自框架的 [TypeScript 类型定义](http://definitelytyped.org/)。
* 如果需要，请更新你的解决方案配置以使 Web 部件包中不包含外部依赖项。

## <a name="create-a-new-web-part-project"></a>创建新的 Web 部件项目

在你最喜爱的位置创建新的项目目录：

```
md jquery-webpart
```
    
> **警告：**请务必在新文件夹中创建此目录，不是作为 `helloworld-webpart` 的子目录。

转到项目目录：

```
cd jquery-webpart
```
    
通过运行 Yeoman SharePoint 生成器创建新的 jQuery Web 部件：

```
yo @microsoft/sharepoint
```

出现提示时：

* 接受默认的 **jquery-webpart** 作为你的解决方案名称，并选择 **Enter**”
* 选择“**使用当前文件夹**”作为放置文件的位置。

接下来的一组提示会要求提供有关你的 Web 部件的特定信息：

* 对 Web 部件名称键入 **jQuery**，并选择 **Enter**。
* 输入“**jQuery Web 部件**”作为 Web 部件的描述，然后选择 **Enter**。
* 接受框架的默认设置“**无 javascript web 框架**”选项并选择 **Enter** 以继续。
* 出现下一个系统提示时，选择 **Enter** 以继续。请不要选择任何要添加的库。 

在这种情况下，Yeoman 将安装必需的依赖项并为解决方案文件提供基架。这可能需要几分钟的时间。Yeoman 将为项目提供基架，以同时包括 **jQueryWebPart** Web 部件。

在控制台中键入以下命令，以在 Visual Studio 代码中打开 Web 部件项目：

```
code .
```

## <a name="install-jquery-and-jquery-ui-npm-packages"></a>安装 JQuery 和 jQuery UI NPM 包

在控制台中，键入以下命令以安装 jQuery npm 包：

```
npm i --save jquery
```

 现在键入以下命令以安装 jQueryUI npm 包：

```
npm i --save jqueryui
```

TypeScript 定义管理器 (TSD) 可以搜索并安装项目的类型定义。使用 TSD 安装 jQuery 和 jQuery UI 的类型定义。

由于该 Web 部件项目主要是一个 TypeScript 项目，TypeScript 编译器必须能够了解各自的类型。它还可以帮助提供代码编辑器中所需的 IntelliSense。

要做到这一点，首先安装 [TypeScript 定义管理器](http://definitelytyped.org/tsd/)。

打开控制台并运行命令：

```
npm i -g tsd
```

现在安装 jQuery 和 jQuery UI 类型定义：

```
tsd install jquery jqueryui --save
```

TSD 将类型定义安装到 **/typings** 文件夹。此文件夹还包括其他键入内容（已通过 Yeoman 生成器提供基架）。

转到 **/typings/jquery** 和 **/typings/jqueryui** 以找到 jQuery 和 jQuery UI 的类型定义。

包括一个其他键入内容：

```
tsd install combokeys --save
```

### <a name="unbundle-external-dependencies-from-web-part-bundle"></a>从 Web 部件捆绑包中解除外部依赖项
默认情况下，你添加的任何依赖项均将被捆绑到 Web 部件捆绑包。在某些情况下，这并不理想。你可以选择从 Web 部件捆绑包中解除这些依赖项。

在 Visual Studio Code 中，打开文件 config\config.json。

此文件包含有关捆绑包和任何外部依赖项的信息。 

`entries` 区域包含默认捆绑包信息 - 在这种情况下是 jQuery Web 部件捆绑包。当向解决方案中添加更多 Web 部件时，将看到每个 Web 部件对应一项。

```json
"entries": [
  {
    "entry": "./lib/webparts/jQuery/jQueryWebPart.js",
    "manifest": "./src/webparts/jQuery/jQueryWebPart.manifest.json",
    "outputPath": "./dist/j-query.bundle.js",
  }
]
```

`externals` 部分包含未与默认捆绑包绑定的库。 

```json
  "externals": {
    "@microsoft/sp-client-base": "node_modules/@microsoft/sp-client-base/dist/sp-client-base.js",
    "@microsoft/sp-client-preview": "node_modules/@microsoft/sp-client-preview/dist/sp-client-preview.js",
    "@microsoft/sp-lodash-subset": "node_modules/@microsoft/sp-lodash-subset/dist/sp-lodash-subset.js",
    "office-ui-fabric-react": "node_modules/office-ui-fabric-react/dist/office-ui-fabric-react.js",
    "react": "node_modules/react/dist/react.min.js",
    "react-dom": "node_modules/react-dom/dist/react-dom.min.js",
    "react-dom/server": "node_modules/react-dom/dist/react-dom-server.min.js"
  },
```

要从默认捆绑包中排除 `jQuery` 和 `jQueryUI`，请将模块添加到 `externals` 部分：

```json
"jquery":"node_modules/jquery/dist/jquery.min.js",
"jqueryui":"node_modules/jqueryui/jquery-ui.min.js"
```

现在，构建项目时，`jQuery` 和 `jQueryUI` 不会捆绑到默认 Web 部件捆绑包。

## <a name="build-the-accordion"></a>构建可折叠项

在 Visual Studio Code 中打开项目文件夹 **jquery-webpart**。在 /src/webparts/jQuery 文件夹下，你的项目应有先前添加的 jQuery web 部件。

### <a name="add-accordion-html"></a>添加可折叠项 HTML
在名为 **MyAccordionTemplate.ts** 的 `src/webparts/jQuery` 文件夹中添加一个新文件。

创建并导出（作为一个模块）类 `MyAccordionTemplate`，它包含可折叠项的 HTML 代码。

```ts
export default class MyAccordionTemplate {
    public static templateHtml: string =  `
      <div class="accordion">
        <h3>Section 1</h3>
        <div>
            <p>
            Mauris mauris ante, blandit et, ultrices a, suscipit eget, quam. Integer
            ut neque. Vivamus nisi metus, molestie vel, gravida in, condimentum sit
            amet, nunc. Nam a nibh. Donec suscipit eros. Nam mi. Proin viverra leo ut
            odio. Curabitur malesuada. Vestibulum a velit eu ante scelerisque vulputate.
            </p>
        </div>
        <h3>Section 2</h3>
        <div>
            <p>
            Sed non urna. Donec et ante. Phasellus eu ligula. Vestibulum sit amet
            purus. Vivamus hendrerit, dolor at aliquet laoreet, mauris turpis porttitor
            velit, faucibus interdum tellus libero ac justo. Vivamus non quam. In
            suscipit faucibus urna.
            </p>
        </div>
        <h3>Section 3</h3>
        <div>
            <p>
            Nam enim risus, molestie et, porta ac, aliquam ac, risus. Quisque lobortis.
            Phasellus pellentesque purus in massa. Aenean in pede. Phasellus ac libero
            ac tellus pellentesque semper. Sed ac felis. Sed commodo, magna quis
            lacinia ornare, quam ante aliquam nisi, eu iaculis leo purus venenatis dui.
            </p>
            <ul>
            <li>List item one</li>
            <li>List item two</li>
            <li>List item three</li>
            </ul>
        </div>
        <h3>Section 4</h3>
        <div>
            <p>
            Cras dictum. Pellentesque habitant morbi tristique senectus et netus
            et malesuada fames ac turpis egestas. Vestibulum ante ipsum primis in
            faucibus orci luctus et ultrices posuere cubilia Curae; Aenean lacinia
            mauris vel est.
            </p>
            <p>
            Suspendisse eu nisl. Nullam ut libero. Integer dignissim consequat lectus.
            Class aptent taciti sociosqu ad litora torquent per conubia nostra, per
            inceptos himenaeos.
            </p>
        </div>
     </div>`;
}
```

保存该文件。

### <a name="import-accordion-html"></a>导入可折叠项 HTML

在 Visual Studio Code 中，打开 **src\webparts\jQuery\JQueryWebPart.ts**。

在文件顶部，可以找到其他导入，添加下面的导入：

```ts
import MyAccordionTemplate from './MyAccordionTemplate';
```

### <a name="import-jquery-and-jqueryui"></a>导入 jQuery 和 jQueryUI
可以使用与导入 MyAccordionTemplate 相同的方式来导入 jQuery Web 部件。

在文件顶部，可以找到其他导入，添加下面的导入：

```
import * as myjQuery from 'jquery';
```

请注意如何jQuery 模块导入一个名为 `myjQuery` 的自定义变量的方式。你可以随时调用它；我们建议你使用与 Web 部件相关的名称。 

由于 jQueryUI 是一个插件，所以要通过 require（而非 import）语句来加载它。在 JQuery 导入正下方，添加以下 require 语句：

```ts
require('jqueryui');
```

接下来，将加载某些外部 css 文件。为此，请使用模块加载程序。添加下列导入：

```ts
import importableModuleLoader from '@microsoft/sp-module-loader';
```

若要加载 jQueryUI 样式，请在 `JQueryWebPart` Web 部件类中，在构造函数内添加以下行：

```ts
importableModuleLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
```

构造函数现在如下所示：

```ts
public constructor(context: IWebPartContext){
    super(context);
    
    importableModuleLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
}
```

此代码将执行以下操作：

* 调用上下文父构造函数来初始化 Web 部件。
* 从 CDN 异步加载可折叠项样式。

### <a name="render-accordion"></a>渲染可折叠项

在 `jQueryWebPart.ts` 中，转至 `render` 方法。

设置 Web 部件的内联 HTML 以渲染可折叠项 HTML：

```ts
this.domElement.innerHTML = MyAccordionTemplate.templateHtml;
```

jQueryUI Accordion 有几个选项，可以设置为自定义可折叠项。在 `this.domElement.innerHTML = MyAccordionTemplate.templateHtml;` 正下方定义可折叠项的几个选项：

```ts
const accordionOptions: JQueryUI.AccordionOptions = {
  animate: true,
  collapsible: false,
  icons: {
    header: 'ui-icon-circle-arrow-e',
    activeHeader: 'ui-icon-circle-arrow-s'
  }
};
```

如你所见，jQueryUI 类型化定义使你可以创建名为 `JQueryUI.AccordionOptions` 的类型化变量，并指定所支持的属性。 

如果你在使用 IntelliSense，你会注意到你将在 `JQueryUI.` 下获得可用方法及方法参数的完全支持。

最后，初始化可折叠项：

```ts
myjQuery(this.domElement).children('.accordion').accordion(accordionOptions);
```

正如你所看到的，使用变量 `myjQuery`（用于导入 `jquery` 模块）。然后，初始化可折叠项。

完整的 `render` 方法类如下所示：

```ts
public render(): void {
  this.domElement.innerHTML = MyAccordionTemplate.templateHtml;

  const accordionOptions: JQueryUI.AccordionOptions = {
    animate: true,
    collapsible: false,
    icons: {
      header: 'ui-icon-circle-arrow-e',
      activeHeader: 'ui-icon-circle-arrow-s'
    }
  };

  myjQuery(this.domElement).children('.accordion').accordion(accordionOptions);
}
```

保存该文件。

## <a name="preview-the-web-part"></a>预览 Web 部件

在控制台中，请确保你仍在 jquery-webpart 文件夹中，并键入以下内容以构建和预览 Web 部件：

```
gulp serve
```

> **注意：**Visual Studio Code 为 gulp 和其他任务运行程序提供内置支持。你可以使用 Windows 中的 **Ctrl+Shift+B** 或 Mac 上的 **Cmd+Shift+B** 进行调试并预览 Web 部件。

Gulp 将执行任务并打开本地 SharePoint Web 部件工作台。

在页面画布中，选择“**+**”（加号）以显示 Web 部件列表，然后添加 jQuery Web 部件。你现在应可以看到 jQueryUI Accordion！

![包含 jQuery Accordion 的 Web 部件的屏幕截图](../../../../images/jquery-accordion-wb.png)

在运行 `gulp serve` 的控制台中，选择 **Ctrl + C** 以终止任务。
