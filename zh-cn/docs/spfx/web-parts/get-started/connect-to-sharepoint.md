# <a name="connect-your-client-side-web-part-to-sharepoint-hello-world-part-2"></a>将客户端 Web 部件连接到 SharePoint（Hello world 第 2 部分）

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

将 Web 部连接到 SharePoint 以访问 SharePoint 中的功能和数据并为最终用户提供更全面的经验。本文继续构建前一篇文章[构建第一个 Web 部件](./build-a-hello-world-web-part)中构建的 hello world web 部件。

也可以通过观看 [SharePoint PnP YouTube 频道](https://www.youtube.com/watch?v=rokWJlXoFWk&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq)上的视频，按照这些步骤进行操作。 

<a href="https://www.youtube.com/watch?v=rokWJlXoFWk&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq">
<img src="../../../../images/spfx-youtube-tutorial2.png" alt="Screenshot of the YouTube video player for this tutorial" />
</a>


## <a name="run-gulp-serve"></a>运行 gulp 服务

请确保 `gulp serve` 命令运行正常。如果尚未运行，请转到 **helloworld-webpart** 项目目录，并使用下面的命令运行。

```
cd helloworld-webpart
gulp serve
```

## <a name="get-access-to-page-context"></a>获取页面上下文的访问权限

在本地托管工作台时，没有 SharePoint 页面上下文。你仍然可以使用许多不同的方法来测试 Web 部件。例如，可以将精力集中在构建 Web 部件的用户体验上，并在没有 SharePoint 上下文时使用虚假数据模拟 SharePoint 交互。

但是，在 SharePoint 中托管工作台时，你有权访问页面上下文，它可提供各种关键属性，如：

* Web 标题
* Web 绝对 URL
* Web 服务器相对 URL
* 用户登录名

通过在 Web 部件类中使用以下变量可获取页面上下文的访问权限：

```ts
this.context.pageContext
```

切换到 Visual Studio Code（或你的首选 IDE），并打开 **src\webparts\helloWorld\HelloWorldWebPart.ts**。

在“**渲染**”方法内，将 **innerHTML** 代码块替换为下面的代码：

```ts
this.domElement.innerHTML = `
<div class='${styles.helloWorld}'>
    <div class='${styles.container}'>
        <div class='ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}'>
        <div class='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
            <span class='ms-font-xl ms-fontColor-white'>Welcome to SharePoint!</span>
            <p class='ms-font-l ms-fontColor-white'>Customize SharePoint experiences using Web Parts.</p>
            <p class='ms-font-l ms-fontColor-white'>${this.properties.description}</p>
            <p class='ms-font-l ms-fontColor-white'>Loading from ${this.context.pageContext.web.title}</p>
            <a href='https://github.com/SharePoint/sp-dev-docs/wiki' class='ms-Button ${styles.button}'>
                <span class='ms-Button-label'>Learn more</span>
            </a>
        </div>
        </div>
</div>
</div>`;
```

请注意如何在 HTML 块中使用 `${ }` 输出变量的值。使用额外 HTML 代码 `p` 来显示 `this.context.pageContext.web.title`。由于此 Web 部件从本地环境加载，标题将是“**本地工作台**”。

保存该文件。在控制台中运行 `gulp serve` 会检测此保存操作并：

* 自动构建和捆绑更新的代码。
* 刷新本地工作台页面（因为需要重新加载 Web 部件代码）。

>**注意：**使控制台窗口和 VS Code 并排，以便在 VS Code 中保存更改时自动查看 gulp 编译。

在浏览器中，转到本地 **workbench.html**。如果已经关闭选项卡，该 URL 是 https://localhost:4321/temp/workbench.html。

你应该会看到以下 Web 部件：

![localhost 中的 SharePoint 页面上下文](../../../../images/sp-mock-localhost-wp.png)

转到 SharePoint 中托管的 **workbench.aspx**。完整的 URL 是 https://your-sharepoint-site-url/Shared%20Documents/workbench.aspx。

>**注意：**如果没有安装 SPFx 开发人员证书，则工作台将通知你，指出它被配置为不从 localhost 加载脚本。在项目目录控制台中执行 `gulp trust-dev-cert` 命令以安装开发人员证书。

你现在应该在 Web 部件中看到 SharePoint 网站 URL，现在该页面上下文适用于 Web 部件。

![SharePoint 网站中的 SharePoint 页面上下文](../../../../images/sp-lists-spsiteurl-wp.png)

## <a name="define-list-model"></a>定义列表模型
你需要列表模型才可开始使用 SharePoint 列表数据。要检索列表，你需要两个模型。 

切换到 Visual Studio Code 并打开 **src\webparts\helloWorld\HelloWorldWebPart.ts**。

定义以下 `interface` 模型（**HelloWorldWebPart** 类正上方）：

```ts
export interface ISPLists {
    value: ISPList[];
}

export interface ISPList {
    Title: string;
    Id: string;
}
```

**ISPList** 接口包含我们将连接到的 SharePoint 列表信息。 

## <a name="retrieve-lists-from-mock-store"></a>从虚假存储检索列表

要在本地工作台中测试，将需要返回虚假数据的虚假存储。

在名为 **MockHttpClient.ts** 的 **src\webparts\helloWorld** 文件夹中创建一个新文件。

将下面的代码复制到 **MockHttpClient.ts**：

```ts
import { ISPList } from './HelloWorldWebPart';

export default class MockHttpClient {

    private static _items: ISPList[] = [{ Title: 'Mock List', Id: '1' }];
    
    public static get(restUrl: string, options?: any): Promise<ISPList[]> {
    return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}
```

代码注意事项：

* 因为 **HelloWorldWebPart.ts** 中有多个导出，要使用 `{ }` 指定要导入的特定项。在此情况下，只需要数据模型 `ISPList`。
* 从默认模块导入（此情况下是 **HelloWorldWebPart**）时，不需要键入文件扩展名。 
* 它将 **MockHttpClient** 类导出为默认模块，以便可以在其他文件中导入。
* 它构建初始 `ISPList` 虚假数组并返回。

保存该文件。

现在，你可以在 **HelloWorldWebPart** 类中使用 **MockHttpClient** 类。首先，你需要导入 **MockHttpClient** 模块。

打开 **HelloWorldWebPart.ts** 文件。

将以下代码复制并粘贴到 `import { IHelloWorldWebPartProps } from './IHelloWorldWebPartProps';` 正下方。

```ts
import MockHttpClient from './MockHttpClient';
```
 
添加以下私有方法以模拟 **HelloWorldWebPart** 类中的列表检索。

```ts
private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl)
        .then((data: ISPList[]) => {
             var listData: ISPLists = { value: data };
             return listData;
         }) as Promise<ISPLists>;
}
```

保存该文件。

## <a name="retrieve-lists-from-sharepoint-site"></a>从 SharePoint 网站检索列表

接下来，你需要从当前网站检索列表。你将使用 SharePoint REST API 从网站检索列表，位于 https://yourtenantprefix.sharepoint.com/_api/web/lists。

添加以下私有方法以从 **HelloWorldWebPart** 类的 SharePoint 中检索列表。

```ts
private _getListData(): Promise<ISPLists> {
return this.context.httpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`)
        .then((response: Response) => {
        return response.json();
        });
}
```

前一种方法使用帮助程序类 **httpClient**（可从 SharePoint 客户端平台获取）来执行 REST API。它使用 **ISPLists** 模型，并应用筛选器以不检索隐藏的列表。

保存该文件。 

切换到运行 `gulp serve` 的控制台窗口，并检查是否有任何错误。如果有错误，gulp 会在控制台中进行报告，并需要在继续之前修复它们。

## <a name="add-new-styles"></a>添加新样式

SharePoint Framework 使用 [Sass](http://sass-lang.com/) 作为 CSS 预处理器并专门使用 [SCSS 语法](http://sass-lang.com/documentation/file.SCSS_FOR_SASS_USERS.html)（与普通 CSS 语法完全兼容)。Sass 扩展 CSS 语言并使你可以使用许多功能（如变量、嵌套规则和内联导入）来为 Web 部件组织和创建有效的样式表。SharePoint Framework 已附带 SCSS 编译器，可将 Sass 文件转换为普通 CSS 文件，还提供了类型化版本以在开发过程中使用它。

要添加新样式，请打开 **HelloWorld.module.scss**。这是你将在其中定义样式的 SCSS 文件。

默认情况下，样式作用于 Web 部件。在 **.helloWorld** 下定义样式时你可以看到。

在 `.button` 样式后添加以下样式：

```css
.list {
    color: #333333;
    font-family: 'Segoe UI Regular WestEuropean', 'Segoe UI', Tahoma, Arial, sans-serif;
    font-size: 14px;
    font-weight: normal;
    box-sizing: border-box;
    margin: 10;
    padding: 10;
    line-height: 50px;
    list-style-type: none;
    box-shadow: 0 4px 4px 0 rgba(0, 0, 0, 0.2), 0 25px 50px 0 rgba(0, 0, 0, 0.1);
}

.listItem {
    color: #333333;
    vertical-align: center;
    font-family: 'Segoe UI Regular WestEuropean', 'Segoe UI', Tahoma, Arial, sans-serif;
    font-size: 14px;
    font-weight: normal;
    box-sizing: border-box;
    margin: 0;
    padding: 0;
    box-shadow: none;
    *zoom: 1;
    padding: 9px 28px 3px;
    position: relative;
}
``` 

保存该文件。

一旦保存该文件，gulp 将在控制台中重新构建代码。这将在 **HelloWorld.module.scss.ts** 文件中生成相应的键入内容。编译到 typescript 后，可以在 Web 部件代码中导入并引用这些样式。

可以在 Web 部件的“**渲染**”方法中看到：

```html
<div class="${styles.container}">
```

## <a name="method-to-render-lists-information"></a>用于渲染列表信息的方法

打开 **HelloWorldWebPart** 类。

SharePoint 工作台使你可以灵活地在本地环境和从 SharePoint 网站中测试 Web 部件。SharePoint Framework 可使用 **EnvironmentType** 模块帮助你了解 Web 部件从哪个环境运行，以实现此功能。 

要使用该模块，你首先需要从 **@microsoft/sp-client-base** 捆绑包导入 **EnvironmentType** 模块。将其添加到顶部的“**导入**”部分，如以下代码所示：

```ts
import { EnvironmentType } from '@microsoft/sp-client-base';
```

在 **HelloWorldWebPart** 类中添加以下私有方法，以调用各自方法来检索列表。

```ts
private _renderListAsync(): void {
    // Local environment
    if (this.context.environment.type === EnvironmentType.Local) {
        this._getMockListData().then((response) => {
        this._renderList(response.value);
        }); }
        else {
        this._getListData()
        .then((response) => {
            this._renderList(response.value);
        });
    }
}
```

**_renderListAsync** 方法中 hostType 的注意事项：

* `this.context.environment.type` 属性可以帮助你检查你是位于本地还是 SharePoint 环境中。
* 根据托管工作台的位置，调用正确的方法。

保存该文件。

现在，你需要使用从 REST API 提取的值来渲染列表数据。

在 **HelloWorldWebPart** 类中添加以下私有方法：

```ts
private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
        html += `
        <ul class="${styles.list}">
            <li class="${styles.listItem}">
                <span class="ms-font-l">${item.Title}</span>
            </li>
        </ul>`;
    });

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
}
```

前一种方法使用 **styles** 变量引用先前添加的新 CSS 样式。 

保存该文件。

## <a name="retrieve-list-data"></a>检索列表数据

导航到 **render** 方法，并使用下面的代码替换方法内的代码：

```ts
this.domElement.innerHTML = `
<div class="${styles.helloWorld}">
    <div class="${styles.container}">
        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
                <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
                <p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
                <p class="ms-font-l ms-fontColor-white">${this.properties.test2}</p>
                <p class='ms-font-l ms-fontColor-white'>Loading from ${this.context.pageContext.web.title}</p>
                <a href="https://github.com/SharePoint/sp-dev-docs/wiki" class="ms-Button ${styles.button}">
                    <span class="ms-Button-label">Learn more</span>
                </a>
            </div>
        </div>
    <div id="spListContainer" />
    </div>
</div>`;

this._renderListAsync();
```

保存该文件。

请注意，在 `gulp serve` 控制台窗口中它会重新构建代码。请确保你看不到任何错误。

切换到本地工作台并添加 HelloWorld Web 部件。

你应该看到返回的虚假数据。

![从 localhost 渲染列表数据](../../../../images/sp-lists-render-localhost.png)

切换到 SharePoint 中托管的工作台。刷新页面，并添加 HelloWorld Web 部件。

你应该看到从当前网站返回的列表。

![从 SharePoint 渲染列表数据](../../../../images/sp-lists-render-spsite.png)

现在，你可以使服务器停止运行。切换到控制台，然后停止 `gulp serve`。选择 `Ctrl+C` 以终止 gulp 任务。

## <a name="next-steps"></a>后续步骤

祝贺你成功将 Web 部件连接到 SharePoint 列表数据！在下一主题[将 Web 部件部署到经典 SharePoint 页面](./serve-your-web-part-in-a-sharepoint-page)中，你可以继续构建出 Hello World Web 部件。你将学习如何在经典 SharePoint 服务器端页面中部署和预览 Hello World Web 部件。
