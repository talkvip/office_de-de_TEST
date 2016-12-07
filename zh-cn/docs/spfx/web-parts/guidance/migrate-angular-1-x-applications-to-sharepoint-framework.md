# <a name="migrate-angular-v1x-applications-to-sharepoint-framework"></a>将 Angular 1.x 应用程序迁移到 SharePoint Framework

许多组织过去已使用 Angular 生成 SharePoint 解决方案。本文介绍如何将使用 [ngOfficeUIFabric](http://ngofficeuifabric.com) - Office UI Fabric 的 Angular 指令样式化的现有 Angular v1.x 应用程序迁移到 SharePoint Framework 客户端 Web 部件。用于本教程的示例应用程序能够执行存储在 SharePoint 列表中的待执行项。

![用于管理存储在 SharePoint 列表中的待执行项的 Angular 应用程序](../../../../images/ng-migration-original-angular-application.png)

Angular 应用程序源可在 GitHub 上获取，网址为 [https://github.com/SharePoint/sp-dev-fx-webparts/tree/dev/samples/angular-migration/angular-todo](https://github.com/SharePoint/sp-dev-fx-webparts/tree/dev/samples/angular-migration/angular-todo)。

迁移到 SharePoint Framework 的 Angular 应用程序的源可在 GitHub 上获取，网址为 [https://github.com/SharePoint/sp-dev-fx-webparts/tree/dev/samples/angular-migration/angular-todo](https://github.com/SharePoint/sp-dev-fx-webparts/tree/master/samples/angular-todo-webpart)。

> **注意：**执行本文中的步骤之前，请确保[设置开发环境](http://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment)以便生成 SharePoint Framework 解决方案。

## <a name="setup-project"></a>安装项目

开始迁移 Angular 应用程序前，创建并安装新的 SharePoint Framework 项目，以托管 Angular 应用程序。

### <a name="create-new-project"></a>新建项目

首先为项目创建新文件夹

```sh
md angular-todo
```

导航到项目文件夹：

```sh
cd angular-todo
```

在项目文件夹中，运行 SharePoint Framework Yeoman 生成器，以搭建新的 SharePoint Framework 项目：

```sh
yo @microsoft/sharepoint
```

出现提示时，请按如下所述定义值：
- **angular-todo** 作为解决方案名称
- **使用当前文件夹**放置文件
- **To do** 作为 Web 部件名称
- **Simple management of to do tasks** 作为 Web 部件说明
- **No JavaScript web framework** 作为生成 Web 部件的起点

![带有默认选项的 SharePoint Framework Yeoman 生成器](../../../../images/ng-migration-yeoman-generator.png)

基架完成后，在代码编辑器中打开项目文件夹。在本教程中，你将使用 Visual Studio Code。

![在 Visual Studio Code 中打开 SharePoint Framework 项目](../../../../images/ng-migration-project-visual-studio-code.png)

### <a name="add-angular-and-ngofficeuifabric"></a>添加 Angular 和 ngOfficeUIFabric

在本教程中，你将从 CDN 加载 Angular 和 ngOfficeUIFabric。为此，请在代码编辑器中打开 **config/config.json** 文件并在 **externals** 属性中添加以下行：

```json
"angular": {
  "path": "https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.5.16/angular.min.js",
  "globalName": "angular"
},
"ng-office-ui-fabric": "https://cdnjs.cloudflare.com/ajax/libs/ngOfficeUiFabric/0.12.3/ngOfficeUiFabric.js"
```

![Angular 和 ngOfficeUIFabric 已添加到 config.json 文件](../../../../images/ng-migration-angular-ngofficeuifabric-config.png)

### <a name="add-angular-typings-for-typescript"></a>添加 TypeScript Angular 键入

由于要在 Web 部件的代码中引用 Angular，因此还需要 TypeScript 的 Angular 键入。若要安装它们，请在命令行中运行：

```sh
npm install @types/angular --save-dev
```

## <a name="migrate-the-angular-application-as-is"></a>按原样迁移 Angular 应用程序

首先迁移只有最少的代码更改的 Angular 应用程序。然后将应用程序的普通 JavaScript 代码升级到 TypeScript，并改进其与客户端 Web 部件的集成。

### <a name="create-sharepoint-list"></a>创建 SharePoint 列表

在 SharePoint 站点中，创建名称为“**Todo**”的新列表。在此列表中，添加名为“**Status**”的新选择列。作为可用选项，请输入：

```
Not started
In progress
Completed
```

![待执行 SharePoint 列表](../../../../images/ng-migration-todo-list.png)

### <a name="copy-angular-application-files-to-the-web-part-project"></a>将 Angular 应用程序文件复制到该 Web 部件项目

在 Web 部件项目中，在 **src/webparts/toDo** 文件夹中创建名为“`app`”的新文件夹。

![“App”文件夹在 Visual Studio Code 资源管理器窗格中突出显示](../../../../images/ng-migration-app-folder-visual-studio-code.png)

从源应用程序将“**app**”文件夹的内容复制到在 Web 部件项目中新建的“**app**”文件夹中。

![“App”的文件在 Visual Studio Code 资源管理器窗格中突出显示](../../../../images/ng-migration-app-files-visual-studio-code.png)

### <a name="load-the-angular-application-in-the-client-side-web-part"></a>在客户端 Web 部件中加载 Angular 应用程序

在代码编辑器中打开 **./src/webparts/toDo/ToDoWebPart.ts** 文件。

在最后一个 `import` 语句后添加以下代码：

```ts
import * as angular from 'angular';
import 'ng-office-ui-fabric';
```

![导入语句以加载 Angular 和 ngOfficeUIFabric](../../../../images/ng-migration-angular-ngofficeuifabric-import.png)

将 **render** 方法的内容更改为：

```ts
export default class ToDoWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {
  // ...
  public render(): void {
    if (this.renderedOnce === false) {
      require('./app/app.module');
      require('./app/app.config');
      require('./app/data.service');
      require('./app/home.controller');

      this.domElement.innerHTML = `
        <div class="${styles.toDo}">
          <div data-ng-controller="homeController as vm">
            <div class="${styles.loading}" ng-show="vm.isLoading">
              <uif-spinner>Loading...</uif-spinner>
            </div>
            <div class="entryform" ng-show="vm.isLoading === false">
              <uif-textfield uif-label="New to do:" uif-underlined ng-model="vm.newItem" ng-keydown="vm.todoKeyDown($event)"></uif-textfield>
            </div>
            <uif-list class="items" ng-show="vm.isLoading === false" >
              <uif-list-item ng-repeat="todo in vm.todoCollection" uif-item="todo" ng-class="{'done': todo.done}">
                <uif-list-item-primary-text>{{todo.title}}</uif-list-item-primary-text>
                <uif-list-item-actions>
                  <uif-list-item-action ng-click="vm.completeTodo(todo)" ng-show="todo.done === false">
                    <uif-icon uif-type="check"></uif-icon>
                  </uif-list-item-action>
                  <uif-list-item-action ng-click="vm.undoTodo(todo)" ng-show="todo.done">
                    <uif-icon uif-type="reactivate"></uif-icon>
                  </uif-list-item-action>
                  <uif-list-item-action ng-click="vm.deleteTodo(todo)">
                    <uif-icon uif-type="trash"></uif-icon>
                  </uif-list-item-action>
                </uif-list-item-actions>
              </uif-list-item>
            </uif-list>
          </div>
        </div>`;

      angular.bootstrap(this.domElement, ['todoapp']);
    }
  }
  // ...
}
```

### <a name="update-site-path"></a>更新站点路径

在代码编辑器中打开 **./src/webparts/toDo/app/app.config.js** 文件。将 **sharepointApi** 常量的值更改为 SharePoint 站点（在其上创建了 Todo 列表，后接 `/_api/`）的服务器相关的 URL。

![Web URL 在 Visual Studio Code 中的 app.config.js 文件中突出显示](../../../../images/ng-migration-app-config-web-url.png)

### <a name="trust-the-development-certificate"></a>信任开发证书

默认情况下，需要用于通过 HTTPS 加载 SharePoint 工作台及其资源的开发证书不受信任，并使 Web 浏览器在浏览至 SharePoint 工作台时显示警告。在需要在 SharePoint 上下文中运行 SharePoint 工作台的情况下，如果 SSL 证书不受信任，则某些 Web 浏览器禁止加载工作台。若要避免此问题，应该信任通过 SharePoint Framework 提供的开发证书。

在命令行中执行：

```sh
gulp trust-dev-cert
```

### <a name="upload-workbenchaspx-to-sharepoint"></a>将 workbench.aspx 上载到 SharePoint

在命令行中执行：

```sh
gulp serve --nobrowser
```

在 **./temp** 文件夹中，复制“**workbench.html**”文件并将其重命名为“**workbench.aspx**”。在 **workbench.aspx** 文件中，将 **webAbsoluteUrl** 属性的值更改为在其中创建 Todo 列表的 SharePoint 站点的完整 URL。

![Web URL 在 Visual Studio Code 中的 workbench.aspx 文件中突出显示](../../../../images/ng-migration-workbench-weburl.png)

将 **workbench.aspx** 文件上载到 SharePoint 站点中的文档库。

![Web URL 在 Visual Studio Code 中的 workbench.aspx 文件中突出显示](../../../../images/ng-migration-workbench-aspx-doclib.png)

如果正确执行了所有步骤，则应该可以在浏览器中看到显示要添加到待执行项的表单的 Web 部件。

![显示在 SharePoint 工作台中的已迁移 Angular 应用程序已上载到 SharePoint](../../../../images/ng-migration-first-run.png)

添加几个待执行项来验证 Web 部件按预期方式工作。

![已迁移的 Angular 应用程序样式不正确](../../../../images/ng-migration-old-office-ui-fabric.png)

### <a name="fix-web-part-styling"></a>修复 Web 部件样式设置

虽然 Web 部件工作正常，但它看起来与开始使用的 Angular 应用程序不相同。这是由于 ngOfficeUIFabric 使用的 Office UI Fabric 比 SharePoint 工作台中的可用版本要低。简单的解决方法是加载 ngOfficeUIFabric 所使用的 CSS 样式。这样做会产生的问题是，这些样式将与 SharePoint 工作台所使用的 Office UI Fabric 样式发生冲突，从而破坏其用户界面。更好的解决方案是将特定组件所需的样式添加到 Web 部件样式。

在代码编辑器中打开 **./src/webparts/toDo/ToDo.module.scss** 文件。将其内容更改为：

```css
.toDo {
  .loading {
    margin: 0 auto;
    width: 6em;
  }

  .done :global .ms-ListItem-primaryText {
    text-decoration: line-through;
  }

  ul, li {
    margin: 0;
    padding: 0;
  }

  :global {
  .ms-ListItem{font-family:"Segoe UI WestEuropean","Segoe UI",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",sans-serif;-webkit-font-smoothing:antialiased;font-size:14px;font-weight:400;box-sizing:border-box;margin:0;padding:0;box-shadow:none;padding:9px 28px 3px;position:relative;display:block}.ms-ListItem::after,.ms-ListItem::before{display:table;content:"";line-height:0}.ms-ListItem::after{clear:both}.ms-ListItem-primaryText,.ms-ListItem-secondaryText,.ms-ListItem-tertiaryText{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:block}.ms-ListItem-primaryText{font-family:"Segoe UI WestEuropean","Segoe UI",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",sans-serif;-webkit-font-smoothing:antialiased;font-size:21px;font-weight:100;padding-right:80px;position:relative;top:-4px}.ms-ListItem-secondaryText{font-family:"Segoe UI WestEuropean","Segoe UI",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",sans-serif;-webkit-font-smoothing:antialiased;font-size:14px;font-weight:400;line-height:25px;position:relative;top:-7px;padding-right:30px}.ms-ListItem-tertiaryText{font-family:"Segoe UI WestEuropean","Segoe UI",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",sans-serif;-webkit-font-smoothing:antialiased;font-size:14px;font-weight:400;position:relative;top:-9px;margin-bottom:-4px;padding-right:30px}.ms-ListItem-metaText{font-family:"Segoe UI WestEuropean","Segoe UI",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",sans-serif;-webkit-font-smoothing:antialiased;font-size:11px;font-weight:400;position:absolute;right:30px;top:39px}.ms-ListItem-image{float:left;height:70px;margin-left:-8px;margin-right:10px;width:70px}.ms-ListItem-selectionTarget{display:none}.ms-ListItem-actions{max-width:80px;position:absolute;right:30px;text-align:right;top:10px}.ms-ListItem-action{color:#a6a6a6;display:inline-block;font-size:15px;position:relative;text-align:center;top:3px;cursor:pointer;height:16px;width:16px}.ms-ListItem-action .ms-Icon{vertical-align:top}.ms-ListItem-action:hover{color:#666666;outline:1px solid transparent}.ms-ListItem.is-unread{border-left:3px solid #0078d7;padding-left:27px}.ms-ListItem.is-unread .ms-ListItem-metaText,.ms-ListItem.is-unread .ms-ListItem-secondaryText{color:#0078d7;font-weight:600}.ms-ListItem.is-unseen:after{border-right:10px solid transparent;border-top:10px solid #0078d7;left:0;position:absolute;top:0}.ms-ListItem.is-selectable .ms-ListItem-selectionTarget{display:block;height:20px;left:6px;position:absolute;top:13px;width:20px}.ms-ListItem.is-selectable .ms-ListItem-image{margin-left:0}.ms-ListItem.is-selectable:hover{background-color:#eaeaea;cursor:pointer;outline:1px solid transparent}.ms-ListItem.is-selectable:hover:before{-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;speak:none;position:absolute;top:12px;left:6px;height:15px;width:15px;border:1px solid #767676}.ms-ListItem.is-selected:before{border:1px solid transparent}.ms-ListItem.is-selected:before,.ms-ListItem.is-selected:hover:before{-moz-osx-font-smoothing:grayscale;-webkit-font-smoothing:antialiased;display:inline-block;font-family:FabricMDL2Icons;font-style:normal;font-weight:400;speak:none;content:'\e041';font-size:15px;color:#767676;position:absolute;top:12px;left:6px}.ms-ListItem.is-selected:hover{background-color:#c7e0f4;outline:1px solid transparent}.ms-ListItem.ms-ListItem--document{padding:0}.ms-ListItem.ms-ListItem--document .ms-ListItem-itemIcon{width:70px;height:70px;float:left;text-align:center}.ms-ListItem.ms-ListItem--document .ms-ListItem-itemIcon .ms-Icon{font-size:38px;line-height:70px;color:#666666}.ms-ListItem.ms-ListItem--document .ms-ListItem-primaryText{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:14px;padding-top:15px;padding-right:0;position:static}.ms-ListItem.ms-ListItem--document .ms-ListItem-secondaryText{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-family:"Segoe UI WestEuropean","Segoe UI",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",sans-serif;-webkit-font-smoothing:antialiased;font-size:11px;font-weight:400;padding-top:6px}.MailList{overflow-y:auto;-webkit-overflow-scrolling:touch;max-height:500px}.MailTile{margin-bottom:5px;padding:10px;background:red}
  }
}
```

在 **./src/webparts/toDo/ToDoWebPart.ts** 文件中，在 **render** 方法中，将应用程序呈现模板更改为使用新的 Office UI Fabric 图标。

```ts
export default class ToDoWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {
  // ...
  public render(): void {
    if (this.renderedOnce === false) {
      require('./app/app.module');
      require('./app/app.config');
      require('./app/data.service');
      require('./app/home.controller');

      this.domElement.innerHTML = `
        <div class="${styles.toDo}">
          <div data-ng-controller="HomeController as vm">
            <div id="loading" ng-show="vm.isLoading">
              <uif-spinner>Loading...</uif-spinner>
            </div>
            <div id="entryform" ng-show="vm.isLoading === false">
              <uif-textfield uif-label="New to do:" uif-underlined ng-model="vm.newItem" ng-keydown="vm.todoKeyDown($event)"></uif-textfield>
            </div>
            <uif-list id="items" ng-show="vm.isLoading === false" >
              <uif-list-item ng-repeat="todo in vm.todoCollection" uif-item="todo" ng-class="{'${styles.done}': todo.done}">
                <uif-list-item-primary-text>{{todo.title}}</uif-list-item-primary-text>
                <uif-list-item-actions>
                  <uif-list-item-action ng-click="vm.completeTodo(todo)" ng-show="todo.done === false">
                    <i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i>
                  </uif-list-item-action>
                  <uif-list-item-action ng-click="vm.undoTodo(todo)" ng-show="todo.done">
                    <i class="ms-Icon ms-Icon--RevToggleKey" aria-hidden="true"></i>
                  </uif-list-item-action>
                  <uif-list-item-action ng-click="vm.deleteTodo(todo)">
                    <i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i>
                  </uif-list-item-action>
                </uif-list-item-actions>
              </uif-list-item>
            </uif-list>
          </div>
        </div>`;

      angular.bootstrap(this.domElement, ['todoapp']);
    }
  }
  // ...
}
```

刷新 Web 浏览器中的 Web 部件时，可以看到样式已正确设置。

![迁移的 Angular 应用程序在 Web 部件中正确标记为已完成项](../../../../images/ng-migration-correct-styling.png)

## <a name="upgrade-the-angular-application-to-typescript"></a>将 Angular 应用程序升级到 TypeScript

原始 Angular 应用程序以普通 JavaScript 编写，这使得对其进行维护时容易出错。生成 SharePoint Framework 客户端 Web 部件时，可以使用 TypeScript，并可得益于其设计时类型安全功能。在本教程的此部分，可以将普通 JavaScript Angular代码迁移到 TypeScript。

### <a name="upgrade-application-configuration"></a>升级应用程序配置

在项目中，将 **./src/webparts/toDo/app/app.config.js** 文件重命名为 `app.config.ts`。将其内容更改为：

```ts
import * as angular from 'angular';

export default function() {
  const todoapp: ng.IModule = angular.module('todoapp');
  todoapp.constant('sharepointApi', '/todo/_api/');
  todoapp.constant('todoListName', 'Todo');
  todoapp.constant('hideFinishedTasks', false);
}
```

### <a name="upgrade-data-service"></a>升级数据服务

在项目中，将 **./src/webparts/toDo/app/data.service.js** 文件重命名为 `DataService.ts`。将其内容更改为：

```ts
import * as angular from 'angular';

export interface ITodo {
  id: number;
  title: string;
  done: boolean;
}

interface ITodoItem {
  Id: number;
  Title: string;
  Status: string;
}

export interface IDataService {
  getTodos: () => angular.IPromise<ITodo[]>;
  addTodo: (todo: string) => angular.IPromise<{}>;
  deleteTodo: (todo: ITodo) => angular.IPromise<{}>;
  setTodoStatus: (todo: ITodo, done: boolean) => angular.IPromise<{}>;
}

export default class DataService implements IDataService {
  public static $inject: string[] = ['$q', '$http', 'sharepointApi', 'todoListName', 'hideFinishedTasks'];

  constructor(private $q: angular.IQService,
    private $http: angular.IHttpService,
    private sharepointApi: string,
    private todoListName: string,
    private hideFinishedTasks: boolean) {
  }

  public getTodos(): angular.IPromise<ITodo[]> {
    const deferred: angular.IDeferred<ITodo[]> = this.$q.defer();

    let url: string = `${this.sharepointApi}web/lists/getbytitle('${this.todoListName}')/items?$select=Id,Title,Status&$orderby=ID desc`;

    if (this.hideFinishedTasks === true) {
      url += "&$filter=Status ne 'Completed'";
    }

    this.$http({
      url: url,
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((result: angular.IHttpPromiseCallbackArg<{ value: ITodoItem[] }>): void => {
      const todos: ITodo[] = [];
      for (let i: number = 0; i < result.data.value.length; i++) {
        const todo: ITodoItem = result.data.value[i];
        todos.push({
          id: todo.Id,
          title: todo.Title,
          done: todo.Status === 'Completed'
        });
      }
      deferred.resolve(todos);
    });

    return deferred.promise;
  }

  public addTodo(todo: string): angular.IPromise<{}> {
    const deferred: angular.IDeferred<{}> = this.$q.defer();

    let listItemEntityTypeFullName: string = undefined;
    this.getListItemEntityTypeFullName()
      .then((entityTypeName: string): angular.IPromise<string> => {
        listItemEntityTypeFullName = entityTypeName;
        return this.getRequestDigest();
      })
      .then((requestDigest: string): void => {
        const body: string = JSON.stringify({
          '__metadata': { 'type': listItemEntityTypeFullName },
          'Title': todo
        });
        this.$http({
          url: `${this.sharepointApi}web/lists/getbytitle('${this.todoListName}')/items`,
          method: 'POST',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'X-RequestDigest': requestDigest
          },
          data: body
        }).then((result: angular.IHttpPromiseCallbackArg<{}>): void => {
          deferred.resolve();
        });
      });

    return deferred.promise;
  }

  public deleteTodo(todo: ITodo): angular.IPromise<{}> {
    const deferred: angular.IDeferred<{}> = this.$q.defer();

    this.getRequestDigest()
      .then((requestDigest: string): void => {
        this.$http({
          url: `${this.sharepointApi}web/lists/getbytitle('${this.todoListName}')/items(${todo.id})`,
          method: 'POST',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'X-RequestDigest': requestDigest,
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }).then((result: angular.IHttpPromiseCallbackArg<{}>): void => {
          deferred.resolve();
        });
      });

    return deferred.promise;
  }

  public setTodoStatus(todo: ITodo, done: boolean): angular.IPromise<{}> {
    const deferred: angular.IDeferred<{}> = this.$q.defer();

    let listItemEntityTypeFullName: string = undefined;
    this.getListItemEntityTypeFullName()
      .then((entityTypeName: string): angular.IPromise<string> => {
        listItemEntityTypeFullName = entityTypeName;
        return this.getRequestDigest();
      })
      .then((requestDigest: string): void => {
        const body: string = JSON.stringify({
          '__metadata': { 'type': listItemEntityTypeFullName },
          'Status': done ? 'Completed' : 'Not started'
        });
        this.$http({
          url: `${this.sharepointApi}web/lists/getbytitle('${this.todoListName}')/items(${todo.id})`,
          method: 'POST',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'X-RequestDigest': requestDigest,
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          data: body
        }).then((result: angular.IHttpPromiseCallbackArg<{}>): void => {
          deferred.resolve();
        });
      });

    return deferred.promise;
  }

  private getRequestDigest(): angular.IPromise<string> {
    const deferred: angular.IDeferred<string> = this.$q.defer();

    this.$http({
      url: this.sharepointApi + 'contextinfo',
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((result: angular.IHttpPromiseCallbackArg<{ FormDigestValue: string }>): void => {
      deferred.resolve(result.data.FormDigestValue);
    }, (err: any): void => {
      deferred.reject(err);
    });

    return deferred.promise;
  }

  private getListItemEntityTypeFullName(): angular.IPromise<string> {
    const deferred: angular.IDeferred<string> = this.$q.defer();

    this.$http({
      url: `${this.sharepointApi}web/lists/getbytitle('${this.todoListName}')?$select=ListItemEntityTypeFullName`,
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((result: angular.IHttpPromiseCallbackArg<{ ListItemEntityTypeFullName: string }>): void => {
      deferred.resolve(result.data.ListItemEntityTypeFullName);
    }, (err: any): void => {
      deferred.reject(err);
    });

    return deferred.promise;
  }
}
```

### <a name="upgrade-home-controller"></a>升级主控制器

在项目中，将 **./src/webparts/toDo/app/home.controller.js** 文件重命名为 `HomeController.ts`。将其内容更改为：

```ts
import * as angular from 'angular';
import { IDataService, ITodo } from './DataService';

export default class HomeController {
  public isLoading: boolean = false;
  public newItem: string = null;
  public todoCollection: ITodo[] = [];

  public static $inject: string[] = ['DataService', '$window'];

  constructor(private dataService: IDataService, private $window: angular.IWindowService) {
    this.loadTodos();
  }

  private loadTodos(): void {
    this.isLoading = true;
    this.dataService.getTodos()
      .then((todos: ITodo[]): void => {
        this.todoCollection = todos;
      })
      .finally((): void => {
        this.isLoading = false;
      });
  }

  public todoKeyDown($event: KeyboardEvent): void {
    if ($event.keyCode === 13 && this.newItem.length > 0) {
      $event.preventDefault();

      this.todoCollection.unshift({ id: -1, title: this.newItem, done: false });

      this.dataService.addTodo(this.newItem)
        .then((): void => {
          this.newItem = null;
          this.dataService.getTodos()
            .then((todos: ITodo[]): void => {
              this.todoCollection = todos;
            });
        });
    }
  }

  public deleteTodo(todo: ITodo): void {
    if (this.$window.confirm('Are you sure you want to delete this todo item?')) {
      let index: number = -1;
      for (let i: number = 0; i < this.todoCollection.length; i++) {
        if (this.todoCollection[i].id === todo.id) {
          index = i;
          break;
        }
      }

      if (index > -1) {
        this.todoCollection.splice(index, 1);
      }

      this.dataService.deleteTodo(todo)
        .then((): void => {
          this.dataService.getTodos()
            .then((todos: ITodo[]): void => {
              this.todoCollection = todos;
            });
        });
    }
  }

  public completeTodo(todo: ITodo): void {
    todo.done = true;

    this.dataService.setTodoStatus(todo, true)
      .then((): void => {
        this.dataService.getTodos()
          .then((todos: ITodo[]): void => {
            this.todoCollection = todos;
          });
      });
  }

  public undoTodo(todo: ITodo): void {
    todo.done = false;

    this.dataService.setTodoStatus(todo, false)
      .then((): void => {
        this.dataService.getTodos()
          .then((todos: ITodo[]): void => {
            this.todoCollection = todos;
          });
      });
  }
}
```

### <a name="upgrade-application-module"></a>升级应用程序模块

在项目中，将 **./src/webparts/toDo/app/app.module.js** 文件重命名为 `app.module.ts`。将其内容更改为：

```ts
import * as angular from 'angular';
import config from './app.config';
import HomeController from './HomeController';
import DataService from './DataService';

import 'ng-office-ui-fabric';

const todoapp: angular.IModule = angular.module('todoapp', [
  'officeuifabric.core',
  'officeuifabric.components'
]);

config();

todoapp
  .controller('HomeController', HomeController)
  .service('DataService', DataService);
```

### <a name="update-reference-to-angular-application-in-the-web-part"></a>更新 Web 部件中对 Angular 应用程序的引用

既然使用 TypeScript 生成了 Angular 应用程序，并且其各个部分相互引用，所以 Web 部件不再需要引用应用程序的所有部分。相反，只需要加载主模块，其结果是将会加载构成 Angular 应用程序的所有其他元素。

在代码编辑器中打开 **./src/webparts/toDo/ToDoWebPart.ts** 文件。将 **render** 方法更改为：

```ts
export default class ToDoWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {
  // ...
  public render(): void {
    if (this.renderedOnce === false) {
      require('./app/app.module');

      this.domElement.innerHTML = `
        <div class="${styles.toDo}">
          <div data-ng-controller="HomeController as vm">
            <div id="loading" ng-show="vm.isLoading">
              <uif-spinner>Loading...</uif-spinner>
            </div>
            <div id="entryform" ng-show="vm.isLoading === false">
              <uif-textfield uif-label="New to do:" uif-underlined ng-model="vm.newItem" ng-keydown="vm.todoKeyDown($event)"></uif-textfield>
            </div>
            <uif-list id="items" ng-show="vm.isLoading === false" >
              <uif-list-item ng-repeat="todo in vm.todoCollection" uif-item="todo" ng-class="{'${styles.done}': todo.done}">
                <uif-list-item-primary-text>{{todo.title}}</uif-list-item-primary-text>
                <uif-list-item-actions>
                  <uif-list-item-action ng-click="vm.completeTodo(todo)" ng-show="todo.done === false">
                    <i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i>
                  </uif-list-item-action>
                  <uif-list-item-action ng-click="vm.undoTodo(todo)" ng-show="todo.done">
                    <i class="ms-Icon ms-Icon--RevToggleKey" aria-hidden="true"></i>
                  </uif-list-item-action>
                  <uif-list-item-action ng-click="vm.deleteTodo(todo)">
                    <i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i>
                  </uif-list-item-action>
                </uif-list-item-actions>
              </uif-list-item>
            </uif-list>
          </div>
        </div>`;

      angular.bootstrap(this.domElement, ['todoapp']);
    }
  }
  // ...
}
```

若要验证是否已成功升级到 TypeScript，请在命令行中运行

```sh
gulp serve --nobrowser
```

在 Web 浏览器中，刷新 SharePoint 工作台，这样它会显示与之前一样的 Web 部件。

![迁移的 Angular 应用程序在 Web 部件中正确标记为已完成项](../../../../images/ng-migration-correct-styling.png)

即使 Web 部件的工作方式没有变化，代码也会得到改进。在将来的更新中，你可以更轻松地验证已在开发的代码的正确性和完整性。

## <a name="improve-integration-of-the-angular-application-with-the-sharepoint-framework"></a>改进 Angular 应用程序与 SharePoint Framework 的集成

此时，Angular 应用程序能够正常工作，并包装在 SharePoint Framework 客户端 Web 部件中。用户可以向页添加 Web 部件，但无法配置 Web 部件的工作方式。所有配置均嵌入到 Angular 应用程序代码中。在本节中，你将扩展 Web 部件，以允许配置存储 todo 项的列表的名称，以及 Web 部件是否显示已完成的任务。

### <a name="define-web-part-properties"></a>定义 Web 部件属性

在代码编辑器中打开 **./src/webparts/toDo/ToDoWebPart.manifest.json** 文件。将 **properties** 部分更改为：

```json
"properties": {
  "todoListName": "Todo",
  "hideFinishedTasks": false
}
```

![Web 部件清单使用 Angular 应用程序的新属性扩展](../../../../images/ng-migration-angular-web-part-properties.png)

将 **./src/webparts/toDo/IToDoWebPartProps.ts** 文件的内容更改为：

```ts
export interface IToDoWebPartProps {
  todoListName: string;
  hideFinishedTasks: boolean;
}
```

在 **./src/webparts/toDo/ToDoWebPart.ts** 文件中，将第一个导入语句更改为：

```ts
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-client-preview';
```

![更新导入语句以加载 PropertyPaneToggle 属性窗格控件](../../../../images/ng-migration-import-propertypanetoggle.png)

接下来，在同一文件中，将 **propertyPaneSettings** 方法更改为：

```ts
export default class ToDoWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {
  // ...
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
                PropertyPaneTextField('todoListName', {
                  label: strings.ListNameFieldLabel
                }),
                PropertyPaneToggle('hideFinishedTasks', {
                  label: strings.HideFinishedTasksFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
  // ...
}
```

通过将 **./src/webparts/toDo/loc/mystrings.d.ts** 文件内容更改为以下内容，添加缺少的资源字符串：

```ts
declare interface IToDoStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListNameFieldLabel: string;
  HideFinishedTasksFieldLabel: string;
}

declare module 'toDoStrings' {
  const strings: IToDoStrings;
  export = strings;
}
```

在 **./src/webparts/toDo/loc/en-us.js** 文件中，为新添加的字符串添加翻译：

```js
define([], function() {
  return {
    "PropertyPaneDescription": "Description",
    "BasicGroupName": "Group Name",
    "ListNameFieldLabel": "List name",
    "HideFinishedTasksFieldLabel": "Hide finished tasks"
  }
});
```

### <a name="pass-web-part-properties-values-to-the-angular-application"></a>将 Web 部件属性值传递给 Angular 应用程序

此时，用户可以配置 Web 部件的工作方式，但 Angular 应用程序不会使用这些值。在本节中，你将扩展 Angular 应用程序，以通过 Web 部件属性窗格使用由用户提供的配置值。达到此目的一个方法是在 **render** 方法中广播 Angula 事件，并在 Web 部件使用的控制器中订阅该事件。

#### <a name="delete-angular-configuration-file"></a>删除 Angular 配置文件

在项目中删除 **./src/webparts/toDo/app/app.config.js** 文件。以下步骤将更新应用程序，以从 Web 部件属性获取配置值。

#### <a name="remove-reference-to-configuration"></a>删除对配置的引用

在 **./src/webparts/toDo/app/app.module.ts** 文件中，通过将其内容更改为以下内容，可删除对 Angular 配置的引用：

```ts
import * as angular from 'angular';
import HomeController from './HomeController';
import DataService from './DataService';

import 'ng-office-ui-fabric';

const todoapp: angular.IModule = angular.module('todoapp', [
  'officeuifabric.core',
  'officeuifabric.components'
]);

todoapp
  .controller('HomeController', HomeController)
  .service('DataService', DataService);
```

#### <a name="update-data-service-to-accept-configuration-value-in-method-parameters"></a>更新数据服务以接受中方法参数中的配置值

数据服务最初从 **app.config.ts** 文件中定义的常数检索其配置。为了转为使用 Web 部件属性中配置的配置值，特定方法必须接受参数。 

在代码编辑器中打开 **./src/webparts/toDo/app/DataService.ts** 文件，并将其内容更改为：

```ts
import * as angular from 'angular';

export interface ITodo {
  id: number;
  title: string;
  done: boolean;
}

interface ITodoItem {
  Id: number;
  Title: string;
  Status: string;
}

export interface IDataService {
  getTodos: (sharePointApi: string, todoListName: string, hideFinishedTasks: boolean) => angular.IPromise<ITodo[]>;
  addTodo: (todo: string, sharePointApi: string, todoListName: string) => angular.IPromise<{}>;
  deleteTodo: (todo: ITodo, sharePointApi: string, todoListName: string) => angular.IPromise<{}>;
  setTodoStatus: (todo: ITodo, done: boolean, sharePointApi: string, todoListName: string) => angular.IPromise<{}>;
}

export default class DataService implements IDataService {
  public static $inject: string[] = ['$q', '$http'];

  constructor(private $q: angular.IQService, private $http: angular.IHttpService) {
  }

  public getTodos(sharePointApi: string, todoListName: string, hideFinishedTasks: boolean): angular.IPromise<ITodo[]> {
    const deferred: angular.IDeferred<ITodo[]> = this.$q.defer();

    let url: string = `${sharePointApi}web/lists/getbytitle('${todoListName}')/items?$select=Id,Title,Status&$orderby=ID desc`;

    if (hideFinishedTasks === true) {
      url += "&$filter=Status ne 'Completed'";
    }

    this.$http({
      url: url,
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((result: angular.IHttpPromiseCallbackArg<{ value: ITodoItem[] }>): void => {
      const todos: ITodo[] = [];
      for (let i: number = 0; i < result.data.value.length; i++) {
        const todo: ITodoItem = result.data.value[i];
        todos.push({
          id: todo.Id,
          title: todo.Title,
          done: todo.Status === 'Completed'
        });
      }
      deferred.resolve(todos);
    });

    return deferred.promise;
  }

  public addTodo(todo: string, sharePointApi: string, todoListName: string): angular.IPromise<{}> {
    const deferred: angular.IDeferred<{}> = this.$q.defer();

    let listItemEntityTypeFullName: string = undefined;
    this.getListItemEntityTypeFullName(sharePointApi, todoListName)
      .then((entityTypeName: string): angular.IPromise<string> => {
        listItemEntityTypeFullName = entityTypeName;
        return this.getRequestDigest(sharePointApi);
      })
      .then((requestDigest: string): void => {
        const body: string = JSON.stringify({
          '__metadata': { 'type': listItemEntityTypeFullName },
          'Title': todo
        });
        this.$http({
          url: `${sharePointApi}web/lists/getbytitle('${todoListName}')/items`,
          method: 'POST',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'X-RequestDigest': requestDigest
          },
          data: body
        }).then((result: angular.IHttpPromiseCallbackArg<{}>): void => {
          deferred.resolve();
        });
      });

    return deferred.promise;
  }

  public deleteTodo(todo: ITodo, sharePointApi: string, todoListName: string): angular.IPromise<{}> {
    const deferred: angular.IDeferred<{}> = this.$q.defer();

    this.getRequestDigest(sharePointApi)
      .then((requestDigest: string): void => {
        this.$http({
          url: `${sharePointApi}web/lists/getbytitle('${todoListName}')/items(${todo.id})`,
          method: 'POST',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'X-RequestDigest': requestDigest,
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }).then((result: angular.IHttpPromiseCallbackArg<{}>): void => {
          deferred.resolve();
        });
      });

    return deferred.promise;
  }

  public setTodoStatus(todo: ITodo, done: boolean, sharePointApi: string, todoListName: string): angular.IPromise<{}> {
    const deferred: angular.IDeferred<{}> = this.$q.defer();

    let listItemEntityTypeFullName: string = undefined;
    this.getListItemEntityTypeFullName(sharePointApi, todoListName)
      .then((entityTypeName: string): angular.IPromise<string> => {
        listItemEntityTypeFullName = entityTypeName;
        return this.getRequestDigest(sharePointApi);
      })
      .then((requestDigest: string): void => {
        const body: string = JSON.stringify({
          '__metadata': { 'type': listItemEntityTypeFullName },
          'Status': done ? 'Completed' : 'Not started'
        });
        this.$http({
          url: `${sharePointApi}web/lists/getbytitle('${todoListName}')/items(${todo.id})`,
          method: 'POST',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'X-RequestDigest': requestDigest,
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          data: body
        }).then((result: angular.IHttpPromiseCallbackArg<{}>): void => {
          deferred.resolve();
        });
      });

    return deferred.promise;
  }

  private getRequestDigest(sharePointApi: string): angular.IPromise<string> {
    const deferred: angular.IDeferred<string> = this.$q.defer();

    this.$http({
      url: sharePointApi + 'contextinfo',
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((result: angular.IHttpPromiseCallbackArg<{ FormDigestValue: string }>): void => {
      deferred.resolve(result.data.FormDigestValue);
    }, (err: any): void => {
      deferred.reject(err);
    });

    return deferred.promise;
  }

  private getListItemEntityTypeFullName(sharePointApi: string, todoListName: string): angular.IPromise<string> {
    const deferred: angular.IDeferred<string> = this.$q.defer();

    this.$http({
      url: `${sharePointApi}web/lists/getbytitle('${todoListName}')?$select=ListItemEntityTypeFullName`,
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((result: angular.IHttpPromiseCallbackArg<{ ListItemEntityTypeFullName: string }>): void => {
      deferred.resolve(result.data.ListItemEntityTypeFullName);
    }, (err: any): void => {
      deferred.reject(err);
    });

    return deferred.promise;
  }
}
```

#### <a name="broadcast-properties-change-event"></a>广播属性更改事件

在 **./src/webparts/toDo/ToDoWebPart.ts** 文件中，将名为 `$injector` 的新属性添加到 **ToDoWebPart** 类：

```ts
export default class ToDoWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {
  private $injector: angular.auto.IInjectorService;
  // ...
}
```

![“Injector”属性在 ToDoWebPart 类中突出显示](../../../../images/ng-migration-injector-property.png)

在同一文件中，将 **render** 方法更新为：

```ts
export default class ToDoWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {
  // ...
  public render(): void {
    if (this.renderedOnce === false) {
      require('./app/app.module');

      this.domElement.innerHTML = `
        <div class="${styles.toDo}">
          <div data-ng-controller="HomeController as vm">
            <div class="${styles.configurationNeeded}" ng-show="vm.configurationNeeded">
              Please configure the web part
            </div>
            <div ng-show="vm.configurationNeeded === false">
              <div id="loading" ng-show="vm.isLoading">
                <uif-spinner>Loading...</uif-spinner>
              </div>
              <div id="entryform" ng-show="vm.isLoading === false">
                <uif-textfield uif-label="New to do:" uif-underlined ng-model="vm.newItem" ng-keydown="vm.todoKeyDown($event)"></uif-textfield>
              </div>
              <uif-list id="items" ng-show="vm.isLoading === false" >
                <uif-list-item ng-repeat="todo in vm.todoCollection" uif-item="todo" ng-class="{'${styles.done}': todo.done}">
                  <uif-list-item-primary-text>{{todo.title}}</uif-list-item-primary-text>
                  <uif-list-item-actions>
                    <uif-list-item-action ng-click="vm.completeTodo(todo)" ng-show="todo.done === false">
                      <i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i>
                    </uif-list-item-action>
                    <uif-list-item-action ng-click="vm.undoTodo(todo)" ng-show="todo.done">
                      <i class="ms-Icon ms-Icon--RevToggleKey" aria-hidden="true"></i>
                    </uif-list-item-action>
                    <uif-list-item-action ng-click="vm.deleteTodo(todo)">
                      <i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i>
                    </uif-list-item-action>
                  </uif-list-item-actions>
                </uif-list-item>
              </uif-list>
            </div>
          </div>
        </div>`;

      this.$injector = angular.bootstrap(this.domElement, ['todoapp']);
    }

    this.$injector.get('$rootScope').$broadcast('configurationChanged', {
      sharePointApi: this.context.pageContext.web.absoluteUrl + '/_api/',
      todoListName: this.properties.todoListName,
      hideFinishedTasks: this.properties.hideFinishedTasks
    });
  }
  // ...
}
```

在 **./src/webparts/toDo/ToDo.module.scss** 文件中，为 **.configurationNeeded** 类添加缺少的样式：

```css
.toDo {
  /* ... */
  .configurationNeeded {
    margin: 0 auto;
    width: 100%;
    text-align: center;
  }
  /* ... */
}
```

#### <a name="subscribe-to-the-properties-changed-event"></a>订阅属性更改事件

在代码编辑器中打开 **./src/webparts/toDo/app/HomeController.ts** 文件。

在 **HomeController** 类中添加以下属性：

```ts
export default class HomeController {
  // ...
  private sharePointApi: string = undefined;
  private todoListName: string = undefined;
  private hideFinishedTasks: boolean = false;
  private configurationNeeded: boolean = true;
  // ...
}
```

使用注入根范围服务扩展 **HomeController** 类的构造函数，并将其内容更改为：

```ts
export default class HomeController {
  // ...
  public static $inject: string[] = ['DataService', '$window', '$rootScope'];

  constructor(private dataService: IDataService,
    private $window: angular.IWindowService,
    $rootScope: angular.IRootScopeService) {
    const vm: HomeController = this;
    this.init(undefined, undefined);

    $rootScope.$on('configurationChanged',
      (event: angular.IAngularEvent,
       args: {
         sharePointApi: string;
         todoListName: string;
         hideFinishedTasks: boolean;
        }): void => {
      vm.init(args.sharePointApi, args.todoListName, args.hideFinishedTasks);
    });
  }

  // ...
}
```

将 **init** 方法添加到 **HomeController** 类：

```ts
export default class HomeController {
  // ...
  private init(sharePointApi: string, todoListName: string, hideFinishedTasks?: boolean): void {
    if (sharePointApi !== undefined && sharePointApi.length > 0 &&
      todoListName !== undefined && todoListName.length > 0) {
      this.sharePointApi = sharePointApi;
      this.todoListName = todoListName;
      this.hideFinishedTasks = hideFinishedTasks;
      this.loadTodos();
      this.configurationNeeded = false;
    }
    else {
      this.configurationNeeded = true;
    }
  }
  // ...
}
```

更新 **HomeController** 类中所有剩余的方法，以使用类属性中的配置值：

```ts
export default class HomeController {
  // ...
  private loadTodos(): void {
    this.isLoading = true;
    this.dataService.getTodos(this.sharePointApi, this.todoListName, this.hideFinishedTasks)
      .then((todos: ITodo[]): void => {
        this.todoCollection = todos;
      })
      .finally((): void => {
        this.isLoading = false;
      });
  }

  public todoKeyDown($event: KeyboardEvent): void {
    if ($event.keyCode === 13 && this.newItem.length > 0) {
      $event.preventDefault();

      this.todoCollection.unshift({ id: -1, title: this.newItem, done: false });

      this.dataService.addTodo(this.newItem, this.sharePointApi, this.todoListName)
        .then((): void => {
          this.newItem = null;
          this.dataService.getTodos(this.sharePointApi, this.todoListName, this.hideFinishedTasks)
            .then((todos: ITodo[]): void => {
              this.todoCollection = todos;
            });
        });
    }
  }

  public deleteTodo(todo: ITodo): void {
    if (this.$window.confirm('Are you sure you want to delete this todo item?')) {
      let index: number = -1;
      for (let i: number = 0; i < this.todoCollection.length; i++) {
        if (this.todoCollection[i].id === todo.id) {
          index = i;
          break;
        }
      }

      if (index > -1) {
        this.todoCollection.splice(index, 1);
      }

      this.dataService.deleteTodo(todo, this.sharePointApi, this.todoListName)
        .then((): void => {
          this.dataService.getTodos(this.sharePointApi, this.todoListName, this.hideFinishedTasks)
            .then((todos: ITodo[]): void => {
              this.todoCollection = todos;
            });
        });
    }
  }

  public completeTodo(todo: ITodo): void {
    todo.done = true;

    this.dataService.setTodoStatus(todo, true, this.sharePointApi, this.todoListName)
      .then((): void => {
        this.dataService.getTodos(this.sharePointApi, this.todoListName, this.hideFinishedTasks)
          .then((todos: ITodo[]): void => {
            this.todoCollection = todos;
          });
      });
  }

  public undoTodo(todo: ITodo): void {
    todo.done = false;

    this.dataService.setTodoStatus(todo, false, this.sharePointApi, this.todoListName)
      .then((): void => {
        this.dataService.getTodos(this.sharePointApi, this.todoListName, this.hideFinishedTasks)
          .then((todos: ITodo[]): void => {
            this.todoCollection = todos;
          });
      });
  }
}
```

通过在以下命令行中进行执行，可验证 Web 部件是否可正常工作：

```
gulp serve --nobrowser
```

在 Web 浏览器中，导航到 SharePoint 工作台并向画布添加 Web 部件。如果切换“**隐藏完成的任务**”选项，则应该可以看到相应显示或隐藏已完成的任务。

![Angular 应用程序根据 Web 部件属性配置隐藏已完成任务](../../../../images/ng-migration-finished-tasks-hidden.png)