# <a name="build-sharepoint-framework-client-side-web-parts-with-angular-v1x"></a>使用 Angular v1.x 生成 SharePoint Framework 客户端 Web 部件。

> **注意：**SharePoint Framework 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

如果你使用的是 Angular，则还可以使用这个极受欢迎的框架来生成客户端 Web 部件。由于有了 Angular 的现代化功能，它可用于从复杂的多视图单页面应用程序到更小的组件（如 Web 部件）的任意范围。许多组织过去已使用 Angular 生成 SharePoint 解决方案。本文介绍如何使用 Angular v1.x 生成 SharePoint Framework 客户端 Web 部件，并使用 [ngOfficeUIFabric](http://ngofficeuifabric.com) - Office UI Fabric 的 Angular 指令将其样式化。在本教程中，你将生成用于管理待执行项的简单 Web 部件。

![使用 Angular 和 SharePoint 工作台中显示的 ngOfficeUIFabric 生成的 SharePoint Framework 客户端 Web 部件](../../../../images/ng-intro-hide-finished-tasks.png)

工作 Web 部件源可在 GitHub 上获取，网址为：[https://github.com/SharePoint/sp-dev-fx-webparts/tree/master/samples/angular-ngofficeuifabric-todo](https://github.com/SharePoint/sp-dev-fx-webparts/tree/master/samples/angular-ngofficeuifabric-todo)

> **注意：**执行本文中的步骤之前，请确保[设置开发环境](https://github.com/SharePoint/sp-dev-docs/blob/staging/docs/spfx/set-up-your-development-environment)，以便生成 SharePoint Framework 解决方案。

## <a name="create-new-project"></a>新建项目

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

![带有默认选项的 SharePoint Framework Yeoman 生成器](../../../../images/ng-intro-yeoman-generator.png)

基架完成后，请在代码编辑器中打开项目文件夹。在本教程中，你将使用 Visual Studio Code。

![在 Visual Studio Code 中打开 SharePoint Framework 项目](../../../../images/ng-intro-project-visual-studio-code.png)

## <a name="add-angular-and-ngofficeuifabric"></a>添加 Angular 和 ngOfficeUIFabric

在本教程中，你将从 CDN 加载 Angular 和 ngOfficeUIFabric。为此，请在代码编辑器中打开 **config/config.json** 文件并在 **externals** 属性中添加以下行：

```json
"angular": {
  "path": "https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.5.8/angular.min.js",
  "globalName": "angular"
},
"ng-office-ui-fabric": "https://cdnjs.cloudflare.com/ajax/libs/ngOfficeUiFabric/0.12.3/ngOfficeUiFabric.js"
```

![Angular 和 ngOfficeUIFabric 已添加到 config.json 文件](../../../../images/ng-intro-angular-ngofficeuifabric-config.png)

## <a name="add-angular-typings-for-typescript"></a>添加 TypeScript Angular 键入

由于要在 Web 部件的代码中引用 Angular，因此还需要 TypeScript Angular 键入。若要安装它们，请在命令行中运行：

```sh
tsd install angular --save
```

## <a name="implement-angular-application"></a>实现 Angular 应用程序

所有先决条件就绪后，就可以开始实现 Angular 示例应用程序。因为它包含几个文件，所以要为其创建名为 **app** 的单独文件夹。

![“app”文件夹在 Visual Studio Code 中突出显示](../../../../images/ng-intro-app-folder.png)

### <a name="implement-to-do-data-service"></a>实现待执行数据服务

在新创建的 **app** 文件夹中，创建名为 **DataService.ts** 的新文件。在文件中粘贴以下代码：

```ts
export interface ITodo {
  id: number;
  title: string;
  done: boolean;
}

export interface IDataService {
  getTodos(hideFinishedTasks: boolean): ng.IPromise<ITodo[]>;
  addTodo(todo: string): ng.IPromise<{}>;
  deleteTodo(todo: ITodo): ng.IPromise<{}>;
  setTodoStatus(todo: ITodo, done: boolean): ng.IPromise<{}>;
}

export default class DataService implements IDataService {
  public static $inject: string[] = ['$q'];

  private items: ITodo[] = [
    {
      id: 1,
      title: 'Prepare demo web part',
      done: true
    },
    {
      id: 2,
      title: 'Show demo',
      done: false
    },
    {
      id: 3,
      title: 'Share code',
      done: false
    }
  ];
  private nextId: number = 4;

  constructor(private $q: ng.IQService) {
  }

  public getTodos(hideFinishedTasks: boolean): ng.IPromise<ITodo[]> {
    const deferred: ng.IDeferred<ITodo[]> = this.$q.defer();

    const todos: ITodo[] = [];
    for (let i: number = 0; i < this.items.length; i++) {
      if (hideFinishedTasks && this.items[i].done) {
        continue;
      }

      todos.push(this.items[i]);
    }

    deferred.resolve(todos);

    return deferred.promise;
  }

  public addTodo(todo: string): ng.IPromise<{}> {
    const deferred: ng.IDeferred<{}> = this.$q.defer();

    this.items.push({
      id: this.nextId++,
      title: todo,
      done: false
    });

    deferred.resolve();

    return deferred.promise;
  }

  public deleteTodo(todo: ITodo): ng.IPromise<{}> {
    const deferred: ng.IDeferred<{}> = this.$q.defer();

    let pos: number = -1;
    for (let i: number = 0; i < this.items.length; i++) {
      if (this.items[i].id === todo.id) {
        pos = i;
        break;
      }
    }

    if (pos > -1) {
      this.items.splice(pos, 1);
      deferred.resolve();
    }
    else {
      deferred.reject();
    }

    return deferred.promise;
  }

  public setTodoStatus(todo: ITodo, done: boolean): ng.IPromise<{}> {
    const deferred: ng.IDeferred<{}> = this.$q.defer();

    for (let i: number = 0; i < this.items.length; i++) {
      if (this.items[i].id === todo.id) {
        this.items[i].done = done;
      }
    }

    deferred.resolve();

    return deferred.promise;
  }
}
```

![DataService.ts 文件在 Visual Studio Code 中打开](../../../../images/ng-intro-dataservice.png)

在之前的代码片断中实现了三种类型：**ITodo** 接口表示应用程序中的待执行项，**IDataService** 接口定义数据服务签名，**DataService** 类负责检索和处理要执行的项。数据服务可实现用于添加和修改待执行项的简单方法。虽然这些操作是瞬时完成的，但为确保一致性，每个 CRUD 函数均返回一个承诺。 

在本教程中，待执行项存储在内存中。但是，你可以轻松扩展此解决方案，以便在 SharePoint 列表中存储项目，并使用数据服务通过其 REST API 与 SharePoint 通信。

### <a name="implement-the-controller"></a>实现控制器

接下来，实现可促进视图和数据服务之间通信的控制器。在 **app** 文件夹中，创建一个名为 **HomeController.ts** 的新文件，并粘贴以下代码：

```ts
import { IDataService, ITodo } from './DataService';

export default class HomeController {
  public isLoading: boolean = false;
  public newItem: string = null;
  public todoCollection: any[] = [];
  private hideFinishedTasks: boolean = false;

  public static $inject: string[] = ['DataService', '$window', '$rootScope'];

  constructor(private dataService: IDataService, private $window: ng.IWindowService, private $rootScope: ng.IRootScopeService) {
    const vm: HomeController = this;
    this.init();
  }

  private init(hideFinishedTasks?: boolean): void {
    this.hideFinishedTasks = hideFinishedTasks;
    this.loadTodos();
  }

  private loadTodos(): void {
    const vm: HomeController = this;
    this.isLoading = true;
    this.dataService.getTodos(vm.hideFinishedTasks)
      .then((todos: ITodo[]): void => {
        vm.todoCollection = todos;
      })
      .finally((): void => {
        vm.isLoading = false;
      });
  }

  public todoKeyDown($event: any): void {
    if ($event.keyCode === 13 && this.newItem.length > 0) {
      $event.preventDefault();

      this.todoCollection.unshift({ id: -1, title: this.newItem, done: false });
      const vm: HomeController = this;

      this.dataService.addTodo(this.newItem)
        .then((): void => {
          this.newItem = null;
          this.dataService.getTodos(vm.hideFinishedTasks)
            .then((todos: any[]): void => {
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

      const vm: HomeController = this;

      this.dataService.deleteTodo(todo)
        .then((): void => {
          this.dataService.getTodos(vm.hideFinishedTasks)
            .then((todos: any[]): void => {
              this.todoCollection = todos;
            });
        });
    }
  }

  public completeTodo(todo: ITodo): void {
    todo.done = true;

    const vm: HomeController = this;

    this.dataService.setTodoStatus(todo, true)
      .then((): void => {
        this.dataService.getTodos(vm.hideFinishedTasks)
          .then((todos: any[]): void => {
            this.todoCollection = todos;
          });
      });
  }

  public undoTodo(todo: ITodo): void {
    todo.done = false;

    const vm: HomeController = this;

    this.dataService.setTodoStatus(todo, false)
      .then((): void => {
        this.dataService.getTodos(vm.hideFinishedTasks)
          .then((todos: any[]): void => {
            this.todoCollection = todos;
          });
      });
  }
}
```

![HomeController.ts 文件在 Visual Studio Code 中打开](../../../../images/ng-intro-homecontroller.png)

首先加载之前实现的数据服务。控制器需要该数据服务，以获取项列表，并按用户要求修改项。使用 Angular 依赖关系注入，将服务注入到控制器中。该控制器实现了大量对视图模型公开的函数，且这些函数将从模板进行调用。使用这些函数，用户将能够添加新项、把项标记为已完成、待执行，或删除项。

### <a name="implement-the-main-module"></a>实现主模块

数据服务和控制器准备就绪后，定义应用程序的主模块，并在其上注册数据服务和控制器。在 **app** 文件夹中，创建名为 **app-module.ts** 的新文件，并粘贴以下代码：

```ts
import * as angular from 'angular';
import 'ng-office-ui-fabric';
import HomeController from './HomeController';
import DataService from './DataService';

const todoapp: ng.IModule = angular.module('todoapp', [
  'officeuifabric.core',
  'officeuifabric.components'
]);

todoapp
  .controller('HomeController', HomeController)
  .service('DataService', DataService);
```

![app-module.ts文件在 Visual Studio Code 中打开](../../../../images/ng-intro-app-module.png)

首先引用 Angular 和 ngOfficeUIFabric，并加载之前实现的控制器和数据服务。接下来，定义应用程序模块，并加载 **officeuifabric.core** 和 **officeuifabric.components** 模块作为应用程序的依赖关系。这样就可以在模板中使用 ngOfficeUIFabric 指令。最后，在应用程序中注册控制器和数据服务。

现在已经生成 Web 部件的 Angular 应用程序。在以下步骤中，你将使用 Web 部件注册 Angular 应用程序，并使用 Web 部件属性使其 变得可配置。

## <a name="register-angular-application-with-web-part"></a>使用 Web 部件注册 Angular 应用程序

下一步是将 Angular 应用程序添加到 Web 部件。在代码编辑器中打开 **ToDoWebPart.ts** 文件。

在类声明前添加以下行：

```ts
import * as angular from 'angular';
import './app/app-module';
```

![导入 Visual Studio Code 中突出显示的 ToDoWebPart.ts 文件中的语句](../../../../images/ng-intro-web-part-import-angular.png)

此操作允许我们将参考加载到 Angula 和应用程序，需要这两项来引导 Angular 应用程序。

将 Web 部件的 **render** 函数更改为：

```ts
public render(): void {
  if (this.renderedOnce === false) {
    this.domElement.innerHTML = `
<div class="${styles.toDo}">
  <div data-ng-controller="HomeController as vm">
    <div class="${styles.loading}" ng-show="vm.isLoading">
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
```

![Visual Studio Code 中的 Web 部件的 render 函数](../../../../images/ng-intro-web-part-render-angular.png)

此代码首先将应用程序模板直接分配给 Web 部件的 DOM 元素。在根元素上，可指定将在模板中处理事件和数据绑定的控制器的名称。然后，使用之前声明主模块时使用的 **todoapp** 名称引导应用程序。使用 **renderedOnce** Web 部件属性可以确保仅引导  Angular 应用程序一次。否则，如果更改了 Web 部件的其中一项属性，则会重新调用 **render** 函数，再次引导 Angular 应用程序，这将导致错误。

你还需要实现所使用模板的 CSS 样式。在代码编辑器中打开 **ToDo.module.scss** 文件，并将其内容替换为：

```css
.toDo {
  .loading {
    margin: 0 auto;
    width: 6em;
  }

  .done [class=ms-ListItem-primaryText] {
    text-decoration: line-through;
  }
}
```

![ToDo.module.scss 文件将在 Visual Studio Code 中打开](../../../../images/ng-intro-web-part-css.png)

## <a name="load-newer-version-of-office-ui-fabric"></a>加载 Office UI Fabric 较新的版本

运行以下命令，确认一切都按预期的方式运行：

```sh
gulp serve
```

在浏览器中，你会看到待执行 Web 部件显示要执行的项。

![显示待执行项的待执行 Web 部件使用 Office UI Fabric 的较早版本呈现](../../../../images/ng-intro-workbench-old-office-ui-fabric.png)

仔细查看，你会看到每一项前面的项目符号。这是由于 ngOfficeUIFabric 使用的 Office UI Fabric 比 SharePoint 工作台中使用的版本要高。在 Web 部件中加载 Office UI Fabric 的较新版本即可修复此问题。

如果尚未停止 **serve** 任务，则打开命令行并按 `CTRL+C` 停止 **gulp serve** 任务。

在代码编辑器中打开 **ToDoWebPart.ts** 文件。在 `import styles` 语句前添加以下行：

```ts
import ModuleLoader from '@microsoft/sp-module-loader';
```

![“import ModuleLoader”在 Visual Studio Code 中突出显示](../../../../images/ng-intro-module-loader.png)

接下来，将 Web 部件的构造函数更改为：

```ts
public constructor(context: IWebPartContext) {
  super(context);

  ModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.min.css');
  ModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.components.min.css');
}
```

![Visual Studio Code 中的 Web 部件的构造函数](../../../../images/ng-intro-constructor-office-ui-fabric.png)

在预览 Web 部件前，必须通过 **ModuleLoader** 类安装所需的附加 TypeScript 键入。在命令行中运行：

```sh
tsd install combokeys --save
```

现在已准备好在以下命令行中预览所做的更改：

```sh
gulp serve
```

现在你应该看到正确呈现的 Web 部件。

![显示待执行项的待执行 Web 部件正确呈现](../../../../images/ng-intro-workbench-new-office-ui-fabric.png)

## <a name="make-web-part-configurable"></a>使 Web 部件可配置

此时，待执行 Web 部件显示一个待执行项的固定列表。接下来，可以使用配置选项将其扩展，以允许用户选择是否要查看标记为完成的项目。

### <a name="add-property-in-the-web-part-manifest"></a>在 Web 部件清单中添加属性

首先在 Web 部件清单中添加配置属性。在代码编辑器中打开 **ToDoWebPart.manifest.json** 文件。在 **preconfiguredEntries** 部分中，导航到 **properties** 数组，并将现有的 **description** 属性替换为以下行：

```json
"hideFinishedTasks": false
```

![hideFinishedTasks 属性在 Web 部件清单中突出显示](../../../../images/ng-intro-manifest-property.png)

### <a name="update-the-signature-of-the-web-part-properties-interface"></a>更新 Web 部件属性接口签名

接下来，更新 Web 部件属性接口签名

在代码编辑器中打开 **IToDoWebPartProps.ts** 文件，并将其内容替换为以下内容︰

```ts
export interface IToDoWebPartProps {
  hideFinishedTasks: boolean;
}
```

![IToDoWebPartProps.ts 文件将在 Visual Studio Code 中打开](../../../../images/ng-intro-property-interface.png)

### <a name="add-the-property-to-the-web-part-property-pane"></a>将属性添加到 Web 部件属性窗格

若要允许用户配置新添加的属性的值，则需要通过 Web 部件属性窗格将其公开。

在代码编辑器中打开 **ToDoWebPart.ts** 文件。在第一个 `import` 语句中，将 `PropertyPaneTextField` 替换为 `PropertyPaneToggle`：

```ts
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneToggle
} from '@microsoft/sp-client-preview';
```

![“PropertyPaneToggle”导入语句在 Visual Studio Code 中突出显示](../../../../images/ng-intro-property-pane-toggle.png)

将 `propertyPaneSettings` 函数的实现更改为：

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
            groupName: strings.ViewGroupName,
            groupFields: [
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
```

若要修复缺少的字符串引用，首先需要更改字符串的签名。在代码编辑器中打开 **loc/mystrings.d.ts** 文件，并将其内容更改为：

```ts
declare interface IToDoStrings {
  PropertyPaneDescription: string;
  ViewGroupName: string;
  HideFinishedTasksFieldLabel: string;
}

declare module 'toDoStrings' {
  const strings: IToDoStrings;
  export = strings;
}
```

![loc/mystrings.d.ts 文件将在 Visual Studio Code 中打开](../../../../images/ng-intro-strings-interface.png)

接下来，需要为新定义的字符串提供实际值。在代码编辑器中打开 **loc/en-us.js** 文件，并将其内容更改为：

```js
define([], function() {
  return {
    "PropertyPaneDescription": "Manage configuration of the To do web part",
    "ViewGroupName": "View",
    "HideFinishedTasksFieldLabel": "Hide finished tasks"
  }
});
```

![loc/en-us.js 文件将在 Visual Studio Code 中打开](../../../../images/ng-intro-strings.png)

运行以下命令，确认一切都按预期的方式运行：

```sh
gulp serve
```

在 Web 部件属性窗格中，你应该可以看到新定义属性的切换按钮。

![切换按钮显示在 Web 部件属性窗格上](../../../../images/ng-intro-property-pane-toggle-browser.png)

此时单击切换按钮对 Web 部件没有任何影响，因为尚未没有连接到 Angular。你将在下一步执行这一操作。

### <a name="make-the-angular-application-configurable-using-web-part-properties"></a>使用 Web 部件属性使 Angular 应用程序可配置

在上一步中定义的 Web 部件属性允许用户选择是否显示完成的任务。接下来将把用户所选的值传递到 Angular 应用程序，这样它就可以加载相应项。

#### <a name="broadcast-angular-event-on-web-part-property-change"></a>针对 Web 部件属性更改广播 Angular 事件

在代码编辑器中打开 **ToDoWebPart.ts** 文件。在 Web 部件构造函数前添加以下行：

```ts
private $injector: ng.auto.IInjectorService;
```

![$injector 类变量在 Visual Studio Code 中突出显示](../../../../images/ng-intro-injector-class-variable.png)

接下来，将 Web 部件的 **render** 函数更改为以下内容：

```ts
public render(): void {
  if (this.renderedOnce === false) {
    this.domElement.innerHTML = `
<div class="${styles.toDo}">
  <div data-ng-controller="HomeController as vm">
    <div class="${styles.loading}" ng-show="vm.isLoading">
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

    this.$injector = angular.bootstrap(this.domElement, ['todoapp']);
  }

  this.$injector.get('$rootScope').$broadcast('configurationChanged', {
    hideFinishedTasks: this.properties.hideFinishedTasks
  });
}
```

在上面的代码示例中，每次更改 Web 部件属性时，相应属性会向订阅 Angular 事件的 Angular 应用程序广播该事件。Angular 应用程序收到应用程序事件后，会进行相应处理。

#### <a name="subscribe-to-web-part-property-change-event-in-angular"></a>订阅 Angular 中的 Web 部件属性更改事件

在代码编辑器中打开 **app/HomeController.ts** 文件。将构造函数扩展如下：

```ts
constructor(private dataService: IDataService, private $window: ng.IWindowService, private $rootScope: ng.IRootScopeService) {
  const vm: HomeController = this;
  this.init();

  $rootScope.$on('configurationChanged', (event: ng.IAngularEvent, args: { hideFinishedTasks: boolean }): void => {
    vm.init(args.hideFinishedTasks);
  });
}
```

![HomeController.ts 文件中的构造函数定义在 Visual Studio Code 中突出显示](../../../../images/ng-intro-homecontroller-event.png)

要验证的 Angular 应用程序按预期方式工作，并对更改的属性作出正确响应，请在命令行中运行：

```sh
gulp serve
```

如果切换**隐藏已完成的任务**属性的值，Web 部件应相应地显示或隐藏已完成的任务。

![启用“隐藏完成的任务”选项后，Web 部件仅显示未完成的任务](../../../../images/ng-intro-hide-finished-tasks.png)