# <a name="use-existing-javascript-libraries-in-sharepoint-framework-client-side-web-parts"></a>使用 SharePoint Framework 客户端 Web 部件中的现有 JavaScript 库

当在 SharePoint Framework 上生成客户端 Web 部件时，可以从使用现有的 JavaScript 库中受益，该库可用于生成功能强大的解决方案。但是，有一些应该加以考虑的注意事项以确保 Web 部件不会负面影响正在使用它们的 SharePoint 页面的性能。

## <a name="reference-existing-libraries-as-packages"></a>以打包方式引用现有库

引用 SharePoint Framework 客户端 Web 部件中的现有 JavaScript 库最常见的方法是将其作为项目中的包进行安装。以 Angular 为例，为了在客户端 Web 部件中使用该 Angular，将首先使用 **npm** 进行安装：

```sh
npm install angular --save
```

接下来，为了将 Angular 与 TypeScript 一起使用，将利用 **tsd** 安装键入：

```sh
tsd install angular --save
```

最后，会在 Web 部件 中通过 `import` 语句引用 Angular：

```ts
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import styles from './HelloWorld.module.scss';
import * as strings from 'helloWorldStrings';
import { IHelloWorldWebPartProps } from './IHelloWorldWebPartProps';

import * as angular from 'angular';

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <!-- omitted for brevity -->
      </div>`;

      angular.module('helloworld', []);

      angular.bootstrap(this.domElement, ['helloworld']);
  }

  // omitted for brevity
}
```

## <a name="bundle-web-part-resources"></a>捆绑 Web 部件资源

SharePoint Framework 使用基于开源工具的生成工具链，该开源工具包括 gulp 和 Webpack。当生成 SharePoint Framework 项目时，这些生成工具在被称为绑定的过程中将所有引用资源自动组合成单个 JavaScript 文件。

![具有在包含输出文件的 SharePoint Framework dist 文件夹旁边可见的 gulp 捆绑任务的命令窗口](../../../../images/external-scripts-gulp-bundle.png)

绑定可提供很多优势。首先，Web 部件所需的全部资源都能在单个 JavaScript 文件中使用。这简化了 Web 部件包含单个文件的部署，并且不可能会漏掉在部署过程中的依赖项。当 Web 部件使用不同的资源时，必须按正确顺序进行加载。Webpack 在生成过程中生成的 Web 部件包为你管理不同资源的加载，包括解决这些资源之间的任何依赖项。绑定 Web 部件还对最终用户有好处：通常情况下与大量的小文件相比，该绑定会更快地下载单个较大文件。通过将大量较小文件组合成一个较大的捆绑，Web 部件将在页面上更快地加载。但是，将现有 JavaScript 库与 SharePoint Framework 客户端 Web 部件绑定并不是没有缺点。

当绑定 SharePoint Framework 中现有的 JavaScript 框架时，所有引用脚本都包括在已生成的捆绑文件中。按照 Angular 示例，包括 Angular 在内的已优化 Web 部件捆绑超过 170 KB。

![在资源管理器中突出显示的捆绑文件大小](../../../../images/external-scripts-bundle-size.png)

如果将另一个 Web 部件添加到也使用 Angular 的项目并生成该项目，你将会获得两个捆绑文件（均超过 170 KB），每个文件用于一个 Web 部件。

![在资源管理器中突出显示的两个捆绑文件](../../../../images/external-scripts-two-bundles-size.png)

如果在页面添加这些 Web 部件，每个用户将会多次下载 Angular - 一次与页面上的每个 Web 部件一起下载。此方法效率低下，并会减缓加载页面的速度。

## <a name="reference-existing-libraries-as-external-resources"></a>引用现有库作为外部资源

针对 SharePoint Framework 客户端 Web 部件中的现有库，更好的利用方法是将其作为外部资源进行引用。这样一来，有关 Web 部件中包含的特定脚本的唯一信息是该脚本的 URL。添加到页面时，Web 部件将自动尝试从指定的 URL 加载所需的全部资源。

引用 SharePoint Framework 中的现有 JavaScript 库非常简单，不需要对代码进行任何特定的更改。因为此库在运行时已从指定的 URL 上加载，不需要再作为项目中的包进行安装。

以 Angular 为例，为了在客户端 Web 部件中将其作为外部资源进行引用，首先使用 **tsd** 安装 TypeScript 键入：

```sh
tsd install angular --save
```

接下来在 **config/config.json** 文件中，向 **externals** 属性添加以下项：

```json
"angular": {
  "path": "https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.5.8/angular.min.js",
  "globalName": "angular"
}
```

完整的 **config/config.json** 文件看起来类似于：

```json
{
  "entries": [
    {
      "entry": "./lib/webparts/helloWorld/HelloWorldWebPart.js",
      "manifest": "./src/webparts/helloWorld/HelloWorldWebPart.manifest.json",
      "outputPath": "./dist/hello-world.bundle.js"
    }
  ],
  "externals": {
    "@microsoft/sp-client-base": "node_modules/@microsoft/sp-client-base/dist/sp-client-base.js",
    "@microsoft/sp-client-preview": "node_modules/@microsoft/sp-client-preview/dist/sp-client-preview.js",
    "@microsoft/sp-lodash-subset": "node_modules/@microsoft/sp-lodash-subset/dist/sp-lodash-subset.js",
    "office-ui-fabric-react": "node_modules/office-ui-fabric-react/dist/office-ui-fabric-react.js",
    "react": "node_modules/react/dist/react.min.js",
    "react-dom": "node_modules/react-dom/dist/react-dom.min.js",
    "react-dom/server": "node_modules/react-dom/dist/react-dom-server.min.js",
    "angular": {
      "path": "https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.5.8/angular.min.js",
      "globalName": "angular"
    }
  },
  "localizedResources": {
    "helloWorldStrings": "webparts/helloWorld/loc/{locale}.js"
  }
}
```

最后，像之前一样在 Web 部件中引用 Angular：

```ts
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import styles from './HelloWorld.module.scss';
import * as strings from 'helloWorldStrings';
import { IHelloWorldWebPartProps } from './IHelloWorldWebPartProps';

import * as angular from 'angular';

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <!-- omitted for brevity -->
      </div>`;

      angular.module('helloworld', []);

      angular.bootstrap(this.domElement, ['helloworld']);
  }

  // omitted for brevity
}
```

如果现在生成项目并看一眼生成的捆绑文件的大小，你会注意到它只有 6 KB。

![在资源管理器中突出显示的捆绑文件大小](../../../../images/external-scripts-external-angular-bundle-size.png)

如果将另一个 Web 部件添加到也使用 Angular 的项目，并重新生成该项目，这两个捆绑文件将均为 6 KB。

![在资源管理器中突出显示的两个捆绑文件](../../../../images/external-scripts-external-angular-two-bundles-size.png)

不能假定刚刚保存超过 300 KB，这样是不对的。这两个 Web 部件仍然需要 Angular，并将在用户第一次访问放置 Web 部件之一的页面时对其进行加载。

![针对具有一个 Web 部件的页面，开发者工具中突出显示 Angular](../../../../images/external-scripts-browser-loading-angular.png)

即使将两个 Angular Web 部件添加到页面上，SharePoint Framework 仍会只下载 Angular 一次。

![针对具有两个 Web 部件的页面，开发者工具中突出显示 Angular](../../../../images/external-scripts-browser-loading-angular-two-web-parts.png)

引用现有 JavaScript 库作为外部资源的真正好处是对于所有常用的脚本，组织有一个集中的位置，或者你可以使用 CDN。在这种情况下，很有可能特定的 JavaScript 库已经存在于用户的浏览器缓存中。因此，仅需加载 Web 部件捆绑，该捆绑会使页面加载速度显著加快。

![尽管页面上的两个 Web 部件均使用 Angular，开发人员工具网络选项卡并不显示 Angular](../../../../images/external-scripts-browser-loading-angular-from-cache.png)

上面的示例虽然显示了如何从 CDN 加载 Angular，但未使用公共 CDN。可以在配置中指向任何位置：从公共 CDN、私人托管存储库到 SharePoint 文档库。只要适用于 Web 部件的用户能访问指定的 URL，Web 部件就会按预期工作。

CDN 在全球范围内优化以用于资源快速交付。从公用 CDN 引用脚本的另一个优点是，很有可能相同的脚本已应用到用户过去访问的一些其他网站上。由于该脚本已存在于本地浏览器的缓存中，不需要再特地为 Web 部件进行下载，该脚本将使具有 Web 部件的页面以更快的速度加载。

某些组织不允许从公司网络访问公共 CDN。在这种情况下，使用常用 JavaScript 框架的私人托管存储位置是很好的备用选择。因为组织托管库，它还能控制缓存标头，该标头可以帮助进一步优化资源性能。 

## <a name="javascript-libraries-formats"></a>JavaScript 库格式

不同的 JavaScript 库以不同的方式生成和打包。一些库作为模块打包，其他库是在全局范围中运行的普通脚本（这些脚本通常引用为非 AMD 脚本）。当从 URL 加载 JavaScript 库时，在 SharePoint Framework 项目中注册外部脚本的方式取决于该脚本的格式。存在例如 AMD、UMD 或 CommonJS 等多个模块格式，但只需了解特殊脚本是否是一个模块。

当注册脚本作为模块打包时，必须指定下载特殊脚本的 URL。对其他脚本的依赖项已在脚本模块构造内处理。

另一方面，非 AMD 脚本至少需要下载该脚本的 URL 和在全局范围中注册该脚本的变量名称。如果非 AMD 脚本依赖于其他脚本，它们可以被列为依赖项。为说明这一点，让我们看一下几个示例。

Angular v1.x 是一个非 AMD 脚本。在 SharePoint Framework 项目中通过指定 URL 和全局变量名称将其作为外部资源注册，在注册时应包含以下内容：

```json
"angular": {
  "path": "https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.5.8/angular.min.js",
  "globalName": "angular"
}
```

**globalName** 属性指定的名称必须对应于该脚本使用的名称。这样一来可以将自身正确地展示给可能依赖它的其他脚本。

[ngOfficeUIFabric](http://ngofficeuifabric.com) - Office UI 结构的 Angular 指令是依赖于 Angular 的 UMD 模块。Angular 上的依赖项已在模块中处理，因此为进行注册你需要指定其 URL：

```json
"ng-office-ui-fabric": "https://cdnjs.cloudflare.com/ajax/libs/ngOfficeUiFabric/0.12.3/ngOfficeUiFabric.js"
```

jQuery 是一个 AMD 脚本。若要注册可以仅仅使用：

```json
"jquery": "https://code.jquery.com/jquery-2.2.4.js"
``` 

现在假设需要使用具有 jQuery 插件的 jQuery，该插件本身作为非 AMD 脚本分发。如果已注册这两个脚本，使用：

```json
"jquery": "https://code.jquery.com/jquery-2.2.4.js",
"simpleWeather": {
  "path": "https://cdnjs.cloudflare.com/ajax/libs/jquery.simpleWeather/3.1.0/jquery.simpleWeather.min.js",
  "globalName": "jQuery"
}
```

加载 Web 部件将很可能导致错误：这两个脚本可能并行加载，并且该插件未能利用 jQuery 注册自身。

![使用非 AMD jQuery 插件尝试加载 Web 部件时出现错误](../../../../images/external-scripts-non-amd-no-deps-error.png)

如前文所述，SharePoint Framework 允许指定非 AMD 插件的依赖项。这些依赖项使用 **globalDependencies** 属性进行指定：

```json
"jquery": "https://code.jquery.com/jquery-2.2.4.js",
"simpleWeather": {
  "path": "https://cdnjs.cloudflare.com/ajax/libs/jquery.simpleWeather/3.1.0/jquery.simpleWeather.min.js",
  "globalName": "jQuery",
  "globalDependencies": [ "jquery" ]
}
```

**globalDependencies** 属性指定的每个依赖项都必须指向 **config/config.json** 文件的 **externals** 部分中的另一个相关项。

如果尝试立即生成项目，则会出现另一个错误，这次错误指出不能指定非 AMD 脚本的依赖项。

![尝试使用非 AMD 脚本生成 SharePoint Framework 时出现错误，该非 AMD 脚本指定本为模块的脚本依赖项](../../../../images/external-scripts-non-amd-deps-error.png)

若要解决此问题，需要做的是作为非 AMD 模块注册 jQuery：

```json
"jquery": {
  "path": "https://code.jquery.com/jquery-2.1.1.min.js",
  "globalName": "jQuery"
},
"simpleWeather": {
  "path": "https://cdnjs.cloudflare.com/ajax/libs/jquery.simpleWeather/3.1.0/jquery.simpleWeather.min.js",
  "globalName": "jQuery",
  "globalDependencies": [ "jquery" ]
}
```

这样一来，可以指定 **simpleWeather** 应在 jQuery 后进行加载，并且该 jQuery 应该在全局可用变量 `jQuery` 下可用，这也是 **simpleWeather** jQuery 插件注册自身时所需要的。

> 请注意用于注册 jQuery 的项如何使用 **jquery** 作为外部资源名称，但 **jQuery** 作为了全局变量名称。外部资源名称是在代码的 `import` 语句中使用的名称。这也是必须与 TypeScript 键入相匹配的名称。使用 **globalName** 属性指定的全局变量名称是其他脚本已知的名称，如在库顶部生成的插件。虽然对于一些库来说这些名称可以相同，但这不是必须的，你应该仔细检查，确认你使用的名称正确，以避免出现任何问题。 

## <a name="non-amd-scripts-considerations"></a>非 AMD 脚本注意事项

过去开发的许多 JavaScript 库和脚本作为非 AMD 脚本分发。虽然 SharePoint Framework 支持加载非 AMD 脚本，但也应在任何可能的情况下尽量避免使用它们。

非 AMD 脚本在页面全局范围内注册：一个 Web 部件加载的脚本可用于页面上其他所有的 Web 部件。如果拥有两个使用 jQuery 不同版本的 Web 部件同时作为非 AMD 脚本加载，最后加载的 Web 部件会覆盖以前注册的所有 jQuery 版本。可以想象，这可能会导致不可预知的结果并很难调试仅出现在特定环境下的问题 - 仅与其他 Web 部件在页面上使用不同版本的 jQuery 和仅以特定顺序进行加载时。此模块体系结构通过隔离脚本并防止它们相互影响解决了此问题。

## <a name="when-you-should-consider-bundling"></a>应考虑绑定的时机

将现有 JavaScript 库绑定到 Web 部件可能会导致 Web 部件大文件并导致使用该 Web 部件的页面性能变差。虽然使用 Web 部件通常应避免绑定 JavaScript 库，但有些应用场景中绑定大有益处。

如果正在生成应用于每个内部网的标准解决方案，将所有资源和 Web 部件绑定可以帮助确保 Web 部件将按预期工作。因为不会提前知道解决方案的安装位置，包括在 Web 部件捆绑文件中的所有依赖项将允许其正确运行 - 即使组织不允许从 CDN 或其他外部位置下载资源。

如果解决方案包含大量相互共享某些功能的 Web 部件，作为一个独立库生成共享功能会更好，可在所有 Web 部件中将其作为外部资源引用。这样一来，用户只需要下载公用库一次并在所有 Web 部件对其重用。

## <a name="summary"></a>摘要

当在 SharePoint Framework 上生成客户端 Web 部件时，可以从现有的 JavaScript 库中受益，该库可用于生成功能强大的解决方案。SharePoint Framework 允许将这些库和 Web 部件捆绑在一起，或者作为外部资源进行加载。虽然从 URL 加载现有库是通常推荐的方法，但在一些应用场景中绑定可能会有好处，有必要仔细评估需求以选择满足需要的最佳方法。