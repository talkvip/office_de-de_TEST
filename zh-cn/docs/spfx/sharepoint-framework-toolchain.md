# <a name="sharepoint-framework-toolchain"></a>SharePoint Framework 工具链

>**注意：**SharePoint Framework 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

## <a name="overview"></a>概述
SharePoint Framework 工具链是工具、框架包和其他项目的集合，用于生成和部署客户端项目的管理。工具链可以帮助你生成客户端组件，如 Web 部件。它还可以帮助你在本地开发环境中使用 SharePoint 工作台等工具对其进行测试。并且可以使用工具链打包和部署到 SharePoint。工具链还提供了一套帮助你完成密钥生成任务的生成命令，例如代码编译、将客户端项目打包到 SharePoint 应用程序包等等。 

## <a name="npm"></a>npm
在深入探讨各种工具链组件之前，请务必了解 SharePoint Framework 使用 [npm](https://www.npmjs.com/) 的方式，以管理项目中的不同模块。npm 是 JavaScript 客户端开发首选的开放源代码包管理者之一。 

一个典型的 npm 包包含一个或多个可重复使用的称为模块的 JavaScript 代码文件，以及它所依赖的依赖项包列表。安装此包时，npm 还会安装这些依赖项。该官方 [npm 注册表](https://www.npmjs.com/)包含数百个可下载用来生成应用程序的包。还可以[发布自己的包](https://docs.npmjs.com/getting-started/what-is-npm)到 npm，并与其他开发人员共享。SharePoint Framework 不仅在工具链中使用一些 npm 包，也会发布它[自己的包到 npm 注册表](https://www.npmjs.com/search?q=%40microsoft%2Fsp-)。 

### <a name="sharepoint-framework-packages"></a>SharePoint Framework 包
SharePoint Framework 由多个协同工作的 npm 包组成，以帮助开发人员在 SharePoint 中构建客户端体验。 

SharePoint Framework 中有以下包：

* [@microsoft/generator-sharepoint](https://www.npmjs.com/package/@microsoft/generator-sharepoint) - 与 SharePoint Framework 一起使用的 [Yeoman](http://yeoman.io/) 插件。使用此生成器，开发人员可以快速设置具有合理默认值和最佳做法的新的客户端 Web 部件项目。

* [@microsoft/sp-client-base](https://www.npmjs.com/package/@microsoft/sp-client-base) - 为使用 SharePoint Framework 构建的客户端应用程序定义核心库

* [@microsoft/sp-client-preview](https://www.npmjs.com/package/@microsoft/sp-client-preview) - 包含仍在开发中的 SharePoint Framework 库。完成后，它们将被重构到单独的 npm 包中。

* [@microsoft/sp-webpart-workbench](https://www.npmjs.com/package/@microsoft/sp-webpart-workbench) - SharePoint 工作台是用于测试和调试客户端 Web 部件的独立环境。

* [@microsoft/sp-module-loader](https://www.npmjs.com/package/@microsoft/sp-module-loader) - 管理客户端组件、Web 部件和其他资产的版本和加载的模块加载器。它还提供基本的诊断服务。它基于 SystemJS 和 WebPack 等熟悉的标准生成，并且是在页面上加载 SharePoint Framework 的第一个部分。

* [@microsoft/sp-module-interfaces](https://www.npmjs.com/package/@microsoft/sp-module-interfaces) - 定义 SharePoint Framework 模块加载器以及生成系统共享的几个模块接口。

* [@microsoft/sp-lodash-subset](https://www.npmjs.com/package/@microsoft/sp-lodash-subset) - 提供自定义 [lodash](https://lodash.com/) 程序包，与 SharePoint Framework 的模块加载器结合使用。为提高运行时性能，它只包含最基本 lodash 功能的一个子集。

* [@microsoft/sp-tslint-rules](https://www.npmjs.com/package/@microsoft/sp-tslint-rules) - 定义自定义 tslint 规则，与 SharePoint 客户端项目结合使用。

* [@microsoft/office-ui-react-bundle](https://www.npmjs.com/package/@microsoft/office-ui-fabric-react-bundle) - 提供经优化的自定义 [office-ui-fabric-react](https://www.npmjs.com/package/office-ui-fabric-react) 程序包，以与 SharePoint Framework 的模块加载器结合使用。

### <a name="common-build-tools-packages"></a>常见的生成工具包
除了前面所述的 SharePoint Framework 包，还有一套通用的生成工具可用于执行生成任务，例如将 TypeScript 代码编译为 JavaScript，以及将 SCSS 转换为 CSS。

SharePoint Framework 中有以下常见的生成工具包：

* [@microsoft/sp-build-core-tasks](https://www.npmjs.com/package/@microsoft/sp-build-core-tasks) - 基于 gulp 的 SharePoint Framework 生成系统的任务集合。 `sp-build-core-tasks` 包实现特定于 SharePoint 的操作，如包解决方案和编写清单。

* [@microsoft/sp-build-web](https://www.npmjs.com/package/@microsoft/sp-build-web) - 导入并配置一套适用于将在 Web 浏览器（而不是 Node.js 环境）中运行的生成目标的生成任务。此包旨在通过使用 SharePoint Framework 生成系统的 gulp文件直接导入。 

* [@microsoft/gulp-core-build](https://www.npmjs.com/package/@microsoft/gulp-core-build) - 用于生成 TypeScript、HTML、less 和其他生成格式的核心 gulp 生成任务。此包依赖于其他几个包含以下任务的 npm 包：
    - [@microsoft/gulp-core-build-typescript](https://www.npmjs.com/package/@microsoft/gulp-core-build-typescript)
    - [@microsoft/gulp-core-build-sass](https://www.npmjs.com/package/@microsoft/gulp-core-build-sass)
    - [@microsoft/gulp-core-build-webpack](https://www.npmjs.com/package/@microsoft/gulp-core-build-webpack)
    - [@microsoft/gulp-core-build-serve](https://www.npmjs.com/package/@microsoft/gulp-core-build-serve)
    - [@microsoft/gulp-core-build-karma](https://www.npmjs.com/package/@microsoft/gulp-core-build-karma)
    - [@microsoft/gulp-core-build-mocha](https://www.npmjs.com/package/@microsoft/gulp-core-build-mocha)

* [@microsoft/loader-cased-file](https://www.npmjs.com/package/@microsoft/loader-cased-file) - webpack 的[文件加载程序](https://www.npmjs.com/package/file-loader)包装，可用于修改生成文件名的大小写。这可用于某些情况，如使用内容传递网络 (CDN) 时，只允许小写文件名。

* [@microsoft/loader-load-themed-styles](https://www.npmjs.com/package/@microsoft/loader-load-themed-styles) - 包装脚本中等效于 <code>require('load-themed-styles').loadStyles( /* css text */ )</code> 的 CSS 加载的加载程序。它旨在代替样式加载程序。

* [@microsoft/loader-raw-script](https://www.npmjs.com/package/@microsoft/loader-raw-script) - 此加载程序将直接在 webpack 包中使用 `eval(…)` 加载脚本文件的内容。

* [@microsoft/loader-set-webpack-public-path](https://www.npmjs.com/package/@microsoft/loader-set-webpack-public-path) -此加载程序将 __webpack_public_path__ 变量设置为自变量中可以选择附加到 SystemJs baseURL 属性的指定值。

## <a name="scaffolding-a-new-client-side-project"></a>构建新的客户端项目
SharePoint 生成器使用 Web 部件构建客户端项目。该生成器还为指定的客户端项目下载并配置所需的工具链组件。  

### <a name="packages-installation"></a>包安装
该生成器在项目文件夹中本地安装所需的 npm 包。npm 允许在本地或全局范围内将包安装到你的项目。这对于两者都有好处，但是如果你的代码依赖于这些包模块，[一般指南](https://nodejs.org/en/blog/npm/npm-1-0-global-vs-local-installation/)是在本地安装 npm 包。对于 Web 部件项目，Web 部件代码取决于各种 SharePoint 和公共生成包，因此需要本地安装这些包。 

本地安装包时，npm 还将安装与每个包关联的依赖项。可以在项目文件夹中查找安装在 `node_modules` 文件夹下的包。此文件夹包含包及其所有依赖项。此文件夹包含数十到数百个文件夹为理想情况，因为 npm 包将始终分成较小的模块，因此造成数十到数百个正在安装的包。SharePoint Framework 包位于 `node_modules\@microsoft` 文件夹下方。`@microsoft` 是共同代表[由 Microsoft 所发布包](https://www.npmjs.com/~microsoft)的 npm 范围。

每次使用该生成器创建新项目时，生成器将为此特定项目本地安装 SharePoint Framework 包及其依赖项。通过此方式，npm 将更容易管理 Web 部件项目，而不会影响本地开发环境中的其他项目。 

### <a name="packagejson"></a>package.json
客户端项目中的 `package.json` 文件指定项目所依赖的依赖项列表。列表定义了要安装的依赖项。如上文所述，每个依赖项可能包含若干个。npm 允许你使用 `dependencies` 和 `devDependencies` 属性为包定义运行时和生成依赖项。在构建 Web 部件的情况下，当你想要在代码中使用该模块时，将用到 `devDependencies` 属性。

下面是 [helloworld-webpart](web-parts/get-started/build-a-hello-world-web-part) 的 `package.json`：

```json 
{
  "name": "helloword-webpart",
  "version": "0.0.1",
  "private": true,
  "engines": {
    "node": ">=0.10.0"
  },
  "dependencies": {
    "@microsoft/sp-client-base": "~0.1.12",
    "@microsoft/sp-client-preview": "~0.1.12"
  },
  "devDependencies": {
    "@microsoft/sp-build-web": "~0.4.32",
    "@microsoft/sp-module-interfaces": "~0.1.8",
    "@microsoft/sp-webpart-workbench": "~0.1.12",
    "gulp": "~3.9.1"
  },
  "scripts": {
    "build": "gulp bundle",
    "clean": "gulp nuke",
    "test": "gulp test"
  }
}
```

虽然有大量为项目安装的包，它们只需在开发环境中用于构建 web 部件。在这些包的帮助下，你将能够依赖于模块，并生成、编译、捆绑及打包 Web 部件，以进行部署。部署到 CDN 服务器或 SharePoint 的最终缩减捆绑在一起的 Web 部件版本不包括这些包。也就是说，可以根据你的要求进行配置以包括某些模块。有关详细信息，请参阅[将外部库添加到 Web 部件](web-parts/basics/add-an-external-library)。

### <a name="working-with-source-control-systems"></a>使用源代码管理系统
随着项目依赖项的增加，要安装的包数目同样会增加。不想将包含所有依赖项的 `node_modules` 文件夹签入到源代码管理系统。应在签入期间将 `node_modules` 从要忽略的文件列表中排除。 

如果正在使用 `git` 作为源代码管理系统，Yeoman 基架的 Web 部件项目还将包含一个排除 `node_modules` 文件夹的 `.gitignore` 文件。签出或克隆时，来自源代码管理系统的 Web 部件项目第一次运行命令以初始化并在本地安装所有项目依赖项：

```
npm i
```

npm 将扫描 `package.json` 文件并安装所需的依赖项。 

## <a name="build-tasks"></a>生成任务
SharePoint Framework 使用 [gulp](http://gulpjs.com/) 作为其任务运行程序来处理流程，如下所示：

* 捆绑和缩小 JavaScript 与 CSS 文件。
* 在每个生成前运行工具以调用捆绑和缩小任务。
* 将 LESS 或 SASS 文件编译为 CSS。
* 将 TypeScript 文件编译为 JavaScript。

工具链包含以下在 [@microsoft/sp-build-core-tasks](https://www.npmjs.com/package/@microsoft/sp-build-core-tasks) 包中定义的 gulp 任务：

* build
  * 生成客户端解决方案项目。
* bundle
  * 将客户端解决方案项目入口点及其所有依赖项捆绑到一个 JavaScript 文件中。
* serve
  * 为客户端解决方案项目和本地计算机中的资产提供服务。
* nuke
  * 清除上一个生成和生成目标目录（lib 和 dist）中的客户端解决方案项目的生成项目
* test
  * 如果可用，则为客户端解决方案项目运行单元测试。 
* package-solution
  * 将客户端解决方案打包到 SharePoint 包。
* deploy-azure-storage
  * 将客户端解决方案项目资产部署到 Azure 存储中。 

若要启动不同的任务，使用 gulp 命令追加任务名称。例如，要进行编译并预览 SharePoint 工作台中的 Web 部件，请运行以下命令：

```
gulp serve
```

>**注意**：不能同时执行多项任务。

`serve` 运行不同的任务，并最终启动 SharePoint 工作台。

![gulp 服务任务](../../images/toolchain-gulp-serve-task.png)

### <a name="build-targets"></a>生成目标
在上面的屏幕快照中，可以看到指示生成目标的任务，如下所示：

```
Build target: DEBUG
```

如果未指定任何参数，则命令指向 BUILD 模式。如果未指定 `ship` 参数，则命令指向 SHIP 模式。

通常情况下，当 Web 部件准备好传送或在生产服务器中进行部署时，目标将会指向 SHIP。对于像测试和调试的其他情况，目标将指向 BUILD。SHIP 目标也可以确保 Web 部件包缩小版本的生成。 

要针对 SHIP 模式，请使用 `--ship` 追加任务：

```
gulp --ship
```

在 DEBUG 模式下，生成任务将所有 Web 部件资产（包括 Web 部件包）复制到 `dist` 文件夹。

在 SHIP 模式下，生成任务将所有 Web 部件资产（包括 Web 部件包）复制到 `temp\deploy` 文件夹。