# <a name="sharepoint-framework-development-tools-and-libraries"></a>SharePoint Framework 开发工具和库

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

SharePoint Framework 包含多个可用于生成解决方案的客户端 JavaScript 库。本文提供了可用于开发客户端 Web 部件的工具和库的概述。

## <a name="typescript"></a>TypeScript
TypeScript 是键入的 JavaScript 超集，编译为纯 JavaScript。SharePoint 客户端开发工具使用 TypeScript 类、模块和接口生成。可以使用这些工具来构建可靠的客户端 Web 部件。 

若要开始使用 TypeScript，请参阅以下资源：

* [TypeScript 快速入门](https://www.typescriptlang.org/docs/tutorial.html)
* [TypeScript 体育场](https://www.typescriptlang.org/play/index.html)
* [TypeScript 手册](https://www.typescriptlang.org/docs/handbook/basic-types.html)
* [Stack Overflow 上的 TypeScript 社区](https://stackoverflow.com/questions/tagged/typescript)

## <a name="javascript-frameworks"></a>JavaScript 框架
可以选择多个 JavaScript 框架中的任意一个来开发客户端 Web 部件。以下是一些最为热门的框架：

* [React](https://facebook.github.io/react/)
* [AngularJS 1.x](https://docs.angularjs.org/tutorial)
* [Angular 2 for TypeScript 2.x](https://angular.io/docs/ts/latest/quickstart.html)
* [Handlebars](http://handlebarsjs.com/)

由于客户端 Web 部件为置于 SharePoint 页面的组件，我们建议选择一个支持类似组件模型的 JavaScript 框架。轻量级框架（如 React、Handlebars 和 Angular 2）均支持组件模型，很适合构建客户端 Web 部件。 

同时，我们建议查看 [SharePoint PnP JavaScript 核心库](https://github.com/SharePoint/PnP-JS-Core)，这是社区推动的针对在 SharePoint REST API 上提供便捷访问所做的努力。 

## <a name="node-package-manager-npm"></a>节点包管理器 (npm)

SharePoint 客户端开发工具使用与 [NuGet](https://www.nuget.org/) 类似的 [npm](https://www.npmjs.com/) 软件包管理器，以管理依赖项和其他所需的 JavaScript 帮助程序。npm 通常是 Node.js 安装程序的一部分。

有关 npm 的详细信息，请参阅 [npm 文档](https://docs.npmjs.com/)。

## <a name="nodejs"></a>Node.js

Node.js 是承载和提供 JavaScript 代码的开放源代码、跨平台运行时环境。可以使用以 JavaScript 编写的 node.js 开发服务器端 Web 应用程序。Node.js 生态系统与 npm 和 gulp 等任务运行程序紧密结合，以提供有效环境构建基于 JavaScript 的应用程序。Nodel.js 类似于 IIS Express 或 IIS，但是包括简化客户端开发的工具。 

有关 Node.js 的详细信息，请参阅以下内容：

* [关于 Node.js](https://nodejs.org/en/about/)
* [Node.js API 参考文档](https://nodejs.org/api/)
* [Node.js 用法和示例](https://nodejs.org/api/synopsis.html)

## <a name="gulp-task-runner"></a>gulp 任务运行程序
SharePoint 客户端开发工具使用 [gulp](http://gulpjs.com/) 作为生成过程任务运行程序，以执行以下操作：

* 捆绑和缩小 JavaScript 与 CSS 文件。
* 在每个生成前运行工具以调用捆绑和缩小任务。
* 将 LESS 或 SASS 编译为 CSS。
* 将 TypeScript 文件编译为 JavaScript。

有关 gulp 的详细信息，请参阅以下内容：

* [Gulp 入门](https://github.com/gulpjs/gulp/blob/master/docs/getting-started.md)
* [TypeScript 和 Gulp](https://www.typescriptlang.org/docs/handbook/gulp.html)
* [关于 Gulp 的文章](https://github.com/gulpjs/gulp/blob/master/docs/README.md#articles)

## <a name="webpack"></a>Webpack

Webpack 是一个模块程序包，它将 Web 部件文件作为依赖项，并生成一个或多个 JavaScript 捆绑，以便为不同的方案加载不同的程序包。

开发工具链使用 [CommonJS](https://webpack.github.io/docs/commonjs.html) 进行构建。这便于你对模块以及其使用位置进行定义。工具链还使用通用的模块加载器 [SystemJS](https://github.com/systemjs/systemjs) 来加载模块。通过确保每个 Web 部件在其各自的命名空间中进行执行，有助于确定 Web 部件的范围。

有关 webpack 的详细信息，请参阅以下内容：

* [Webpack 文档](http://webpack.github.io/docs/what-is-webpack.html)
* [TypeScript、React 和 Webpack](https://www.typescriptlang.org/docs/handbook/react-&-webpack.html)

## <a name="yeoman-generators"></a>Yeoman 生成器
[Yeoman](http://yeoman.io/) 可帮助启动新项目、规定的最佳做法和工具以帮助你保持高效率。SharePoint Yeoman 生成器将可作为框架的一部分，来启动新的客户端 Web 部件项目。一旦生成项目，可以使用所选的 IDE（如 [Visual Studio](https://www.visualstudio.com/products/visual-studio-community-vs)）或 HTML/JavaScript 代码编辑器（如 [Visual Studio Code](https://code.visualstudio.com/)、[Sublime Text](https://www.sublimetext.com/) 或 [Atom](https://atom.io/)）。  

有关 Yeoman 的详细信息，请参阅以下内容：

* [使用 Yeoman 为 Web 应用程序提供基架](http://yeoman.io/codelab/index.html)
* [可用的 Yeoman 生成器列表](http://yeoman.io/generators/)

以下是一些可以尝试的常见 Yeoman 生成器，具体取决于你选择的框架：

* [generator-react-webpack](https://github.com/newtriks/generator-react-webpack)
* [generator-angular](https://www.npmjs.com/package/generator-angular)

## <a name="sharepoint-rest-apis"></a>SharePoint REST API

SharePoint Framework 提供与 SharePoint 体验的关键集成，并面向 Web 开发。SharePoint REST API 使你能够与 SharePoint 以及其他使 Web 部件功能成形的工作负荷进行交互。 

我们建议你熟悉以下 REST API 集：

* [SharePoint 列表 REST API](https://msdn.microsoft.com/EN-US/library/office/dn292552.aspx)

## <a name="patterns-and-practices"></a>模式和做法

[Office 开发模式和做法 / SharePoint 模式和做法 (PnP)](http://aka.ms/officedevpnp) 主动提供代码示例、模式和其他资源，以帮助你将现有解决方案转换为 SharePoint Framework。请务必熟悉代码示例和通过 PnP 努力提供的指导。

## <a name="additional-resources"></a>其他资源

* [SharePoint Framework](sharepoint-framework-overview)
* [构建 Hello World 客户端 Web 部件](web-parts/get-started/build-a-hello-world-web-part)
