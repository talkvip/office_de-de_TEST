# <a name="add-an-external-library-to-your-sharepoint-client-side-web-part"></a>将外部库添加到 SharePoint 客户端 Web 部件

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

Web 部件中必须包含一个或多个 JavaScript 库。本文介绍了如何捆绑外部库和共享库。


## <a name="bundling-a-script"></a>捆绑脚本

默认情况下，Web 部件捆绑程序会自动包含所有属于 Web 部件模块依赖项的库。这表示将在同一个 JavaScript 捆绑文件中部署库作为你的 Web 部件。这对于不会在多个 Web 部件中使用的小型库更有帮助。

### <a name="example"></a>示例

将字符串验证库[验证程序](https://www.npmjs.com/package/validator)包添加到 Web 部件。

从 npm 下载验证程序包：
    
```
npm install validator --save
```
    
>**注意：**因为你使用的是 TypeScript，所以需要要添加的包的键入。编写代码时这一步很重要，因为 TypeScript 只是 JavaScript 的一个超集。编译时所有的 TypeScript 代码仍会被转换为 JavaScript 代码。可以使用 **tsd** 包搜索和查找键入，例如 `tsd install {package} --save`
    
在 Web 部件中名为 `validator.d.ts` 的文件夹中创建一个文件并添加以下内容：
    
>**注意：**一些库没有键入。验证程序也是它们中的其中之一。这种情况下需要为库定义自己的键入定义 `.d.ts` 文件。以下代码是一个示例。
    
```typescript
declare module "validator" {
    export function isEmail(email: string): boolean;
    export function isAscii(text: string): boolean;
}
```
    
在 Web 部件文件中导入键入：
    
```typescript
import * as validator from 'validator';
```
    
在 Web 部件代码中使用验证程序库：
    
```typescript
validator.isEmail('someone@example.com');
```

## <a name="sharing-a-library-among-multiple-webparts"></a>在多个 Web 部件中共享库

客户端解决方案可能包含多个 Web 部件。这些 Web 部件可能需要导入或共享同一个库。在这种情况下，不能捆绑库，而应该将库添加到单独的 JavaScript 文件中以提升性能并增强缓存功能。此操作尤其适用于大型库。

### <a name="example"></a>示例

在此示例中要将 [marked](https://www.npmjs.com/package/marked) 包 - 一种标记编译器 - 共享到单独的捆绑包中。

从 npm 下载 **marked** 包：
    
```
npm install marked --save
```
    
下载键入：
    
```
tsd install marked --save
```
    
编辑 **config/config.json** 并向 **externals** 映射添加项。此操作可以指示捆绑程序将其放入单独的文件中。此操作可以阻止捆绑程序绑定该库：
    
```json
"marked": "node_modules/marked/marked.min.js"
```
    
目前我们已经添加了库的包和键入，则可添加语句将 `marked` 库导入 Web 部件。
    
```typescript
import * as marked from 'marked';
```
     
在 Web 部件中使用库：
    
```typescript
console.log(marked('I am using __markdown__.'));
```

## <a name="loading-a-script-from-a-cdn"></a>从 CDN 中加载脚本

如果不从 npm 包中加载库，则可以从 CDN 中加载脚本。若要实现此操作，可以编辑 **config.json** 文件来通过库的 CDN URL 加载库。

### <a name="example"></a>示例

在此示例中要从 CDN 加载 jQuery。无需安装 npm 包。但是仍然需要安装键入。 

安装 jQuery 的键入：
    
```
tsd install jquery --save
```
    
更新 `config` 文件夹中的 `config.json` 以从 CDN 加载 jQuery。向 `externals` 字段添加项：
    
```json
"jquery": "https://code.jquery.com/jquery-3.1.0.min.js"
```
    
将 jQuery 导入到 Web 部件：
    
```typescript
import * as $ from 'jquery';
```
    
在 Web 部件中使用 jQuery：
    
```javascript
alert( $('#foo').val() );
```

## <a name="loading-a-non-amd-module"></a>加载非 AMD 模块

一些脚本遵循旧的 JavaScript 模式将库存储在全局命名空间上。此模式目前已被弃用，现支持使用[通用模块定义 (UMD)](https://github.com/umdjs/umd)/[异步模块定义 (AMD)](https://en.wikipedia.org/wiki/Asynchronous_module_definition) 或 [ES6 模块](https://github.com/lukehoban/es6features/blob/master/README.md#modules)但是需要将这些库加载到 Web 部件中。 

若要加载非 AMD 模块，可以向 **config.json** 文件中的项中添加附加属性。

### <a name="example"></a>示例

在此示例中要从 Contoso 的 CDN 中加载虚拟的非 AMD 模块。这些步骤适用于 src/ 或 node_modules/ 目录中的所有非 AMD 脚本。

此脚本名称为 **Contoso.js** 且存储在 **https://contoso.com/contoso.js** 上。其内容为：

```javascript
var ContosoJS = {
  say: function(text) { alert(text); },
  sayHello: function(name) { alert('Hello, ' + name + '!'); }
};
```


在 Web 部件文件夹中名称为 **contoso.d.ts** 的文件中创建此脚本的键入。
    
```typescript
declare module "contoso" {
    interface IContoso {
        say(text: string): void;
        sayHello(name: string): void;
    }
    var contoso: IContoso;
    export = contoso;
}
```
    
更新 **config.json** 文件夹以将该脚本包含在内。向 **externals** 映射添加项：
    
```json
{
    "contoso": {
        "path": "https://contoso.com/contoso.js",
        "globalName": "ContosoJS"
    }
}
```
    
向 Web 部件代码添加导入项：
    
```typescript
import contoso from 'contoso';
```
    
在代码中使用 contoso 库：
    
```typescript
contoso.sayHello(username);
```

## <a name="loading-a-library-that-has-a-dependency-on-another-library"></a>加载对其他库存在依赖关系的库

许多库对其他库存在依赖关系。可以使用 **globalDependencies** 属性在 **config.json** 文件中指定此类依赖关系。

请注意，无需为非 AMD 模块指定该字段；它们可以正确地相互导入。但是务必注意，非 AMD 模块将 AMD 模块作为依赖项。

有两个相关示例。

#### <a name="non-amd-module-has-a-non-amd-module-dependency"></a>非 AMD 模块具有非 AMD 模块依赖项

此示例涉及两个虚拟脚本。它们位于 **src/** 文件夹中，但是也可从 CDN 中加载它们。

**ContosoUI.js**

```javascript
Contoso.EventList = {
    alert: function() {
        var events = Contoso.getEvents();
        events.forEach( function(event) {
            alert(event);
        });
    }
}
```

**ContosoCore.js**

```javascript
var Contoso = {
    getEvents: function() {
        return ['A', 'B', 'C'];   
    }
};
```

添加或创建此类的键入。在此示例中将创建 `Contoso.d.ts`，其中会包含两个 JavaScript 文件的键入。 
    
**contoso.d.ts**

```typescript
declare module "contoso" {
interface IEventList {
 alert(): void;
}
interface IContoso {
 getEvents(): string[];
 EventList: IEventList;
}
var contoso: IContoso;
export = contoso;
}
```

更新 **config.json** 文件。向 **externals** 添加两个项：
    
```json
{
     "contoso": {
         "path": "/src/ContosoCore.js",
         "globalName": "Contoso"
     },
     "contoso-ui": {
         "path": "/src/ContosoUI.js",
         "globalName": "Contoso",
         "globalDependencies": ["contoso"]
     }
}
```
    
添加 Contoso 和 ContosoUI 的导入项：
       
```typescript
import contoso = require('contoso');
require('contoso-ui');
```

在代码中使用库：
    
```typescript
contoso.EventList.alert();
```

## <a name="loading-sharepoint-jsom"></a>加载 SharePoint JSOM

加载 SharePoint JSOM 与加载具有依赖项的非 AMD 脚本本质上属于同一方案。这表示同时使用 **globalName** 和 **globalDependency** 两个选项。

安装属于 JSOM 键入依赖项的 Microsoft Ajax 的键入：

```
tsd install microsoft.ajax --save
```

安装 JSOM 的键入：

```
tsd install sharepoint --save
``` 

向 `config.json` 添加项：

```json
{
"sp-init": {
 "path": "https://CONTOSO.sharepoint.com/_layouts/15/init.js",
 "globalName": "$_global_init"
},
"microsoft-ajax": {
 "path": "https://CONTOSO.sharepoint.com/_layouts/15/MicrosoftAjax.js",
 "globalName": "Sys",
 "globalDependencies": [ "sp-init" ]
},
"sp-runtime": {
 "path": "https://CONTOSO.sharepoint.com/_layouts/15/SP.Runtime.js",
 "globalName": "SP",
 "globalDependencies": [ "microsoft-ajax" ]
},
"sharepoint": {
 "path": "https://CONTOSO.sharepoint.com/_layouts/15/SP.js",
 "globalName": "SP",
 "globalDependencies": [ "sp-runtime" ]
}
}
```

在 Web 部件中添加 require 语句：
    
```typescript
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
```

## <a name="load-localized-resources"></a>加载本地化资源

加载本地化资源很简单。**config.json** 中存在一个名为 **localizedResources** 的映射，你可以使用它描述如何加载本地化资源。此映射中的路径与 **lib** 文件夹相对，其中不能包含前导斜杠 (**/**)。

在此示例中具有文件夹 **src/strings/**。在此文件夹中是一些名称为 **en-us.js**、**fr-fr.js**、**de-de.js** 的 JavaScript 文件。因为这些文件中的每一个都必须可通过模块加载程序加载，因此它们必须包含一个 CommonJS 包装器。例如，在 **en-us.js** 中：

```typescript
  define([], function() {
    return {
      "PropertyPaneDescription": "Description",
      "BasicGroupName": "Group Name",
      "DescriptionFieldLabel": "Description Field"
    }
  });
```

编辑 **config.json** 文件。向 **localizedResources** 添加项。**{locale}** 是区域名称的占位符标记：

```json
{
"strings": "strings/{locale}.js"
}
```
    
添加字符串的键入。在此示例中具有文件 **MyStrings.d.ts**：

```typescript
declare interface IStrings {
webpartTitle: string;
initialPrompt: string;
exitPrompt: string;
}

declare module 'mystrings' {
const strings: IStrings;
export = strings;
}
```
    
添加项目中字符串的导入项：
    
```typescript
import * as strings from 'strings';
```
    
在项目中使用字符串：

```typescript
alert(strings.initialPrompt);
```
