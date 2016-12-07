# <a name="notes-on-solution-packaging"></a>关于解决方案封装的说明

>**注意：**SharePoint Framework 当前处于预览模式，并可能会发生更改。不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。
>

**package-solution** gulp 任务查看 **/config/package-solution.json** 以了解各种配置详细信息，其中包括一些通用的文件路径，以及包中组件（_Web 部件_ & _应用程序_）之间的关系定义。

## <a name="configuration-file"></a>配置文件

配置文件的架构如下所示：

```javascript
interface IPackageSolutionOptions {
  paths?: IPackageSolutionPathOptions;
  solution: ISolution;
}
```

每个包配置文件均有一些可选的设置来替代任务将在其中查找各种源文件和清单的位置，并定义写入包的位置。此外，它还具有所需的解决方案定义，该定义向包装程序说明各种组件的关系。

### <a name="solution-definition-isolution"></a>解决方案定义 (_ISolution_)

```javascript
interface ISolution {
  name: string;
  id: string;
  version?: string;
  features?: IFeature[];
}
```

每个解决方案文件必须具有**名称**，它标识 SharePoint UI 中的包。此外，每个包必须包含一个全局唯一标识符 (**id**)，后者由 SharePoint 在内部使用。或者，还可以指定格式为“X.X.X.X”的**版本**号，用来标识升级时包的各种版本。

（可选）解决方案定义还包含 SharePoint 功能定义列表。**注意：**如果它被忽略或为空，该任务将为每个组件创建单个功能（1:1 映射）。

### <a name="feature-definition-ifeature"></a>功能定义 (_IFeature_)

```javascript
interface IFeature {
  title: string;
  description: string;
  id: string;
  version?: string;
  componentIds?: string[];
}
```

值得注意的是，这是用于创建 SharePoint 功能的定义，并且其中某些选项将在 UI 中公开。与解决方案类似，每个功能都有强制**标题**、**说明**、**id**，和**版本**号（格式为 X.X.X.X）。请注意，该功能 **id** 也应该是一个全局唯一标识符。

每个功能还可以包含任何数量的组件，当功能被激活时它们也将被激活。这通过 **componentIds** 列表定义，该全局唯一标识符**必须**匹配组件的清单文件中的 **ID**。如果此列表未定义或为空，包装程序将包含功能中的**每个** 组件。

### <a name="file-paths-ipackagesolutionpathoptions"></a>文件路径 (_IPackageSolutionPathOptions_)

```javascript
interface IPackageSolutionPathOptions {
  rawPackageDir: string;
  zippedPackage: string;
  featureXmlDir: string;
}
```

* **rawPackageDir** 是解压缩包文件将在其中写入的目录。
* **zippedPackage** 是压缩包的名称，其中包括扩展名 (.spapp)。
* **featureXmlDir** 是自定义 feature xml 所在的目录。

### <a name="examples"></a>示例

#### <a name="default-configuration"></a>默认配置

```json
{
  "paths": {
    "rawPackageDir": "./sharepoint/solution/debug",
    "zippedPackage": "ClientSolution.spapp",
    "featureXmlDir": "./sharepoint/feature_xml"
  },
  "solution": {
    "name": "A Sample Solution",
    "id": "00000000-0000-0000-0000-000000000000",
    "version": "1.0.0.0"
  }
}
```

### <a name="custom-featurexml"></a>自定义 Feature.xml

为支持不同 SharePoint 资源（如列表模板、页面或内容类型）的设置，自定义 Feature XML 也可能被注入到包中。这用于设置应用程序所需的资源，但也可能用于 Web 部件。有关 Feature XML 的文档位于[此处](https://msdn.microsoft.com/en-us/library/office/ms475601.aspx?f=255&MSPPError=-2147217396)。

默认情况下，包任务在 **./sharepoint/feature_xml** 中查找自定义 Feature XML。此文件夹中的所有文件都包括在最终的应用程序包中。但是，任务依赖于 _rels/ folder 文件夹的内容以确定定义了哪些自定义功能。实质上，它假定每个 **. xml.rels** 文件与 **feature_xml/** 顶层中具有相同名称的 feature.xml 相关，并将关系到添加包中包含此功能的 **AppManifest.xml.rels** 文件。