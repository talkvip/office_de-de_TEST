# <a name="overview-of-the-sharepoint-framework"></a>SharePoint Framework 概述

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

SharePoint Framework (SPFx) 是提供客户端 SharePoint 开发和与 SharePoint 数据简单集成完全支持及开放源代码工具支持的页面和 Web 部件模型。借助 SharePoint Framework，可以在自己首选的开发环境中使用新式 Web 技术和工具生成从一开始便具有可响应性和移动性的生产体验与应用。SharePoint Framework 适用于本地 SharePoint 和 SharePoint Online。
 
SharePoint Framework 的主要功能包括：

* 可在当前用户和浏览器中的连接的上下文中运行。无任何 iFrame。
* 控件在普通页面 DOM 中呈现。
* 控件生成后即可响应并可供访问。
* 允许开发人员访问生命周期 - 其中包括**呈现** -  **加载**、**序列化**和**反序列化**、**配置更改**，等等。
* 其未指定框架。可以使用喜欢的任何浏览器框架：React、Handlebars、Knockout、Angular 等。
* 工具链基于 npm、ypeScript、Yeoman、webpack 和 gulp 等常见开放源代码客户端开发工具。
* 性能可靠。
* 最终用户可以在所有网站上使用租户管理员（或其代理）批准的 SPFx 客户端解决方案，其中包括自助式团队、组或个人网站。 
* 可在经典 Web 部件及发布页面和新版页面中部署解决方案。
 
改进了脚本编辑器 Web 部件上的运行时模型。其中包括强大的客户端 API、处理 SharePoint 和 Office 365 身份验证的 HttpClient 对象、上下文信息、简单的属性定义和配置等。 

如果使用 C# 和 Visual Studio 或 JavaScript，则需要了解一些有关客户端 JavaScript 开发的知识。大部分知识可供完全迁移。根据需要使用相同的 [REST](https://msdn.microsoft.com/en-us/library/office/jj860569.aspx) 服务或 [JavaScript 对象模型 (JSOM)](https://msdn.microsoft.com/en-us/library/office/jj193034.aspx)。数据模型未发生任何形式的更改。如果你是 C# 开发人员，则 TypeScript 是过渡到 JavaScript 世界的不错方式。可以选择 IDE。许多开发人员喜欢使用跨平台 IDE Visual Studio Code，而 Visual Studio 的插件则可用于使用新的开发环境。许多开发人员还使用 Sublime 和 ATOM 等类似产品。使用最适合自己的即可。

## <a name="why-the-sharepoint-framework"></a>为什么要使用 SharePoint Framework？

SharePoint 在 2001 年作为内部部署产品发行。随着时间的推移，一个大型的开发人员社区已经形成并壮大，并从多方面对其进行了改进。总体上，开发人员社区遵循了 SharePoint 产品开发团队所使用的同一种模式和做法，其中包括 Web 部件、SharePoint 功能 XML 等。许多功能使用 C# 编写，被编译到 DLL，然后部署到服务器中。
 
这种解决方案非常适用于只包含一个企业的环境，但是无法扩展到存在多个租户并行运行的云中。因此，我们引入了两种备用模型：客户端 JavaScript 注入和 SharePoint 外接程序。两种解决方案各有优势和劣势。 

### <a name="javascript-injection"></a>JavaScript 注入

SharePoint Online 中最热门的 Web 部件之一是脚本编辑器。可以使用脚本编辑器将 JavaScript 粘贴到 Web 部件并使该 JavaScript 在页面呈现时执行。它虽然简单原始，但很有用。它和页面在同一个浏览器上下文中运行，并在同一个 DOM 中，因此它可与页面上的其他控件交互。相对而言，其性能良好且易于使用。 

但是此方法存在一些弊端。首先，虽然可以打包解决方案以便最终用户可以将该控件放置到页面上，但是无法快捷地提供配置选项。此外，最终用户可以编辑页面并修改脚本，这样则可能会破坏 Web 部件。另一个重大问题在于脚本编辑器 Web 部件未被标记为“**对脚本编写安全**”。大多数自助式网站集（个人网站、团队网站、组网站）都已启用名称为**“NoScript”**的功能。严格来说，它是删除在 SharePoint 中添加/自定义页 (ACP) 权限的功能。这表示会阻止脚本编辑器 Web 部件在这些网站上执行。  

### <a name="sharepoint-add-in-model"></a>SharePoint 外接程序模型

可在 NoScript 网站中运行的解决方案的当前选项是外接程序/应用部件模型。这种实现将创建 **iFrame**，实际体验位于此处并在此执行。其优势在于由于其位于系统外部，且无权访问当前 DOM/连接，因此更易受到信息工作者的信任且更易部署。最终用户可以在 NoScript 网站上安装外接程序。 

但此方法也存在一些弊端。首先，它们在 **iFrame** 中运行。iFrame 比脚本编辑器 Web 部件慢，因为其需要对其他页面的新请求。该页面必须经过身份验证和授权、生成自己的调用以获取 SharePoint 数据、加载各种 JavaScript 库等等。例如，脚本编辑器通常可能需要 100 毫秒进行加载和呈现，而应用部件可能需要 2 秒或更长时间。此外，**iFrame** 边界使得创建响应式设计及继承 CSS 和主题信息更为困难。iFrame 安全性更为强大，这对自己（无法通过页面上的其他控件访问自己的页面时）和最终用户（控件无权访问到 Office 365 的连接时）可能会有一定好处。


### <a name="sharepoint-framework"></a>SharePoint Framework 

之前，我们创建了 Web 部件作为安装到云服务器上的完全信任 C# 程序集。但是，大多数部件的当前开发模型涉及在浏览器中运行 JavaScript 生成对 SharePoint 和 Office 365 后端工作负荷的 REST API 调用。C# 程序集不适用于此环境。我们需要新的开发模型。SharePoint Framework 是 SharePoint 开发的下一个发展方向。

## <a name="whats-next"></a>接下来会是什么？

这是第一个预览版本。我们将基于反馈和体验随时间推移提供更新和优化。在 SharePoint Framework 处于预览状态期间，在 API 名称、流等方面可能会不定期发生重大更改。对 SharePoint Framework 的未来更新将向后兼容，以便你的解决方案可以继续工作。

## <a name="sharepoint-framework-license"></a>SharePoint Framework 许可证

SharePoint Framework 组件通过 [Microsoft EULA](https://github.com/SharePoint/sp-dev-docs/blob/master/Microsoft%20Sharepoint%20Framework%20Preview%20EULA.DOCX) 许可授权。

## <a name="questions"></a>有疑问？

如果存在任何疑问，请将其发布到 [SharePoint StackExchange](http://sharepoint.stackexchange.com/)。请使用 [#spfx](http://sharepoint.stackexchange.com/tags/spfx/)、[#spfx-webparts](http://sharepoint.stackexchange.com/tags/spfx-webparts/) 和 [#spfx-tooling](http://sharepoint.stackexchange.com/tags/spfx-tooling/) 标记你的问题和评论。 

还可在 [GitHub 存储库](https://github.com/SharePoint/sp-dev-docs/issues)中发布有关文档或 SharePoint Framework 方面的问题或反馈。

## <a name="additional-resources"></a>其他资源

- [SharePoint 客户端 Web 部件的概述](./web-parts/overview-client-side-web-parts)
