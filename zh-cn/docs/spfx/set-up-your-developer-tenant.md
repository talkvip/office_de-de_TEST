# <a name="set-up-your-office-365-tenant"></a>设置 Office 365 租户

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

若要使用 SharePoint Framework 预览版生成并部署客户端 web 部件，将需要一个已启用[首次发布](https://support.office.com/en-us/article/Set-up-the-Standard-or-First-Release-options-in-Office-365-3b3adfa4-1777-4ff0-b606-fb8732101f47)选项的 Office 365 租户。 

## <a name="sign-up-for-an-office-365-tenant"></a>注册 Office 365 租户
如果已经有 Office 365 租户，请参阅[创建你自己的应用目录网站](#create-app-catalog-site)。请确保已启用[首次发布](https://support.office.com/en-us/article/Set-up-the-Standard-or-First-Release-options-in-Office-365-3b3adfa4-1777-4ff0-b606-fb8732101f47)选项。 

如果没有此租户，则注册 [Office 开发人员程序](https://profile.microsoft.com/RegSysProfileCenter/wizardnp.aspx?wizid=14b845d0-938c-45af-b061-f798fbb4d170&lcid=1033)。你将收到带有注册 Office 365 开发人员租户的链接的欢迎邮件。已为 Office 365 开发人员租户启用首次发布选项。 

>**注意：**在注册之前请确保注销任何现有的 Office 365 租户。

## <a name="create-app-catalog-site"></a>创建应用目录网站
将需要应用目录来上载和部署 Web 部件。如果已经设置了一个应用目录，请参阅[创建新的开发人员网站集](#create-a-new-developer-site-collection)。  

通过在浏览器中输入以下 URL，转到 **SharePoint 管理中心**。将 **yourtenantprefix** 替换为你的 Office 365 开发人员租户前缀。
    
```
https://yourtenantprefix-admin.sharepoint.com
```
    
在左侧边栏中，选择“**应用**”菜单项，然后选择“**应用目录**”。

选择“**确定**”来创建新的应用目录站点。

在下一页中，输入以下详细信息：

* **标题**：输入**应用目录**。
* **网站地址_后缀_**：为该应用目录输入首选后缀；例如：**apps**。
* **管理员**：输入用户名并选择“**解析**”按钮来解析用户名。

选择“**确定**”创建应用目录站点。

SharePoint 将创建应用目录站点，并且可以查看 SharePoint 管理中心中的进度。

## <a name="create-a-new-developer-site-collection"></a>创建新的开发人员网站集
还需要开发人员网站。如果已经有一个开发人员网站集，则参阅[设置文档库](#set-up-a-document-library)。

 通过在浏览器中输入以下 URL，转到 **SharePoint 管理中心**。将 **yourtenantprefix** 替换为你的 Office 365 开发人员租户前缀。
    
```
https://yourtenantprefix-admin.sharepoint.com
```
    
在 SharePoint 功能区，选择“**新建**” -> “**专用网站集**”。

在对话框中，输入以下详细信息：

* **标题**：输入开发人员网站集标题；例如：**开发人员网站**。
* **网站地址_后缀_**：为开发人员网站集输入一个后缀；例如：**dev**。
* **模板选择**：选择“**开发人员网站**”作为网站集模板。
* **管理员**：输入用户名并选择“**解析**”按钮来解析用户名。

选择“**确定**”来创建网站集。

SharePoint 将创建开发人员网站，并且可以查看 SharePoint 管理中心中的进度。创建网站后，可以浏览开发人员网站集。

## <a name="set-up-a-document-library"></a>设置文档库
为了使用 SharePoint Framework 的功能，将需要设置具有新列的文档库，并上载 SharePoint 工作台。此过程将在开发人员网站集中使用默认文档库。这在左侧导航中称为**文档**。

* 在浏览器中转到开发人员网站。
* 选择右上角的齿轮图标，然后选择“**网站设置**”以打开设置页面。
* 选择“**网站管理**”类别下方的“**网站库和列表**”。
* 选择“**自定义文档**”
* 选择“**列**”下方的“**创建列**”
* 输入 **ClientSideApplicationId** 作为列名称，其他字段保持其当前值。
* 选择“**确定**”来创建列。

## <a name="put-the-sharepoint-workbench-in-the-document-library"></a>将 SharePoint 工作台置于文档库中
需要开发人员网站上的 SharePoint 工作台来测试 SharePoint 上的 web 部件。此过程将在网站集中使用默认文档库。这在左侧导航中称为**文档**。

* 将 [workbench.aspx](https://github.com/SharePoint/sp-dev-docs/blob/master/workbench.aspx) 文件下载到本地计算机。要完成该任务，请执行以下操作：
    - 打开页面中间文件内容的上下文菜单（右键单击），然后选择“**目标另存为**”或“**链接另存为**”（具体取决于你的浏览器），在本地计算机将文件另存为 **workbench.aspx**。 
* 将文件上载至开发者网站集中的**文档**库。要完成该任务，请执行以下操作：
    - 打开 SharePoint 上的**文档**库。
    - 将 workbench.aspx 拖放到**文档**库。也可以从**文档**库选择“**上载**”，然后找到并上传 workbench.aspx 文件。

## <a name="troubleshooting"></a>疑难解答

如果使用的是使用自助式网站创建设置的 SharePoint 网站，则会发生错误。我们建议使用开发人员网站集，或不是使用自助式网站创建来创建的 SharePoint 网站。

移动到 workbench.aspx 页面时，会收到以下异常信息： 

```
"The requested operation is part of an experimental feature that is not supported in the current environment. The requested operation is part of an experimental feature that is not supported in the current environment." 
```
那么，你所使用的不是已启用首次发布选项的租户。可能使用的是普通租户

在开发人员预览版中，还将需要在网站集中启用“自定义脚本”功能，你将在此网站集中使用客户端 Web 部件。 

## <a name="next-steps"></a>后续步骤
现在，你已配置 SharePoint 租户，请[设置开发环境](./set-up-your-development-environment)以生成客户端 web 部件。
