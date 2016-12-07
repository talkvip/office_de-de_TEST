# <a name="sharepoint-webhooks-sample-reference-implementation"></a>SharePoint webhook 示例参考实现

**参与者**：Bert Jansen (Microsoft)，SharePoint PnP 核心小组

>**注意：**SharePoint webhook 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint webhook。

SharePoint 模式和做法 (PnP) 参考实现显示如何在应用程序中使用 SharePoint webhook。webhook 以企业就绪方式实现，将各种 Microsoft Azure 组件（如 Azure Web 作业、Azure SQL Server 和 Azure 存储队列）用于异步 Web 作业通知处理。

此参考实现仅适用于 [SharePoint 列表 webhook](./lists/overview-sharepoint-list-webhooks)。 

也可以通过观看 [SharePoint PnP YouTube 频道](https://www.youtube.com/watch?v=j3hWCAI9R20)上的视频，按照这些步骤进行操作。

<a href="https://www.youtube.com/watch?v=j3hWCAI9R20">
<img src="../../../images/youtube-introducing-sharepoint-webhooks.png" alt="PnP webcast - Introducing SharePoint webhooks" />
</a>

## <a name="applies-to"></a>适用于

-  使用[已启用的初版](https://support.office.com/en-us/article/Set-up-the-Standard-or-First-Release-options-in-Office-365-3b3adfa4-1777-4ff0-b606-fb8732101f47)的 Office 365 多租户 (MT)。

## <a name="prerequisites"></a>先决条件

Microsoft Azure 用于托管实现 Azure webhook 所需的各种组件。

## <a name="source-code-for-this-reference-implementation"></a>用于此参考实现的源代码

[SharePoint 开发人员示例 GitHub 存储库](https://aka.ms/sp-webhooks-sample-reference)中有用于参考实现的源代码和其他材料。

## <a name="deploying-the-reference-implementation"></a>部署此参考实现

此应用程序将向你展示如何管理专门针对 SharePoint 列表的 webhook。它还包含了webhook 服务终结点的参考实现，可在 webhook 项目中重复使用。 

![SharePoint webhook 参考实现应用程序](../../../images/webhook-sample-application.png)

[SharePoint webhook 参考实现 - 部署指南](https://github.com/SharePoint/sp-dev-samples/blob/master/Samples/WebHooks.List/Deployment%20guide.md)列出了用于部署参考实现的部署步骤。 

## <a name="introduction-to-webhooks"></a>webhook 简介

webhook 通知该应用程序 SharePoint 中应用程序需要监视的的更改。应用程序不再需要定期轮询更改。有了 webhook，只要进行了更改就会通知应用程序（**推送**模型）。webhook 并不特定于 Microsoft。它们是一种通用的 web 标准，同时也被其他供应商（如 WordPress、GitHub、MailChimp，等等）采用。

### <a name="adding-a-webhook-to-your-sharepoint-list"></a>将 webhook 添加到 SharePoint 列表

此参考实现适用于 SharePoint 列表。若要将 webhook 添加到 SharePoint 列表，应用程序首先要通过发送 [`POST /_api/web/lists('list-id')/subscriptions`](./lists/create-subscription) 请求创建一个 webhook 订阅。请求包含以下项：

* 标识要添加 webhook 的列表的有效负载。
* 你的 webhook 服务 URL 发送通知的位置。
* webhook 的到期日期。 

请求 SharePoint 添加 webhook 后，SharePoint 将验证是否存在 webhook 服务终结点。它会将验证字符串发送到服务终结点。SharePoint 期望服务终结点在 5 秒钟内返回验证字符串。如果此过程失败，则取消 webhook 的创建。如果服务已部署，此操作将生效，并且 SharePoint 会针对应用程序最初发送的 POST 请求返回 HTTP 201 消息。响应中的有效负载包含新的 webhook 订阅 ID。

![添加 webhook](../../../images/webhook-sample-add-process.png)

查看此参考实现，你将看到所有 webhook CRUD 操作都合并在 **SharePoint.WebHooks.Common** 项目的 [WebHookManager](https://github.com/SharePoint/sp-dev-samples/blob/master/Samples/WebHooks.List/SharePoint.WebHooks.Common/WebHookManager.cs) 类中。使用 **AddListWebHookAsync** 方法添加 webhook：

```cs
/// <summary>
/// This method adds a webhook to a SharePoint list. Note that you need your webhook endpoint being passed into this method to be up and running and reachable from the internet
/// </summary>
/// <param name="siteUrl">Url of the site holding the list</param>
/// <param name="listId">Id of the list</param>
/// <param name="webHookEndPoint">Url of the webhook service endpoint (the one that will be called during an event)</param>
/// <param name="accessToken">Access token to authenticate against SharePoint</param>
/// <param name="validityInMonths">Optional webhook validity in months, defaults to 3 months, max is 6 months</param>
/// <returns>subscription ID of the new webhook</returns>
public async Task<SubscriptionModel> AddListWebHookAsync(string siteUrl, string listId, string webHookEndPoint, string accessToken, int validityInMonths = 3)
{
    // webhook add code...
}
```

调用 SharePoint 时，需要提供身份验证信息，并且在这种情况下，你所使用的是带**访问令牌**的**载体**身份验证标头。要获得访问令牌，通过 **ExecutingWebRequest** 事件处理程序截获令牌：

```cs
ClientContext cc = null;

// Create SharePoint ClientContext object...

// Add ExecutingWebRequest event handler
cc.ExecutingWebRequest += Cc_ExecutingWebRequest;

// Capture the OAuth access token since we want to reuse that one in our REST requests
private void Cc_ExecutingWebRequest(object sender, WebRequestEventArgs e)
{
    this.accessToken = e.WebRequestExecutor.RequestHeaders.Get("Authorization").Replace("Bearer ", "");
}
```

### <a name="sharepoint-calls-out-to-your-webhook-service"></a>SharePoint 调用 webhook 服务

当 SharePoint 检测到为其创建 webhook 订阅的列表中的更改时，SharePoint 将调用服务终结点。在 SharePoint 中查看有效负载时，请注意以下属性非常重要：

属性|说明
--------|-----------
**subscriptionId**|webhook 订阅 ID。如果要更新 webhook 订阅，例如，延长 webhook 到期日的时间，则需要此 ID。
**resource**|发生更改的列表 ID。
**siteUrl**|托管资源发生更改的站点的服务器相对 URL。

> **注意：**SharePoint 只发送发生更改的通知，但通知不包含实际变化内容。由于获取有关网站和列表发生更改的信息，这意味着可以使用相同的服务终结点来处理多个网站和列表中的 webhook 事件。

调用服务时，务必使服务在 5 秒内响应 HTTP 200 消息。稍后本文将对响应时间进行介绍，但实质上，这要求你**异步**处理通知。在此参考实现中，你将通过使用 Azure Web 作业和 Azure 存储队列来执行此操作。

![SharePoint 调用 webhook 终结点](../../../images/webhook-sample-call-webhook.png)

### <a name="grab-the-changes-your-service-needs-to-act-upon"></a>捕捉服务需要对其进行响应的更改

上一步中已调用服务终结点，但是 SharePoint 只提供了有关发生更改的位置信息，并没有提供实际更改的信息。要了解发生更改的内容，需要使用 SharePoint `GetChanges()` API，如下图所示。

![Async GetChanges](../../../images/webhook-sample-async-getchanges.png)

可以在 **SharePoint.WebHooks.Common** 项目的 [ChangeManager](https://github.com/SharePoint/sp-dev-samples/blob/master/Samples/WebHooks.List/SharePoint.WebHooks.Common/ChangeManager.cs) 类中的 **ProcessNotification** 方法中了解有关 `GetChanges()` 实现的详细信息。 

为避免重复获得相同的更改，务必通知 SharePoint 要从中进行更改的点。此操作通过传递 **changeToken** 完成，这也意味着服务终结点需要保持最后一次使用的 **changeToken**，以便它在下一次调用服务终结点时使用。

以下为有关更改要注意的一些重要事项：

- SharePoint 不会实时调用服务：当具有 webhook 的列表发生更改时，SharePoint 将列出 webhook 调用。此队列每分钟读取一次，并将调用适当的服务终结点。此批处理请求非常重要。例如，如果一次批量上载 1000 条记录，批处理课防止 SharePoint 调用终结点 1000 次。因此终结点仅调用一次，但当你调用 `GetChanges()` 方法时，将获得 1000 个需要处理的更改事件。
- 为保证立即响应，不管更改的数目是多少，服务终结点的工作负荷务必以异步方式运行。在此参考实现中，我们利用 Azure 的强大功能：该服务将序列化传入的负载并将其存储在 Azure 存储队列中，同时 Azure web 作业连续运行，并检查队列中的消息。队列中出现消息时， Web 作业将对其进行处理，同时以异步方式执行逻辑。

### <a name="complete-end-to-end-flow"></a>完整的端到端流

下图描述了完整的端到端 webhook 流：

![webhook 参考实现端到端流](../../../images/webhook-sample-end-to-end-flow.png)

1. 应用程序创建一个 webhook 订阅。执行此操作时，它将从为其创建 webhook 的列表中获取当前 **changeToken**。
2. 应用程序会将 **changeToken** 保留在持久性存储区中，如此种情况下的 SQL Azure。
3. SharePoint 中发生更改，SharePoint 调用服务终结点。
4. 服务终结点将序列化通知请求，并将其存储在存储队列中。
5. Web 作业在队列中查看此消息，并启动消息处理逻辑。
6. 消息处理逻辑从持久性存储区中检索上次使用的更改令牌。
7. 消息处理逻辑使用 `GetChanges()`API 来确定更改内容。
8. 在处理好返回的更改后，现在应用程序将基于这些更改执行所需的操作。
9. 最后，应用程序将保留上次检索的 **changeToken**，以便下次不会收到已处理的更改。

## <a name="how-to-work-with-webhook-renewal"></a>如何使用 webhook 续订

默认情况下，webhook 订阅设置的过期时间为 6 个月，或在创建时指定一个日期。通常需要 webhook 的可用时间更长。如下所述的模式非常适用于增加 webhook 订阅的生存期。第一个模式为轻量型，第二个稍微复杂些，而且需要承载其他 web 作业。

### <a name="basic-model"></a>基本模型

在服务接收通知时，它还获取有关订阅生存期的信息。如果订阅即将过期，只需在通知处理逻辑内延长订阅生存期。此模型在此参考实现中实现，并适用于大多数情况。但是，如果在为其创建 webhook 订阅的列表上连续 6 个月未发生更改，在这种情况下，将不会延长 webhook 订阅的生存期并会将此订阅删除。

### <a name="reliable-but-more-complex-model"></a>可靠但更复杂的模型

每周创建一个 web 作业，以从持久性存储区中读取所有订阅 ID。每次逐一扩展找到的订阅。 

> **注意：**此 web 作业不属于此参考实现。

可以使用 [`PATCH /_api/web/lists('list-id')/subscriptions(‘subscriptionID’)`](./lists/update-subscription) REST 调用，来完成 SharePoint 列表 webhook的实际续订。在参考实现中， webhook 的更新在 **SharePoint.WebHooks.Common** 项目的 [WebHookManager](https://github.com/SharePoint/sp-dev-samples/blob/master/Samples/WebHooks.List/SharePoint.WebHooks.Common/WebHookManager.cs) 类中实现。使用 **UpdateListWebHookAsync** 方法完成 webhook 的更新：

```csharp
/// <summary>
/// Updates the expiration datetime (and notification URL) of an existing SharePoint list webhook
/// </summary>
/// <param name="siteUrl">Url of the site holding the list</param>
/// <param name="listId">Id of the list</param>
/// <param name="subscriptionId">Id of the webhook subscription that we need to update</param>
/// <param name="webHookEndPoint">Url of the webhook service endpoint (the one that will be called during an event)</param>
/// <param name="expirationDateTime">New webhook expiration date</param>
/// <param name="accessToken">Access token to authenticate against SharePoint</param>
/// <returns>true if successful, exception in case something went wrong</returns>
public async Task<bool> UpdateListWebHookAsync(string siteUrl, string listId, string subscriptionId, string webHookEndPoint, DateTime expirationDateTime, string accessToken)
{
    // webhook update code...
}
```

## <a name="debugging-webhooks"></a>调试 webhook

由于 SharePoint 正在调用 Webhook 服务终结点，终结点必须可通过 SharePoint 访问。这使得开发和调试更加复杂一些。下面是可用来使你的生活更加简单的一些策略：

* 在初始开发过程中，为服务处理逻辑提供序列化负载。这将使它能完全测试处理逻辑，而无需部署服务终结点（甚至无需配置 webhook）。
* 如果你对 Azure 资源具有访问权限，可以使用调试版本将你的终结点部署到 Azure，并配置用于调试的 Azure 应用程序服务。然后，这将允许你设置远程断点，然后使用 Visual Studio 执行远程调试。
- 如果不想在开发期间部署服务，则需要针对服务使用安全隧道。也就是说，告诉 SharePoint 通知服务位于共享公用终结点上。在客户端中，安装一个连接到该共享公共服务的组件，只要对公共终结点进行调用，就会通知客户端组件，而该组件会向在本地主机上运行的服务推送有效负载。[ngrok](https://ngrok.com/) 是这样一个安全隧道工具的实现，可用于在本地调试 webhook 服务。
