# <a name="overview-of-sharepoint-webhooks"></a>SharePoint webhook 概述

>**注意：**SharePoint webhook 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint webhook。

SharePoint [webhook](http://en.wikipedia.org/wiki/Webhook) 允许开发人员构建应用程序，对其进行订阅以接收 SharePoint 中发生的特定事件通知。事件被触发时，SharePoint 会向订阅服务器发送 HTTP POST 有效负载。webhook 的开发和使用相比 SharePoint 外接程序远程事件接收器使用的 Windows Communication Foundation (WCF) 服务更加容易。这是因为 webhook 是常规 HTTP 服务 (web API)。

当前 webhook 仅针对 SharePoint 列表项启用。SharePoint 列表项 webhook 覆盖对应于指定 SharePoint 列表或文档库的列表项更改的事件。SharePoint webhook 提供一个简单的通知管道，使应用程序得以了解对 SharePoint 列表的更改，而无需轮询服务。有关详细信息，请参阅 [SharePoint 列表 webhook](./lists/overview-sharepoint-list-webhooks)： 

## <a name="creating-webhooks"></a>创建 webhook
若要创建新的 SharePoint webhook，请将新的订阅添加到特定 SharePoint 资源，如 SharePoint 列表。 

创建新的订阅需要以下信息：

- 资源 - 为其创建订阅的资源端点 URL。例如 SharePoint 列表 API URL。
- 服务器通知 URL - 服务终结点 URL。事件在指定资源中发生时，SharePoint 将为此终结点发送 HTTP POST。
- 到期日期 - 订阅的到期日期。到期日期不应超过 6 个月。默认情况下，订阅设置为从其创建时到 6 个月后过期。 

如果需要，还可以包括以下附加信息：

- 客户端状态 - 传递回所有通知上的客户端的一个不透明字符串。可以使用它对通知进行验证、标记不同的订阅或其他原因。

## <a name="handling-webhook-validation-requests"></a>处理 webhook 验证请求

创建新的订阅时，SharePoint 将验证通知 URL 是否支持接收 webhook 通知。订阅创建请求期间将执行此验证。只有在服务及时响应并返回验证令牌时，才会创建该订阅。

### <a name="example-validation-request"></a>示例验证请求

创建新的订阅时，SharePoint 将以类似于以下示例的格式，将 HTTP POST 请求发送到注册的 URL：


```http
POST https://contoso.azurewebsites.net/your/webhook/service?validationToken={randomString}
Content-Length: 0
```

### <a name="response"></a>响应

若要成功创建订阅，服务必须通过返回 **validationToken** 查询字符串参数值作为纯文本格式的响应对请求作出响应。

```http
HTTP/1.1 200 OK
Content-Type: text/plain

{randomString}
```

如果应用程序返回了一个状态代码而不是 `200`，或无法使用 **validationToken** 参数值做出响应，则创建新订阅的请求将会失败。

## <a name="receiving-notifications"></a>接收通知
通知负载将告知应用程序有关指定订阅的指定资源中发生的事件。如果同个时间段内对资源进行多处更改，应用程序的多个通知可能会分批放到单个请求中。

如果在创建订阅时使用通知负载体，它还将包含客户端状态。

### <a name="webhook-notification-resource"></a>webhook 通知资源

当 SharePoint webhook 通知请求提交到已注册的通知 URL 时，通知资源将定义提供给服务的数据形状。

#### <a name="json-representation"></a>JSON 表示形式

服务生成的每个通知都会序列化为 **webhookNotifiation** 实例：

```json
{
    "subscriptionId":"91779246-afe9-4525-b122-6c199ae89211",
    "clientState":"00000000-0000-0000-0000-000000000000",
    "expirationDateTime":"2016-04-30T17:27:00.0000000Z",
    "resource":"b9f6f714-9df8-470b-b22e-653855e1c181",
    "tenantId":"00000000-0000-0000-0000-000000000000",
    "siteUrl":"/",
    "webId":"dbc5a806-e4d4-46e5-951c-6344d70b62fa"
}
```

由于多个通知可能被提交到单个请求中的服务，它们会一起组合在一个具有单个数组**值**的对象中：

```json
{
   "value":[
      {
         "subscriptionId":"91779246-afe9-4525-b122-6c199ae89211",
         "clientState":"00000000-0000-0000-0000-000000000000",
         "expirationDateTime":"2016-04-30T17:27:00.0000000Z",
         "resource":"b9f6f714-9df8-470b-b22e-653855e1c181",
         "tenantId":"00000000-0000-0000-0000-000000000000",
         "siteUrl":"/",
         "webId":"dbc5a806-e4d4-46e5-951c-6344d70b62fa"
      }
   ]
}
```

#### <a name="properties"></a>属性

| 属性名称          | 类型              | description                                                                                                                         |
|:-----------------------|:------------------|:------------------------------------------------------------------------------------------------------------------------------------|
| **resource**           | 字符串            | 注册订阅的列表的唯一标识符。                                                                 |
| **subscriptionId**     | 字符串            | 订阅资源的唯一标识符。                                                                                 |
| **clientState**        | 字符串 - 可选 | 可选字符串值重新传入此订阅的通知消息。                                     |
| **expirationDateTime** | DateTime          | 如果没有更新或续订，订阅将过期的日期和时间。                                                      |
| **tenantId**           | 字符串            | 生成此通知的租户的唯一标识符。                                                                 |
| **siteUrl**            | 字符串            | 注册订阅的网站的服务器相对 URL。                                                               |
| **webId**              | 字符串            | 注册订阅的 Web 的唯一标识符。                                                                  |

#### <a name="example-notification"></a>示例通知
对服务通知 URL 的 HTTP 请求的正文中将包含一个 webhook 通知。以下示例显示具有一个通知的负载：

```json
{
   "value":[
      {
         "subscriptionId":"91779246-afe9-4525-b122-6c199ae89211",
         "clientState":"00000000-0000-0000-0000-000000000000",
         "expirationDateTime":"2016-04-30T17:27:00.0000000Z",
         "resource":"b9f6f714-9df8-470b-b22e-653855e1c181",
         "tenantId":"00000000-0000-0000-0000-000000000000",
         "siteUrl":"/",
         "webId":"dbc5a806-e4d4-46e5-951c-6344d70b62fa"
      }
   ]
}
```

通知并不包括任何有关触发它的更改的信息。应用程序应使用列表上的 [GetChanges API](https://msdn.microsoft.com/EN-US/library/office/dn531433.aspx#bk_ListGetChanges)，以在通知应用程序时查询更改日志中的更改集合，并存储更改令牌值用于任何后续调用。

## <a name="event-types"></a>事件类型
SharePoint webhook 仅支持异步事件。这意味着在发生更改后才触发 webhook（类似于 **-ed** 事件），因此不可能同步（**-ing** 事件）。

## <a name="error-handling"></a>错误处理
如果将通知发送到应用程序时发生错误，SharePoint 将遵循指数回退逻辑。200-299 范围外任何具有 HTTP 状态代码的响应或超时的响应，将在接下来的几分钟内再次尝试。如果该请求在 15 分钟后未成功，将丢弃该通知。

尽管该服务在检测到足够数量的故障时可能会删除订阅，但仍然会为应用程序对以后的通知进行尝试。

## <a name="expiration"></a>过期
默认情况下，如果未指定 **expirationDateTime** 值，webhook 订阅设置为 6 个月后过期。 

创建订阅时需要设置到期日期。到期日期不应超过 6 个月。应用程序需要根据应用程序需要，通过定期更新订阅来处理到期日期。 

## <a name="retry-mechanism"></a>重试机制

如果事件发生在 SharePoint 中并且服务端点不可访问（例如，出于维护原因），SharePoint 将重试发送通知。SharePoint 将重试 **5 次，每次尝试之间等待 5 分钟**。如果由于某种原因服务停用时间较长，那么下一次在它处于活动状态并通过 SharePoint 调用时，对 `GetChanges()` 的调用将恢复服务不可用时遗漏的更改。