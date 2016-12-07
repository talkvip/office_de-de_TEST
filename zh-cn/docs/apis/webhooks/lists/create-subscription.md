# <a name="create-a-new-subscription"></a>新建订阅 

>**注意：**SharePoint webhook 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint webhook。

在 SharePoint 列表中新建一个 webhook 订阅。 

## <a name="permissions"></a>Permissions

应用程序必须至少具有对 SharePoint 列表的编辑权限，将在该列表中创建订阅。

**如果应用程序是 Microsoft Azure Active Directory (AD) 应用程序：**

必须向 Azure AD 应用授予下表中指定的权限：

应用程序 | 权限 
------------|------------
Office 365 SharePoint Online|在所有网站集中读取和写入项及列表。

**如果应用程序是 SharePoint 外接程序：**

必须向 SharePoint 外接程序授予以下权限或更高权限：

范围 | 权限 
------|------------
列表|管理

## <a name="http-request"></a>HTTP 请求

```
POST /_api/web/lists('list-id')/subscriptions
```

## <a name="request-body"></a>请求正文

请求正文中包含以下属性。

名称 | 类型 | 说明 
-----|------|------------
资源|string|从中接收通知的列表的 URL。
notificationUrl|string|向其发送通知的服务 URL。
expirationDateTime|date|通知将过期并被删除的日期。
client-clientState|string|可选。传递回所有通知上的客户端的一个不透明字符串。可以使用此对通知进行验证，或标记不同的订阅。


### <a name="example"></a>示例

```http
POST /_api/web/lists('5C77031A-9621-4DFC-BB5D-57803A94E91D')/subscriptions
Accept: application/json
Content-Type: application/json

{
  "resource": "https://contoso.sharepoint.com/_api/web/lists('5C77031A-9621-4DFC-BB5D-57803A94E91D')",
  "notificationUrl": "https://91e383a5.ngrok.io/api/webhook/handlerequest",
  "expirationDateTime": "2016-04-27T16:17:57+00:00"
}
```

## <a name="response"></a>响应

如果添加订阅，则返回一个包含新建订阅对象的 `201 Created` 响应。

### <a name="example"></a>示例

```http
HTTP/1.1 201 Created
Content-Type: application/json

{
    "id": "a8e6d5e6-9f7f-497a-b97f-8ffe8f559dc7",
    "expirationDateTime": "2016-04-27T16:17:57Z",    
    "notificationUrl": "https://91e383a5.ngrok.io/api/webhook/handlerequest",
    "resource": "5c77031a-9621-4dfc-bb5d-57803a94e91d"
}
```

## <a name="url-validation"></a>URL 验证

在创建新的订阅之前，SharePoint 将在向所提供的服务 URL 的请求正文中发送具有验证令牌的请求。服务必须通过返回验证令牌来响应此请求。

如果服务无法以此方式验证请求，将不创建订阅。有关详细信息，请参阅 [SharePoint webhook 概述](../overview-sharepoint-webhooks)
