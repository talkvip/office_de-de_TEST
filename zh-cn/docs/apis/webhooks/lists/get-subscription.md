# <a name="get-subscriptions"></a>获取订阅

>**注意：**SharePoint webhook 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint webhook。

在 SharePoint 列表中获取一个或多个 webhook 订阅。

## <a name="permissions"></a>Permissions

### <a name="get-a-single-subscription"></a>获取单个订阅

应用程序必须至少具有对 SharePoint 列表的编辑权限，将在该列表中检索订阅。

**如果应用程序是 Microsoft Azure Active Directory (AD) 应用程序：**

必须授予 Azure AD 应用程序在下表中指定的权限。只能由创建订阅的 Azure AD 应用程序来检索订阅。 

应用程序 | 权限 
------------|------------
Office 365 SharePoint Online|在所有网站集中读取和写入项及列表。

**如果应用程序是 SharePoint 外接程序：**

必须向 SharePoint 外接程序授予以下权限或更高权限。只能由创建订阅的 SharePoint 外接程序来检索订阅。 

范围 | 权限 
------|------------
列表|管理

### <a name="get-all-subscriptions"></a>获取所有订阅

应用程序必须具有对 SharePoint 列表的管理列表权限，将在该列表中检索订阅。

**如果应用程序是 Microsoft Azure Active Directory (AD) 应用程序：**

必须向 Azure AD 应用授予下表中指定的权限。 

应用程序 | 权限 
------------|------------
Office 365 SharePoint Online|具有对所有网站集的完全控制权限。

**如果应用程序是 SharePoint 外接程序：**

必须向 SharePoint 外接程序授予以下权限或更高权限。 

范围 | 权限 
------|------------
列表|完全控制

## <a name="http-request"></a>HTTP 请求

### <a name="get-a-single-subscription"></a>获取单个订阅

#### <a name="list-webhook"></a>列表 webhook
```
GET _api/web/lists('list-id')/subscriptions('id')
```

##### <a name="example"></a>示例

```http
GET _api/web/lists('5C77031A-9621-4DFC-BB5D-57803A94E91D')/subscriptions('6D77031A-2345-5GRT-BV3D-55234B56FR43')
```

#### <a name="request-body"></a>请求正文

请勿提供此方法的请求正文。

##### <a name="response"></a>响应

这将返回可由调用应用程序查看的订阅。

```http
HTTP/1.1 200 OK
Content-Type: application/json

{
  "odata.metadata": "https://contoso.sharepoint.com/_api/$metadata#SP.ApiData.Subscriptions/@Element",
  "odata.type": "Microsoft.SharePoint.Webhooks.Subscription",
  "odata.id": "https://contoso.sharepoint.com/_api/web/lists('5C77031A-9621-4DFC-BB5D-57803A94E91D')/subscriptions('a8e6d5e6-9f7f-497a-b97f-8ffe8f559dc7')",
  "odata.editLink": "web/lists('5C77031A-9621-4DFC-BB5D-57803A94E91D')/subscriptions('a8e6d5e6-9f7f-497a-b97f-8ffe8f559dc7')",
  "expirationDateTime": "2016-04-30T16:17:57Z",
  "id": "a8e6d5e6-9f7f-497a-b97f-8ffe8f559dc7",
  "notificationUrl": "https://contoso.azurewebistes.net/api/webhook/handlerequest",
  "resource": "5c77031a-9621-4dfc-bb5d-57803a94e91d"
}
```

### <a name="get-all-subscriptions"></a>获取所有订阅

#### <a name="list-webhook"></a>列表 webhook
```
GET _api/web/lists('list-id')/subscriptions
```

##### <a name="example"></a>示例

```http
GET _api/web/lists('5C77031A-9621-4DFC-BB5D-57803A94E91D')/subscriptions
```

#### <a name="request-body"></a>请求正文

请勿提供此方法的请求正文。

##### <a name="response"></a>响应

这将返回 SharePoint 资源上所有订阅的集合。 

```http
HTTP/1.1 200 OK
Content-Type: application/json

{
  "odata.metadata": "https://a830edad9050849295j16032914.sharepoint.com/_api/$metadata#SP.ApiData.Subscriptions",
  "value": [
    {
      "odata.type": "Microsoft.SharePoint.Webhooks.Subscription",
      "odata.id": "https://contoso.sharepoint.com/_api/Microsoft.SharePoint.Webhooks.Subscriptionc3175b9c-1491-454f-b5da-980431e36146",
      "odata.editLink": "Microsoft.SharePoint.Webhooks.Subscriptionc3175b9c-1491-454f-b5da-980431e36146",
      "clientState": "{A0A354EC-97D4-4D83-9DDB-144077ADB449}",
      "expirationDateTime": "2016-04-30T16:17:57Z",
      "id": "a8e6d5e6-9f7f-497a-b97f-8ffe8f559dc7",
      "notificationUrl": "https://contoso.azurewebsites.net/api/webhook/handlerequest",
      "resource": "5c77031a-9621-4dfc-bb5d-57803a94e91d"
    }
  ]
}
```
