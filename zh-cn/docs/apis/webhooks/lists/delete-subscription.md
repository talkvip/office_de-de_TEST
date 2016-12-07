# <a name="delete-a-subscription"></a>删除订阅

>**注意：**SharePoint webhook 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint webhook。

删除 SharePoint 列表中的 webhook 订阅。删除订阅后将不再传送通知。

## <a name="permissions"></a>Permissions

应用程序必须至少具有对 SharePoint 列表的编辑权限，将在该列表中删除订阅。

**如果应用程序是 Microsoft Azure Active Directory (AD) 应用程序：**

必须向 Azure AD 应用授予下表中指定的权限。只能由创建订阅的 Azure AD 应用程序来删除订阅。

应用程序 | 权限 
------------|------------
Office 365 SharePoint Online|在所有网站集中读取和写入项及列表。

**如果应用程序是 SharePoint 外接程序：**

必须向 SharePoint 外接程序授予以下权限或更高权限。只能由创建订阅的 SharePoint 外接程序来删除订阅。

范围 | 权限 
------|------------
列表|管理

## <a name="http-request"></a>HTTP 请求

```
DELETE _api/web/lists('list-id')/subscriptions('id')
```

### <a name="example"></a>示例

```http
DELETE _api/web/lists('5C77031A-9621-4DFC-BB5D-57803A94E91D')/subscriptions('6D77031A-2345-5GRT-BV3D-55234B56FR43')
```

## <a name="request-body"></a>请求正文

请勿提供此方法的请求正文。

## <a name="response"></a>响应

如果找到并成功删除订阅，则返回 `204 No Content` 响应。

### <a name="example"></a>示例

```http
HTTP/1.1 204 No Content
```
