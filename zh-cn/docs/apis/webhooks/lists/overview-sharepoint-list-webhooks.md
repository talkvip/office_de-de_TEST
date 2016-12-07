# <a name="sharepoint-list-webhooks"></a>SharePoint 列表 webhook

>**注意：**SharePoint webhook 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint webhook。

SharePoint 列表 webhook 覆盖对应于指定 SharePoint 列表或文档库的列表项更改的事件。SharePoint webhook 提供一个简单的通知管道，使应用程序得以了解对 SharePoint 列表的更改，而无需轮询服务。

## <a name="tasks"></a>任务
| 任务                                                | HTTP 方法                                                  |
|-----------------------------------------------------|--------------------------------------------------------------|
| [新建订阅](./create-subscription) | `POST    /_api/web/lists('list-guid')/subscriptions`         |
| [获取订阅](./get-subscription)          | `GET     /_api/web/lists('list-guid')/subscriptions`         |
| [删除订阅](./delete-subscription)       | `DELETE  /_api/web/lists('list-guid')/subscriptions('id')`   |
| [更新订阅](./update-subscription)     | `PATCH   /_api/web/lists('list-guid')/subscriptions('id')`   |

## <a name="list-event-types"></a>列表事件类型
通知将被发送到你的应用程序，用于 SharePoint 中的以下异步列表项事件：

* ItemAdded
* ItemUpdated
* ItemDeleted
* ItemCheckedOut
* ItemCheckedIn
* ItemUncheckedOut
* ItemAttachmentAdded
* ItemAttachmentDeleted
* ItemFileMoved
* ItemVersionDeleted
* ItemFileConverted

webhook 不支持同步事件。

## <a name="additional-resources"></a>其他资源

* [SharePoint webhook 概述](../overview-sharepoint-webhooks)
