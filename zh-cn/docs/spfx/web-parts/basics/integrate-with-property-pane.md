# <a name="integrate-your-sharepoint-client-side-web-part-with-the-property-pane"></a>将 SharePoint 客户端 Web 部件与属性窗格相集成。

>**注意：**SharePoint Framework 当前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

属性窗格允许最终用户使用一组属性来配置 Web 部件。文章[构建你的第一个 Web 部件](../get-started/build-a-hello-world-web-part)描述了属性窗格在 **HelloWorldWebPart** 类中的定义方式。属性窗格属性在  **propertyPaneSettings** 中定义。

下图显示了 SharePoint 中的属性窗格的一个示例。

![属性窗格示例](../../../../images/property-pane-example.png)

属性窗格中有三个重要的元数据：

* 页面
* 标头
* 组

页面使你可以灵活地分隔复杂的交互，并将其放入一个或多个页面。然后，页面包含标头和组。

标头使你可以定义属性窗格中的标题，组使你可以定义属性窗格的各个部分，通过此可以对你的字段集进行分组。 

属性窗格应包含一个页面、一个可选标头和至少一个组。

然后在组内定义属性字段。 

## <a name="using-the-property-pane"></a>使用属性窗格

以下代码示例初始化并配置 Web 部件中的属性窗格。创建类型 **IPropertyPaneSettings** 的方法，并返回属性窗格页的集合。

```ts
protected get propertyPaneSettings(): IPropertyPaneSettings {
  return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
              })
            ]
          }
        ]
      }
    ]
  };
}
```

### <a name="property-pane-fields"></a>属性窗格字段

支持以下字段类型：

* 标签
* 文本框
* 多行文本框
* 复选框
* 下拉列表
* 链接
* 滑块
* 切换
* 自定义

字段类型可作为 **sp-client-platform** 中的模块使用。需要对其导入才能在代码中使用：

```ts
import {
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneLink,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-client-preview';
```

每个字段类型构造函数的定义如下所示，以 **PropertyPaneTextField** 为例：

```ts
PropertyPaneTextField('targetProperty',{
  //field properties are defined here
})
```

**targetProperty** 定义该字段类型的关联对象，同时在 Web 部件的属性界面中定义。

若要将类型分配给这些属性，请在包含一个或多个目标属性的 Web 部件类中定义一个接口。

```ts
export interface IHelloWorldWebPartProps {
    targetProperty: string
}
```

然后可以在 Web 部件中使用 **this.properties.targetProperty**。

```ts
<p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
```

定义属性后，可以在 Web 部件中使用 **this.properties.<property-value>** 访问它们：有关详细信息，请参阅 [**HelloWorldWebPart**的 **render** 方法 ](../get-started/build-a-hello-world-web-part#web-part-render-method)：

## <a name="handling-field-changes"></a>处理字段更改

属性窗格具有两种交互模式：

* 反应
* 非反应

在反应模式中，每次更改时都将触发更改事件。反应行为会自动使用新值更新 Web 部件。

虽然反应模式满足大多数情况，但有时需要非反应行为。非反应不会自动更新 Web 部件，除非用户确认所做的更改。

## <a name="custom-field-example"></a>自定义字段示例

在 **groupFields** 数组中添加以下字段定义。

```ts
{
  type: IPropertyPaneFieldType.Custom,
  targetProperty: 'custom',
  properties: {
    onRender: this._customFieldRender.bind(this),
    value: undefined,
    context: undefined
  }
}
```

将以下类型添加到 **@microsoft/sp-client-platform** 导入。

```ts
IPropertyPaneFieldType,
IOnCustomPropertyFieldChanged
```

添加以下专用方法以呈现自定义字段。

```ts
private _customFieldRender(elem: HTMLElement, context: any, onChanged?: IOnCustomPropertyFieldChanged): void {
    elem.innerHTML = '<input id="password" type="password" name="password" class="ms-TextField-field">';
}
```
