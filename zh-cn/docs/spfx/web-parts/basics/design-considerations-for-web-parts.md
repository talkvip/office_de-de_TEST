# <a name="design-considerations-for-sharepoint-client-side-web-parts"></a>SharePoint 客户端 Web 部件的设计注意事项

>**注意：**SharePoint Framework 目前处于预览模式，并可能会发生更改。当前不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。
>

要开始设计 Web 部件，你需要熟悉 [Office UI Fabric](http://dev.office.com/fabric)。来自 [Fabric 核心](https://github.com/OfficeDev/office-ui-fabric-core)的所有样式 — 包括图标、版式、颜色用法、动画和响应网格 – 已默认加载且可用于 Web 部件。因为这可能会与全局副本冲突，所以请勿导入 Web 部件的 Fabric 副本。这些类提供 Web 部件样式的基础，如果需要不同的视觉效果以匹配公司品牌，你可以始终离开。

## <a name="office-ui-fabric-react-components"></a>Office UI Fabric React 组件

结合 Office UI Fabric，你可以使用 Office UI Fabric React 组件来构建 Web 部件。Fabric React 是一个快速响应、移动优先的组件集合，旨在使用 Office 设计语言来快速简便地创建 Web 体验。

下面待办事项列表示例在属性窗格中使用 Fabric 组件以使页面作者可配置 Web 部件。

![使用 Fabric 的 Todo Web 部件示例](../../../../images/design-wp-todo-example.png)

你可以在“[Office UI Fabric 样式](http://dev.office.com/fabric/styles)”中查找 Office UI Fabric 样式、版式、颜色、图标和动画的完整列表。


## <a name="responsive-behavior"></a>响应行为

新 SharePoint 创作体验中的页面使用 Office UI Fabric 响应网格来帮助确保每个页面的效果。 

### <a name="max-width"></a>最大宽度

我们建议所有 Web 部件都使用 100% 最大宽度，以确保他们会重排并在任何页面上正常运行。页面和列的宽度由页面模板定义，但可以通过作者修改。如果在 Web 部件中设置了最大像素值，当看到页面具有不同宽度时，可能会导致功能和布局出现意外结果。

![最大宽度 Web 部件响应行为](../../../../images/design-wp-responsive-max-width.png)

### <a name="min-width"></a>最小宽度

页/列宽度减小到最小宽度 320 像素时，所有 Web 部件应可进行重排。

![最小宽度 Web 部件响应行为](../../../../images/design-wp-responsive-min-width.png)

## <a name="web-part-published-mode-vs-edit-mode"></a>Web 部件发布的模式与编辑模式对比

新的 SharePoint 页面创作体验有两种模式：

* **发布的模式**允许你的团队或受众查看内容并与 Web 部件交互。
* **编辑模式**允许页面作者添加和配置 Web 部件以将内容添加到页面。

### <a name="edit-mode"></a>编辑模式

#### <a name="add-hint-and-toolbox"></a>添加提示和工具箱

添加提示是具有加号图标的横线，在选择 Web 部件并悬停鼠标时可见以指示页面作者可在哪里将新 Web 部件添加到相应的页面。当用户单击/点按加号图标时，将打开工具箱。工具箱包含可添加到页面中的所有 Web 部件。

![Web 部件添加提示和工具箱](../../../../images/design-wp-toolbox.png)

#### <a name="toolbar"></a>工具栏

垂直工具栏和边界框是每个 Web 部件框架的一部分，并且由页面提供。每个 Web 部件的工具栏中都具有编辑和删除操作。

![Web 部件垂直工具栏](../../../../images/design-wp-toolbar.png)

#### <a name="contextual-edits"></a>上下文编辑

应为 Web 部件设计 WYSIWYG/所见即所得体验，以填写信息或添加发布时向用户显示的内容。应在页面中输入此内容，以便用户了解查看者将如何看到内容。例如，应在显示文本的位置填写标题和描述或者应在页面的上下文中添加和修改新任务。

![Web 部件上下文编辑](../../../../images/design-wp-contexual-edits.png)

#### <a name="item-level-edits"></a>项目级编辑

UI 会在 Web 部件中更改；例如，将文本插入文本字段以填写链接时、显示 UI 以对项目重新排序或检查 Web 部件中的任务时

![Web 部件项目级编辑](../../../../images/design-wp-item-level-edits.png)

## <a name="property-panes"></a>属性窗格

通过工具栏上的编辑操作图标调用属性窗格。窗格应主要包含配置设置，启用/禁用功能以显示在页面上或调用服务来显示内容。

![Web 部件属性窗格设计](../../../../images/design-wp-pp.png)

有三种类型的属性窗格使你可以设计和开发 Web 部件，使其符合你的业务或客户需求。

### <a name="single-pane"></a>单个窗格

单个窗格用于简单的 Web 部件（仅有少量的属性要配置）。

![Web 部件单个属性窗格设计](../../../../images/design-wp-pp-single.png)

### <a name="accordion-pane"></a>可折叠项窗格

可折叠项窗格用于包含一组或多组具有许多选项的属性，其中组将生成较长的选项滚动列表。例如，可能有三个组（名为“属性”、“外观”和“布局”），每个组都有十个组件。

#### <a name="accordion---one-group-open"></a>可折叠项 - 一个组打开

![Web 部件可折叠项属性窗格有一个组打开](../../../../images/design-wp-pp-accordion.png)

#### <a name="accordion--two-groups-open-and-scrolled"></a>可折叠项 - 两个组打开并滚动

![Web 部件可折叠项属性窗格有两个组打开并滚动](../../../../images/design-wp-pp-accordion-groups.png)


### <a name="property-pane-stepspages"></a>属性窗格步骤/页面

当需要按线性顺序配置 Web 部件或对第二个步骤中显示的第一个步骤影响选项做出选择时，步骤窗格用于对多个步骤或页面中的属性进行分组。

**步骤 1/3**
![属性窗格，具有步骤设计 1/3](../../../../images/design-wp-pp-pages-step1.png)

**步骤 2/3**
![属性窗格，具有步骤设计 2/3](../../../../images/design-wp-pp-pages-step2.png)

**步骤 3/3**
![属性窗格，具有步骤设计 3/3](../../../../images/design-wp-pp-pages-step3.png)


## <a name="reactive-vs-non-reactive-web-parts"></a>反应与非反应 Web 部件

反应 Web 部件旨在作为完整的客户端 Web 部件，即当页面上的 Web 部件进行更改时属性窗格中配置的每个组件均会反映出来。对于待办事项列表 Web 部件，取消选中“已完成任务”将隐藏 Web 部件中的此视图。

![反应 Web 部件设计](../../../../images/design-wp-pp-reactive.png)

非反应 Web 部件完全不是客户端且通常需要调用一个或多个属性以设置/拉取或将数据存储在服务器上。在这种情况下，你应该启用属性窗格底部的“应用”和“取消”按钮。

![非反应 Web 部件设计](../../../../images/design-wp-pp-non-reactive.png)

## <a name="constructing-the-to-do-list-property-pane"></a>构造待办事项列表属性窗格

待办事项列表示例使用单个窗格且是反应 Web 部件。下面显示每个 Fabric React 组件和最终设计。

添加待办事项列表的描述![添加待办事项列表的描述](../../../../images/design-wp-todo-pp-description.png)

下拉列表 – 用于从现有列表中选择任务![下拉列表用于从现有列表中选择任务](../../../../images/design-wp-todo-pp-dropdown.png)

复选框 – 允许作者显示或隐藏不同的视图![复选框允许作者显示或隐藏不同的视图](../../../../images/design-wp-todo-pp-checkbox.png)

滑块 – 用于设置可见的任务数![滑块用于设置可见的任务数](../../../../images/design-wp-todo-pp-slider.png)

从下拉列表中选择列表后，Web 部件显示项目加载到页面上的指示器![Web 部件显示加载项目的指示器](../../../../images/design-wp-todo-loading-indicator.png)

当加载新的任务时，视图淡出使用来自 Office UI Fabric 的动画样式![Web 部件与 Fabric 动画](../../../../images/design-wp-todo-fabric-animations.png)
