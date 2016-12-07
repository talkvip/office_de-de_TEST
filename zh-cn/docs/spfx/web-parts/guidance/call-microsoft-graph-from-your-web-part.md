# <a name="call-the-microsoft-graph-api-using-oauth-from-your-web-part"></a>使用 OAuth 从 Web 部件调用 Microsoft Graph API

通过与 Microsoft Graph 集成，可以向 Web 部件添加很多非常棒的功能，如电子邮件、文档和日历。若要调用 Microsoft Graph 中的 API，需要使用 JavaScript 库的 Active Directory Authentication Library (ADAL)，并使用 OAuth 流进行身份验证。需要注意一些设计注意事项和代码修改，Web 部件才能正确和安全地使用 ADAL JS 并调用 Microsoft Graph API。

## <a name="azure-active-directory-authorization-flows"></a>Azure Active Directory 授权流程

Office 365 使用 Active Directory (AD Azure) 来保护它的 API，这些 API 可通过 Microsoft Graph 进行访问。Azure AD 将 OAuth 用作授权协议。当应用程序完成 OAuth 授权流时，即可获取临时访问令牌。通过使用用户授予该应用程序的权限，该令牌代表用户提供访问特定资源的权限。

存在不同的 OAuth 流类型，具体取决于应用程序的类型。Web 应用程序使用 Azure AD 重定向至托管应用程序的 URL 的 OAuth 流。重定向是附加安全措施，用于验证该应用程序的真实性。

客户端应用程序（如 Android 和 iOS 应用）没有 URL，无法使用重定向。因此，它们无需重定向即可完成 OAuth 流程。Web 应用程序和客户端应用程序均使用公开的客户端 ID，以及仅 Azure AD 和此应用程序已知的私有客户端密码。

客户端 Web 应用程序类似于 Web 应用程序，但使用 JavaScript 实现，并在浏览器上下文中运行。如果不向用户显示，这些应用程序无法使用客户端密码。因此，这些应用程序使用名为 OAuth 隐式流的授权流访问受 Azure AD 保护的资源。在此流中，应用程序和 Azure AD 之间的合同在公开的客户端 ID 和托管应用程序的 URL 的基础上建立。这是 SharePoint Framework 客户端 Web 部件连接到受 Azure AD 保护的资源必须使用的流。若要在 Web 部件中实现基于 Azure AD 的身份验证和授权，可以使用 Microsoft 提供的[适用于 JavaScript 的 Active Directory Authentication Library (ADAL)](https://github.com/AzureAD/azure-activedirectory-library-for-js)。

## <a name="difference-between-client-side-applications-and-sharepoint-framework-web-parts"></a>客户端应用程序和 SharePoint Framework Web 部件之间的区别

典型的客户端 Web 应用程序拥有整个页面，并且是应用程序 URL 的唯一资源。页面上的所有元素在开发期间均已确定，使用此应用程序的用户无法动态地添加新元素。虽然客户端 Web 应用程序可能包含多个视图、使用外部数据，并提供一定程度的配置，但开发人员在构建应用程序时均已考虑到这些方面。

使用 SharePoint 时，用户通过添加和配置 Web 部件撰写页面。单个 Web 部件仅是整个页面的一小部分。构建 Web 部件的开发人员无法预先知道其他 Web 部件是否会放在同一页面上，而且如果是这钟情况，开发人员也无法预知这些 Web 部件将使用哪些资源。与 SharePoint 外接程序不同，Web 部件是页面的一部分。所有 Web 部件均可访问页面的 DOM，并且一个 Web 部件加载的资源，其他 Web 部件均可使用。

## <a name="limitations-when-using-adal-js-with-sharepoint-framework-client-side-web-parts"></a>使用 ADAL JS 与 SharePoint Framework 客户端 Web 部件时的限制

ADAL JS 库极大地简化了在客户端 Web 应用程序中实现  Azure AD OAuth 流的过程。但在与 SharePoint Framework 客户端 Web 部件一同使用时，则存在诸多限制。之所以存在这些限制，是因为 ADAL JS 设计初衷是由具有整个页面并且不与其他应用程序共享该页面的客户端 Web 应用程序使用。

### <a name="shared-storage"></a>共享存储

根据在项目中配置 ADAL JS 库的方式，它将在浏览器的本地存储或会话存储中存储其数据。无论哪种方式，同一页面上的任何组件均能访问 ADAL JS 存储的信息。例如，如果一个 Web 部件检索到 Microsoft Graph 访问令牌，然后同一页面上的任何其他 Web 部件均可以重复使用该令牌，并调用 Microsoft Graph。

### <a name="adal-js-as-a-singleton"></a>ADAL JS 作为单一实例

ADAL JS 旨在作为在页面的全局范围内的注册的单一实例服务工作。处理 Iframe 中的 OAuth 流回调时，ADAL JS 代码使用主窗口中已注册的单一实例引用来解析令牌流。因为每个 Web 部件使用不同的 Azure AD 应用程序注册，重复使用同一个 ADAL JS 实例会导致意外结果：每个 Web 部件将使用与该页面加载的第一个 Web 部件所注册的配置相同的配置。

### <a name="resource-based-storage"></a>基于资源的存储

默认情况下，ADAL JS 库存储信息，例如当前访问令牌或与特定资源（如 Microsoft Graph 或 SharePoint）关联的一组密钥的过期时间。如果一个页面上有多个组件访问具有不同权限的同一资源，则首个检索到的令牌将由 ADAL JS 提供给所有组件。如果不同的组件需要访问同一资源，它们可能会失败。若要阐明此限制，请考虑如下示例。

假设页面上有两个 Web 部件：一个列出了即将召开的会议，另外一个创建了新会议。若要获取即将召开的会议，即将召开会议的 Web 部件将使用具有读取用户日历的权限的 Microsoft Graph 访问令牌。另一方面，创建新会议的 Web 部件需要具有写入到用户日历的权限的 Microsoft Graph 访问令牌。如果首先加载即将召开会议的 Web 部件，则 ADAL JS 检索到仅具有读取权限的令牌。然后，同一令牌由 ADAL JS 提供给创建新会议的 Web 部件。如你所见，Web 部件尝试创建新会议时，Web 部件将由于访问被拒绝错误而失败。由于采用基于资源的存储，所以在上一个检索到的令牌过期前，ADAL JS 无法检索 Microsoft Graph 的新令牌。

### <a name="processing-authorization-callbacks"></a>处理授权回调

OAuth 隐式流基于 Azure AD 和参与此流的应用程序之间的重定向。身份验证过程启动时，应用程序将你重定向到 Azure AD 登录页。完成身份验证后，Azure AD 将你重定向到该应用程序，并在 URL 哈希中发送标识令牌，以便应用程序进行处理。从置于 SharePoint 页面上的 Web 部件的角度来看，这种行为有两个缺点。

虽然已登录到 SharePoint，但你需要通过 Azure AD 的单独身份验证，Web 部件才能访问受 Azure AD 保护的资源。Web 部件是整页的一小部分，但为完成身份验证过程，需要离开整个页面，这是一种较差的用户体验。

身份验证完成后，Azure AD 会将你重定向到应用程序。URL 哈希包括应用程序需要处理的标识令牌。遗憾的是，因为标识令牌不包含有关其来源的任何信息，因此，如果同一页面上的存在多个 Web 部件，则全部 Web 部件均将尝试从 URL 处理标识令牌。

## <a name="use-adal-js-with-sharepoint-framework-client-side-web-parts"></a>使用 ADAL JS与 SharePoint Framework 客户端 Web 部件

可以克服 ADAL JS 中的限制，在 SharePoint Framework 客户端 Web 部件中成功实现 OAuth。

### <a name="load-adal-js-in-your-web-part"></a>在 Web 部件中加载 ADAL JS

Web 部件使用 TypeScript 生成，作为 AMD 模块进行分发。其模块化的体系结构可以隔离 Web 部件使用的不同资源。ADAL JS 设计为在全局范围内进行注册，但可以通过使用以下代码在 Web 部件中加载它：

```ts
const AuthenticationContext = require('adal-angular');
```

此语句导入可公开 ADAL JS 功能的 `AuthenticationContext` 类，以便在 Web 部件中使用。尽管模块名称是这样，但 `AuthenticationContext` 类可与每个 JavaScript 库协同工作。

### <a name="make-adal-js-suitable-for-use-with-sharepoint-framework-web-parts"></a>使 ADAL JS 适合与 SharePoint Framework Web 部件一起使用

当前 ADAL JS 功能不适合在 Web 部件中使用。使用以下修补程序来使 ADAL JS 在 Web 部件中工作。

**WebPartAuthenticationContext.js:**
```js
const AuthenticationContext = require('adal-angular');

AuthenticationContext.prototype._getItemSuper = AuthenticationContext.prototype._getItem;
AuthenticationContext.prototype._saveItemSuper = AuthenticationContext.prototype._saveItem;
AuthenticationContext.prototype.handleWindowCallbackSuper = AuthenticationContext.prototype.handleWindowCallback;
AuthenticationContext.prototype._renewTokenSuper = AuthenticationContext.prototype._renewToken;
AuthenticationContext.prototype.getRequestInfoSuper = AuthenticationContext.prototype.getRequestInfo;

AuthenticationContext.prototype._getItem = function (key) {
  if (this.config.webPartId) {
    key = this.config.webPartId + '_' + key;
  }

  return this._getItemSuper(key);
};

AuthenticationContext.prototype._saveItem = function (key, object) {
  if (this.config.webPartId) {
    key = this.config.webPartId + '_' + key;
  }

  return this._saveItemSuper(key, object);
};

AuthenticationContext.prototype.handleWindowCallback = function (hash) {
  if (hash == null) {
    hash = window.location.hash;
  }

  if (!this.isCallback(hash)) {
    return;
  }

  var requestInfo = this.getRequestInfo(hash);
  if (requestInfo.requestType === this.REQUEST_TYPE.LOGIN) {
    return this.handleWindowCallbackSuper(hash);
  }

  var resource = this._getResourceFromState(requestInfo.stateResponse);
  if (!resource || resource.length === 0) {
    return;
  }

  if (this._getItem(this.CONSTANTS.STORAGE.RENEW_STATUS + resource) === this.CONSTANTS.TOKEN_RENEW_STATUS_IN_PROGRESS) {
    return this.handleWindowCallbackSuper(hash);
  }
}

AuthenticationContext.prototype._renewToken = function (resource, callback) {
  this._renewTokenSuper(resource, callback);
  var _renewStates = this._getItem('renewStates');
  if (_renewStates) {
    _renewStates = _renewStates.split(';');
  }
  else {
    _renewStates = [];
  }
  _renewStates.push(this.config.state);
  this._saveItem('renewStates', _renewStates);
}

AuthenticationContext.prototype.getRequestInfo = function (hash) {
  var requestInfo = this.getRequestInfoSuper(hash);
  var _renewStates = this._getItem('renewStates');
  if (!_renewStates) {
    return requestInfo;
  }

  _renewStates = _renewStates.split(';');
  for (var i = 0; i < _renewStates.length; i++) {
    if (_renewStates[i] === requestInfo.stateResponse) {
      requestInfo.requestType = this.REQUEST_TYPE.RENEW_TOKEN;
      requestInfo.stateMatch = true;
      break;
    }
  }

  return requestInfo;
}

window.AuthenticationContext = function() {
  return undefined;
}
```

此修补程序将以下更改应用到 ADAL JS：
- 所有信息均存储在 Web 部件特有的密钥中（请参阅 `_getItem` 和 `_saveItem` 函数替代）
- 仅启动了回调的 Web 部件可以处理它们（请参阅 `handleWindowCallback` 函数替代）
- 从回调验证数据时，使用特定 Web 部件 `AuthenticationContext` 类的实例，而不是全局注册的单一实例（请参阅 `_renewToken`、`getRequestInfo` 和 `window.AuthenticationContext` 函数的空注册）

### <a name="use-the-adal-js-patch-in-sharepoint-framework-web-parts"></a>使用 SharePoint Framework Web 部件中的 ADAL JS 修补程序

要使 ADAL JS 在 SharePoint Framework Web 部件中正常工作，必须以特定方式对其进行配置。要做到这一点，首先需要定义可扩展标准 ADAL JS `Config` 接口的自定义接口，以便公开配置对象上的其他属性。

```ts
export interface IAdalConfig extends adal.Config {
  popUp?: boolean;
  callback?: (error: any, token: string) => void;
  webPartId?: string;
}
```

`popUp` 属性和 `callback` 函数均已在 ADAL JS 中实现，但尚未在标准配置接口的 TypeScript 键入中公开。你需要上述属性和函数，这样用户使用弹出窗口而不是重定向到 Azure AD 登录页面即可登录到 Web 部件。

下一步将把 ADAL JS 修补程序、自定义配置接口和配置对象加载到 Web 部件的主要组件。

```tsx
const AuthenticationContext = require('adal-angular');
import adalConfig from '../AdalConfig';
import { IAdalConfig } from '../../IAdalConfig';
import '../../WebPartAuthenticationContext';
```

在组件的构造函数中，使用其他属性扩展 ADAL JS 配置，并将其用于创建 `AuthenticationContext` 类的新实例。

```tsx
export default class UpcomingMeetings extends React.Component<IUpcomingMeetingsProps, IUpcomingMeetingsState> {
  private authCtx: adal.AuthenticationContext;

  constructor(props: IUpcomingMeetingsProps, state: IUpcomingMeetingsState) {
    super(props);

    this.state = {
      loading: false,
      error: null,
      upcomingMeetings: [],
      signedIn: false
    };

    const config: IAdalConfig = adalConfig;
    config.webPartId = this.props.webPartId;
    config.popUp = true;
    config.callback = (error: any, token: string): void => {
      this.setState((previousState: IUpcomingMeetingsState, currentProps: IUpcomingMeetingsProps): IUpcomingMeetingsState => {
        previousState.error = error;
        previousState.signedIn = !(!this.authCtx.getCachedUser());
        return previousState;
      });
    };

    this.authCtx = new AuthenticationContext(config);
    AuthenticationContext.prototype._singletonInstance = undefined;
  }
}
```

代码首先检索标准 ADAL JS 配置对象，然后将其转换为新定义的配置接口的类型。然后，它将传递 Web 部件实例的 ID，这样所有 ADAL JS 值就能以不与该页面上其他 Web 部件冲突的方式进行存储。接下来，它将启用基于弹出窗口的身份验证，并定义身份验证完成时运行的回调函数来更新组件。最后，它将创建 `AuthenticationContext` 类的新实例，并将重置 `AuthenticationContext` 类的构造函数中设置的 `_singletonInstance`。这是允许每个 Web 部件使用其自己的 ADAL JS 身份验证上下文版本所必需的。

## <a name="considerations-when-using-oauth-implicit-flow-in-sharepoint-framework-client-side-web-parts"></a>在 SharePoint Framework 客户端 Web 部件中使用 OAuth 隐式流时的注意事项

在生成与受 Azure AD 保护的资源进行通信的 SharePoint Framework 客户端 Web 部件之前，有一些应考虑的约束条件。以下信息将帮助你确定 Web 部件中的 OAuth 授权成为解决方案和组织的正确选择的最佳时机。

### <a name="sharepoint-framework-web-parts-are-highly-trusted"></a>SharePoint Framework Web 部件可高度信任

与 SharePoint 外接程序不同，Web 部件使用与当前用户相同的权限运行。用户可以执行的操作，该 Web 部件均可执行。这就是只有租户管理员可以安装和部署 Web 部件的原因。租户管理员应验证所部署的 Web 部件来自受信任的源，并已批准在租户中使用。

Web 部件是页面的一部分，并且与 SharePoint 外接程序不同，可与页面上的其他元素共享 DOM 和资源。授予由 Azure AD 保护的资源访问权限的访问令牌通过对 Web 部件所在的同一页面回调进行检索。该回调可以由此页面上的任何元素进行处理。此外，一旦从回调处理访问令牌后，它们存储在浏览器的本地存储或会话存储中，页面上任何组件均可从中进行检索。恶意 Web 部件可以读取令牌，并公开该令牌或使用该令牌将它所检索的数据公开到外部服务。

### <a name="url-is-a-part-of-the-oauth-implicit-flow-contract"></a>URL 是 OAuth 隐式流合同的一部分

如果不向用户显示，客户端应用程序无法存储客户端密码。为了验证特定应用程序的真实性，托管应用程序的 URL 用作 Azure AD 和该应用程序之间的信任协定的一部分。这样做的后果是每个具有使用 OAuth 隐式流的 Web 部件的页面，它的 URL 必须注册 Azure AD。如果未注册页面，OAuth 流将失败，该 Web 部件将被拒绝访问受保护的资源。

对于允许用户将 Web 部件添加到页面的组织来说，此项必需的安全措施是一个重大限制。我们建议限制使用 OAuth 隐式流的 Web 部件的位置数量。使用几个众所周知的位置（如主页或特定的登录页），并在 Azure AD 中注册这些位置。

### <a name="internet-explorer-security-zones"></a>Internet Explorer 安全区域

Internet Explorer 使用安全区域，将安全策略应用到所访问的网站。分配给不同安全区域的网站独立运行，并且无法相互协作。当 SharePoint Framework 客户端 Web 部件中使用 OAuth 隐式流时，具有使用 OAuth 的 Web 部件的页面和 Azure AD 登录页（位于 https://login.microsoftonline.com）必须位于同一安全区域中。没有此类配置，身份验证流程将失败，Web 部件将无法与受 Azure AD 保护的资源进行通信。

### <a name="user-must-sign-in-frequently"></a>用户必须经常登录

使用常规 OAuth 流 Web 应用程序和客户端应用程序获取短期访问令牌和刷新令牌（有效期更长）时。如果访问令牌过期，则刷新令牌可用于获取新的访问令牌。这种方法降低了用户需要登录到 Azure AD 的频率。

客户端应用程序无法安全存储密钥，而泄漏刷新令牌是一个严重的安全威胁，因为它们仅收到作为 OAuth 隐式流一部分的访问令牌。一旦该令牌过期，应用程序必须启动新的 OAuth 流，这就要求用户重新登录。

## <a name="more-information"></a>详细信息

- [Azure AD 的身份验证场景](https://Azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-scenarios/)