# <a name="working-with-the-original-requestdigest"></a>使用原始 __RequestDigest

>**注意：**SharePoint Framework 当前处于预览模式，并可能会发生更改。不支持在生产环境中使用 SharePoint Framework 客户端 Web 部件。

有大量编写以使用经典 SharePoint 页面的代码，可以将其与 SharePoint Framework 结合使用，但有时其中不包含某些组件或变量。__REQUESTDIGEST 表单域是一个示例。在理想世界中，你不会使用全局变量来访问摘要，而仅使用更新的 **HttpRequest** 对象来发出 SharePoint 调用，并且它将为你处理所有摘要/身份验证逻辑（包括过期令牌）。[将客户端 Web 部件连接到 SharePoint（Hello world 第 2 部分）](https://github.com/SharePoint/sp-dev-docs/wiki/HelloWorld,-Talking-to-SharePoint)文章介绍了如何执行此操作。

但是，如果现有的代码使用某些较旧的结构，通过强大的客户端代码和 DOM 操作，可以很容易地向页面添加回这些内容。关键是要挂钩到基 Web 部件类中的 **onInit** 方法，并拉取要在那里创建的 DOM 元素。下面是创建 __REQESTDIGEST 表单元素的一个示例。

```JavaScript
    public onInit<T>(): Promise<T>
    {
    // does the digest exist?
    if ( !document.getElementById('__REQUESTDIGEST') )
    {
      // OK, the request digest does not exist.  Let's create it.
      // first, grab the digest value out of the contextWebInfo object (if it exists).
      var digestValue: string;
      try{
        digestValue = (window as any)._spClientSidePageContext.contextWebInfo.FormDigestValue;
      }
      catch (exception){
        // there is no digest on this page, so just return.  This can easily happen on the local workbench
        return Promise.resolve();
      }

      if (digestValue){
        // OK, now lets create the digest input form.  It looks like this -
        // <input type="hidden" name="__REQUESTDIGEST" id="__REQUESTDIGEST" value="blahblahblahblahblahblah, July23 -0000 or something like that">
        const requestDigestInput: Element = document.createElement('input');
        requestDigestInput.setAttribute('type', 'hidden');
        requestDigestInput.setAttribute('name', '__REQUESTDIGEST');
        requestDigestInput.setAttribute('id', '__REQUESTDIGEST');
        requestDigestInput.setAttribute('value', digestValue);

        // lastly, add the digest to the page
        document.body.appendChild(requestDigestInput);
      }
    }

    // no promise to return
    return Promise.resolve();
    }
```

>**注意：**有更好的方法来获取将处理所有缓存/过期/正在重取/等的当前摘要值。试一试。将需要从 sp-client-base 导入 digestCacheServiceKey 和 IDigestCache

```JavaScript
    var digestCache:IDigestCache = this.context.serviceScope.consume(digestCacheServiceKey);
    digestCache.fetchDigest(this.context.pageContext.web.serverRelativeUrl).then((digest: string) => {
      // Do Something with the digest
      console.log(digest);
    });
```