## 一 应用程序注册
1.在 Microsoft Azure 门户中打开 Azure Active Directory 面板.
2.打开“应用注册”面板
3.在“应用注册”面板中，单击“新建注册”
4.填写必填字段并创建应用注册.

* 命名应用程序。
* 为应用程序选择“支持的帐户类型”。
* 对于“重定向 URI”
  - 选择“Web”。
  - 将 URL 设置为 https://token.botframework.com/.auth/web/redirect

* 单击“注册”
  - 创建应用以后，Azure 会显示应用的“概览”页。
  - 记录“应用程序(客户端) ID”值 。 稍后在将 Azure AD 应- 用程序注册到机器人时，需将此值用作“客户端 ID”。
  - 另请记录“目录(租户) ID”值 。 也可以用它将此应用程序注册到机器人
##  二 配置证书和机密
5.在导航窗格中单击“证书和机密”，为应用程序创建机密 。
 * 在“客户端机密”下，单击“新建客户端机密”。
 * 添加一项说明，用于将此机密与可能需要为此应用创建的其他机密（例如 bot login）区别开来。
 * 将“过期”设置为“永不”。
 * 单击“添加” 。
 * 在离开此页面之前，记录该机密。 稍后在将 Azure AD 应用程序注册到机器人时，需将此值用作“客户端机密”

6.在导航窗格中，单击“API 权限”打开“API 权限”面板 。 最佳做法是为应用显式设置 API 权限。
 * 单击“添加权限”，显示“请求 API 权限”窗格。
 * 对于此Demo，请选择“Microsoft API”和“Microsoft Graph”。
 * 选择“委托的权限”，确保选中所需权限。
## 三 在机器人中注册 Azure AD 应用程序
1.在机器人通道注册页，单击“设值”，在页面底部附近的“OAuth 连接设置”下，单击“添加设置”。
  - 对于“名称”，输入连接的名称 。 在机器人代码中会用到。
  - 对于“服务提供程序”，选择“Azure Active Directory v2”。
  - 对于“客户端 ID”，请输入为 Azure AD v1 应用程序记录的应用程序。
  - 对于“客户端机密”，请输入所创建的机密，以便为机器人授予访问 Azure AD 应用的权限。
  - 对于“作用域”，输入从应用程序注册中选择的权限的名称。

## 四 上传机器人代码
1.更新.env
  - 将 connectionName 设置为要添加到机器人的 OAuth 连接设置的名称。
  - 将 MicrosoftAppId 和 MicrosoftAppPassword 值设置为机器人的应用 ID 和应用机密。
2. 上传机器人代码到云服务平台。

## 五 配置Teams Bot
1. 利用 App Studio 插件，配置Teams Bot。
2. 安装Bot从下拉按钮中选择应用注册时的名称。
  - 应用注册时生成的AppId和App Password应该自动显示出来。
  - Bot endpoint address是https://云服务程序地址/api/messages。
  - 点击添加按钮，把插件添加到Teams。
