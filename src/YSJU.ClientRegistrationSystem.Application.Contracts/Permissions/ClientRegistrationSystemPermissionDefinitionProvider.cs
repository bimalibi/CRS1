using YSJU.ClientRegistrationSystem.Localization;
using Volo.Abp.Authorization.Permissions;
using Volo.Abp.Localization;

namespace YSJU.ClientRegistrationSystem.Permissions;

public class ClientRegistrationSystemPermissionDefinitionProvider : PermissionDefinitionProvider
{
    public override void Define(IPermissionDefinitionContext context)
    {
        var myGroup = context.AddGroup(ClientRegistrationSystemPermissions.GroupName);
        //Define your own permissions here. Example:
        //myGroup.AddPermission(ClientRegistrationSystemPermissions.MyPermission1, L("Permission:MyPermission1"));
    }

    private static LocalizableString L(string name)
    {
        return LocalizableString.Create<ClientRegistrationSystemResource>(name);
    }
}
