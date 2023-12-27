using Volo.Abp.Ui.Branding;
using Volo.Abp.DependencyInjection;

namespace YSJU.ClientRegistrationSystem;

[Dependency(ReplaceServices = true)]
public class ClientRegistrationSystemBrandingProvider : DefaultBrandingProvider
{
    public override string AppName => "ClientRegistrationSystem";
}
