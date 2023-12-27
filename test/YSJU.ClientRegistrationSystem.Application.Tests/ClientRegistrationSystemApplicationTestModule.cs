using Volo.Abp.Modularity;

namespace YSJU.ClientRegistrationSystem;

[DependsOn(
    typeof(ClientRegistrationSystemApplicationModule),
    typeof(ClientRegistrationSystemDomainTestModule)
    )]
public class ClientRegistrationSystemApplicationTestModule : AbpModule
{

}
