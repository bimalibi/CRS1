using YSJU.ClientRegistrationSystem.EntityFrameworkCore;
using Volo.Abp.Modularity;

namespace YSJU.ClientRegistrationSystem;

[DependsOn(
    typeof(ClientRegistrationSystemEntityFrameworkCoreTestModule)
    )]
public class ClientRegistrationSystemDomainTestModule : AbpModule
{

}
