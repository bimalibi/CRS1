using YSJU.ClientRegistrationSystem.EntityFrameworkCore;
using Volo.Abp.Autofac;
using Volo.Abp.Caching;
using Volo.Abp.Caching.StackExchangeRedis;
using Volo.Abp.Modularity;

namespace YSJU.ClientRegistrationSystem.DbMigrator;

[DependsOn(
    typeof(AbpAutofacModule),
    typeof(AbpCachingStackExchangeRedisModule),
    typeof(ClientRegistrationSystemEntityFrameworkCoreModule),
    typeof(ClientRegistrationSystemApplicationContractsModule)
    )]
public class ClientRegistrationSystemDbMigratorModule : AbpModule
{
    public override void ConfigureServices(ServiceConfigurationContext context)
    {
        Configure<AbpDistributedCacheOptions>(options => { options.KeyPrefix = "ClientRegistrationSystem:"; });
    }
}
