using System;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using YSJU.ClientRegistrationSystem.Data;
using Volo.Abp.DependencyInjection;

namespace YSJU.ClientRegistrationSystem.EntityFrameworkCore;

public class EntityFrameworkCoreClientRegistrationSystemDbSchemaMigrator
    : IClientRegistrationSystemDbSchemaMigrator, ITransientDependency
{
    private readonly IServiceProvider _serviceProvider;

    public EntityFrameworkCoreClientRegistrationSystemDbSchemaMigrator(
        IServiceProvider serviceProvider)
    {
        _serviceProvider = serviceProvider;
    }

    public async Task MigrateAsync()
    {
        /* We intentionally resolve the ClientRegistrationSystemDbContext
         * from IServiceProvider (instead of directly injecting it)
         * to properly get the connection string of the current tenant in the
         * current scope.
         */

        await _serviceProvider
            .GetRequiredService<ClientRegistrationSystemDbContext>()
            .Database
            .MigrateAsync();
    }
}
