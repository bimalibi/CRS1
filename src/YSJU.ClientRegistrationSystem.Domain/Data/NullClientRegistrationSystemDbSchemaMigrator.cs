using System.Threading.Tasks;
using Volo.Abp.DependencyInjection;

namespace YSJU.ClientRegistrationSystem.Data;

/* This is used if database provider does't define
 * IClientRegistrationSystemDbSchemaMigrator implementation.
 */
public class NullClientRegistrationSystemDbSchemaMigrator : IClientRegistrationSystemDbSchemaMigrator, ITransientDependency
{
    public Task MigrateAsync()
    {
        return Task.CompletedTask;
    }
}
