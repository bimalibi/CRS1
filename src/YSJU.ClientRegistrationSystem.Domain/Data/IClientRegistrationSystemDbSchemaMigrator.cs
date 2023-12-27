using System.Threading.Tasks;

namespace YSJU.ClientRegistrationSystem.Data;

public interface IClientRegistrationSystemDbSchemaMigrator
{
    Task MigrateAsync();
}
