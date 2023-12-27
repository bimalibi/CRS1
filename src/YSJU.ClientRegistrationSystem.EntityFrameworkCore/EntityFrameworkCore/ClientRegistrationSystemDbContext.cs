using Microsoft.EntityFrameworkCore;
using System.Linq;
using Volo.Abp.AuditLogging.EntityFrameworkCore;
using Volo.Abp.BackgroundJobs.EntityFrameworkCore;
using Volo.Abp.Data;
using Volo.Abp.DependencyInjection;
using Volo.Abp.EntityFrameworkCore;
using Volo.Abp.EntityFrameworkCore.Modeling;
using Volo.Abp.FeatureManagement.EntityFrameworkCore;
using Volo.Abp.Identity;
using Volo.Abp.Identity.EntityFrameworkCore;
using Volo.Abp.OpenIddict.EntityFrameworkCore;
using Volo.Abp.PermissionManagement.EntityFrameworkCore;
using Volo.Abp.SettingManagement.EntityFrameworkCore;
using Volo.Abp.TenantManagement;
using Volo.Abp.TenantManagement.EntityFrameworkCore;
using YSJU.ClientRegistrationSystem.AppEntities.ClientDetails;
using YSJU.ClientRegistrationSystem.AppEntities.ClientProductCategories;
using YSJU.ClientRegistrationSystem.AppEntities.ProductCategories;

namespace YSJU.ClientRegistrationSystem.EntityFrameworkCore;

[ReplaceDbContext(typeof(IIdentityDbContext))]
[ReplaceDbContext(typeof(ITenantManagementDbContext))]
[ConnectionStringName("Default")]
public class ClientRegistrationSystemDbContext :
    AbpDbContext<ClientRegistrationSystemDbContext>,
    IIdentityDbContext,
    ITenantManagementDbContext
{
    /* Add DbSet properties for your Aggregate Roots / Entities here. */

    #region Entities from the modules

    /* Notice: We only implemented IIdentityDbContext and ITenantManagementDbContext
     * and replaced them for this DbContext. This allows you to perform JOIN
     * queries for the entities of these modules over the repositories easily. You
     * typically don't need that for other modules. But, if you need, you can
     * implement the DbContext interface of the needed module and use ReplaceDbContext
     * attribute just like IIdentityDbContext and ITenantManagementDbContext.
     *
     * More info: Replacing a DbContext of a module ensures that the related module
     * uses this DbContext on runtime. Otherwise, it will use its own DbContext class.
     */

    //Identity
    public DbSet<IdentityUser> Users { get; set; }
    public DbSet<IdentityRole> Roles { get; set; }
    public DbSet<IdentityClaimType> ClaimTypes { get; set; }
    public DbSet<OrganizationUnit> OrganizationUnits { get; set; }
    public DbSet<IdentitySecurityLog> SecurityLogs { get; set; }
    public DbSet<IdentityLinkUser> LinkUsers { get; set; }
    public DbSet<IdentityUserDelegation> UserDelegations { get; set; }
    public DbSet<ProductCategory> ProductCategories { get; set; }
    public DbSet<ClientDetail> ClientDetails { get; set; }
    public DbSet<ClientProductCategory> ClientProductCategories { get; set; }

    // Tenant Management
    public DbSet<Tenant> Tenants { get; set; }
    public DbSet<TenantConnectionString> TenantConnectionStrings { get; set; }

    #endregion

    public ClientRegistrationSystemDbContext(DbContextOptions<ClientRegistrationSystemDbContext> options)
        : base(options)
    {

    }

    protected override void OnModelCreating(ModelBuilder builder)
    {
        base.OnModelCreating(builder);

        /* Include modules to your migration db context */

        builder.ConfigurePermissionManagement();
        builder.ConfigureSettingManagement();
        builder.ConfigureBackgroundJobs();
        builder.ConfigureAuditLogging();
        builder.ConfigureIdentity();
        builder.ConfigureOpenIddict();
        builder.ConfigureFeatureManagement();
        builder.ConfigureTenantManagement();

        builder.Entity<ProductCategory>(b =>
        {
            b.ToTable("ProductCategories");
            b.ConfigureByConvention();
            b.Property(x => x.SystemName).IsRequired();
            b.Property(x => x.DisplayName).IsRequired();
        });

        builder.Entity<ClientDetail>(b =>
        {
            b.ToTable("ClientDetails");
            b.ConfigureByConvention();
            b.Property(x => x.ClientId).IsRequired();
            b.Property(x => x.FirstName).IsRequired();
            b.Property(x => x.MiddleName).IsRequired();
            b.Property(x => x.LastName).IsRequired();
            b.Property(x => x.Address).IsRequired();
            b.Property(x => x.PhoneNumber).HasMaxLength(16).IsRequired();
            b.Property(x => x.Email).IsRequired();
        });

        builder.Entity<ClientProductCategory>(b =>
        {
            b.ToTable("ClientProductCategories");
            b.ConfigureByConvention();
            b.Property(x => x.ClientId).IsRequired();
            b.Property(x => x.ProductCategoryId).IsRequired();
        });


        var cascadeFKs = builder.Model.GetEntityTypes()
            .SelectMany(t => t.GetForeignKeys())
            .Where(fk => !fk.IsOwnership && fk.DeleteBehavior == DeleteBehavior.Cascade);

        foreach (var fk in cascadeFKs)
            fk.DeleteBehavior = DeleteBehavior.Restrict;
        /* Configure your own tables/entities inside here */

        //builder.Entity<YourEntity>(b =>
        //{
        //    b.ToTable(ClientRegistrationSystemConsts.DbTablePrefix + "YourEntities", ClientRegistrationSystemConsts.DbSchema);
        //    b.ConfigureByConvention(); //auto configure for the base class props
        //    //...
        //});
    }
}
