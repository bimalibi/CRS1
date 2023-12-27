using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Volo.Abp.Data;
using Volo.Abp.DependencyInjection;
using Volo.Abp.Domain.Repositories;
using YSJU.ClientRegistrationSystem.AppEntities.ProductCategories;
namespace YSJU.ClientRegistrationSystem.DataSeedContributer
{
    public class ProductCategoryDataSeedContributer : IDataSeedContributor, ITransientDependency
    {
        private readonly IRepository<ProductCategory, Guid> _productCategoryRepository;

        public ProductCategoryDataSeedContributer(
            IRepository<ProductCategory, Guid> productCategoryRepository)
        {
            _productCategoryRepository = productCategoryRepository;
        }

        public async Task SeedAsync(DataSeedContext context)
        {
            if (await _productCategoryRepository.CountAsync() > 0)
            {
                return;
            }
            var productCategory1 = new ProductCategory
            {
                DisplayName = "Software",
                SystemName = "Software",
            };
            var productCategory2 = new ProductCategory
            {
                DisplayName = "Electronic and Gadgets",
                SystemName = "Electronic and Gadgets",
            };
            var productCategory3 = new ProductCategory
            {
                DisplayName = "Consultation",
                SystemName = "Consultation",
            };
            await _productCategoryRepository.InsertManyAsync(new[] { productCategory1, productCategory2, productCategory3 }, autoSave: true);
        }
    } 
}
