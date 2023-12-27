using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Volo.Abp.Domain.Entities.Auditing;

namespace YSJU.ClientRegistrationSystem.AppEntities.ProductCategories
{
    public class ProductCategory : FullAuditedAggregateRoot<Guid>
    {
        public string DisplayName { get; set; }
        public string SystemName { get; set; }
    }
}
