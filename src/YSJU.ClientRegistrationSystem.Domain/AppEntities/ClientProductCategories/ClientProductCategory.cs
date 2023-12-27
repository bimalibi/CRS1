using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Volo.Abp.Domain.Entities.Auditing;

namespace YSJU.ClientRegistrationSystem.AppEntities.ClientProductCategories
{
    public class ClientProductCategory : AuditedAggregateRoot<Guid>
    {
        public Guid ClientId { get; set; }
        public Guid ProductCategoryId { get; set; }
    }
}
