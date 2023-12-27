using System;
using System.Collections.Generic;
using System.Text;
using Volo.Abp.Application.Dtos;

namespace YSJU.ClientRegistrationSystem.Dtos.ClientDetailDtos
{
    public class PagedAndSortedClientDetailListDto : PagedAndSortedResultRequestDto
    {
        public string? SearchKeyword { get; set; }
        public string? SortOrder { get; set; } = "asc";
        public Guid? ProductCategoryId { get; set; }
    }
}
