using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Volo.Abp.Application.Dtos;
using Volo.Abp.Application.Services;
using YSJU.ClientRegistrationSystem.Dtos.ClientDetailDtos;
using YSJU.ClientRegistrationSystem.Dtos.ResponseDtos;

namespace YSJU.ClientRegistrationSystem.Interfaces.ClientDetailManagement
{
    public interface IClientDetailAppService : IApplicationService
    {
        Task<ResponseDto<ClientDetailResponseDto>> CreateClientDetailAsync(CreateClientDetailDto input);
        Task<ClientDetailResponseDto> GetClientDetailById(Guid clientPersonalDetailId);
        Task<ResponseDto<ClientDetailResponseDto>> UpdateClientDetail(Guid clientPersonalDetailId, UpdateClientDetailDto input);
        Task<PagedResultDto<ClientDetailResponseDto>> GetPagedAndSortedClientDetailAsync(PagedAndSortedClientDetailListDto input);
        Task<ResponseDto<ClientDetailResponseDto>> DeleteClientDetailAsync(Guid clientPersonalDetailId);
    }
}
