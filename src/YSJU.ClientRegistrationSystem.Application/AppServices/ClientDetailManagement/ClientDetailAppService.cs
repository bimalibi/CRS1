using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Text;
using System.Threading.Tasks;
using Volo.Abp.Application.Dtos;
using Volo.Abp;
using Volo.Abp.Application.Services;
using Volo.Abp.Domain.Repositories;
using YSJU.ClientRegistrationSystem.AppEntities.ClientDetails;
using YSJU.ClientRegistrationSystem.AppEntities.ProductCategories;
using YSJU.ClientRegistrationSystem.Interfaces.ClientDetailManagement;
using YSJU.ClientRegistrationSystem.Dtos.ResponseDtos;
using YSJU.ClientRegistrationSystem.Dtos.ClientDetailDtos;
using System.Net.WebSockets;
using YSJU.ClientRegistrationSystem.Dtos.ClientDetailManagementDtos;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using YSJU.ClientRegistrationSystem.AppEntities.ClientProductCategories;
using YSJU.ClientRegistrationSystem.Dtos.ProductCategoryManagementDtos;

namespace YSJU.ClientRegistrationSystem.AppServices.ClientDetailManagement
{
    public class ClientDetailAppService : ApplicationService, IClientDetailAppService
    {
        private readonly IRepository<ClientDetail, Guid> _clientPersonalDetailRepository;
        private readonly IRepository<ProductCategory, Guid> _productCategoryRepository;
        private readonly IRepository<ClientProductCategory, Guid> _clientProductCategory;

        public ClientDetailAppService(
            IRepository<ClientDetail, Guid> clientPersonalDetailRepository,
            IRepository<ProductCategory, Guid> productCategoryRepository,
            IRepository<ClientProductCategory, Guid> clientProductCategory)
        {
            _clientPersonalDetailRepository = clientPersonalDetailRepository;
            _productCategoryRepository = productCategoryRepository;
            _clientProductCategory = clientProductCategory;
        }

        public async Task<ResponseDto<ClientDetailResponseDto>> CreateClientDetailAsync(CreateClientDetailDto input)
        {
            try
            {
                Logger.LogInformation($"CreateClientDetailAsync requested by User: {CurrentUser.Id}");
                Logger.LogDebug($"CreateClientDetailAsync requested for User: {(CurrentUser.Id, input)}");

                var clientPersonalDetailQuery = await _clientPersonalDetailRepository.GetQueryableAsync();
                var productCategoryQuery = await _productCategoryRepository.GetQueryableAsync();

                if (productCategoryQuery.Any(x => !input.ProductCategoryId.Contains(x.Id)))
                {
                    throw new UserFriendlyException("Product not found", code: "400");
                }

                if (clientPersonalDetailQuery.Where(x => x.Email == input.Email).Any())
                {
                    throw new UserFriendlyException("Email already exist", code: "400");
                }

                if (clientPersonalDetailQuery.Where(x => x.PhoneNumber == input.PhoneNumber).Any())
                {
                    throw new UserFriendlyException("Phone number already exist", code: "400");
                }

                var nextClientId = clientPersonalDetailQuery.Select(x => x.ClientId).Max();
                var newClientPersonalDetails = new ClientDetail
                {
                    ClientId = (nextClientId  == 0) ? 100 : nextClientId + 1,
                    FirstName = input.FirstName,
                    MiddleName = input.MiddleName,
                    LastName = input.LastName,
                    Address = input.Address,
                    PhoneNumber = input.PhoneNumber,
                    Email = input.Email,
                };

                await _clientPersonalDetailRepository.InsertAsync(newClientPersonalDetails, true);

                foreach(var item in input.ProductCategoryId)
                {
                    var clientProductCategory = new ClientProductCategory
                    {
                        ClientId = newClientPersonalDetails.Id,
                        ProductCategoryId = item
                    };

                    await _clientProductCategory.InsertAsync(clientProductCategory, true);
                }

                var response = new ResponseDto<ClientDetailResponseDto>
                {
                    Success = true,
                    Code = 200,
                    Message = "New Client Details added successfully",
                    Data = null
                };

                Logger.LogInformation($"CreateClientDetailAsync responded for User: {CurrentUser.Id}");

                return response;
            }
            catch (Exception)
            {
                Logger.LogError(nameof(CreateClientDetailAsync));
                throw;
            }
        }

        public async Task<ClientDetailResponseDto> GetClientDetailById(Guid clientPersonalDetailId)
        {
            try
            {
                Logger.LogInformation($"GetClientDetailById requested by User: {CurrentUser.Id}");
                Logger.LogDebug($"GetClientDetailById requested for User: {(CurrentUser.Id, clientPersonalDetailId)}");

                var clientPersonalDetailQuery = await _clientPersonalDetailRepository.GetQueryableAsync();
                var productCategoryQuery = await _productCategoryRepository.GetQueryableAsync();
                var clientProductCategoryQuery = await _clientProductCategory.GetQueryableAsync();

                var clientPersonalDetail = clientPersonalDetailQuery.Where(x => x.Id == clientPersonalDetailId).FirstOrDefault()
                    ?? throw new UserFriendlyException("Client Personal Detail not found", code: "400");

                var clientPersonalDetailData = (from clientPersonalDetails in clientPersonalDetailQuery
                                                join clientProductCategory in clientProductCategoryQuery on clientPersonalDetail.Id equals clientProductCategory.ClientId into clientProductCategoryLeft
                                                from clientProductCategory in clientProductCategoryLeft.DefaultIfEmpty()
                                                join productCategory in productCategoryQuery on clientProductCategory.ProductCategoryId equals productCategory.Id into productCategoryLeft
                                                from productCategory in productCategoryLeft.DefaultIfEmpty()
                                                where clientPersonalDetail.Id == clientPersonalDetailId
                                                select new ClientDetailResponseDto
                                                {
                                                    ClientId = clientPersonalDetail.ClientId,
                                                    FirstName = clientPersonalDetail.FirstName,
                                                    MiddleName = clientPersonalDetail.MiddleName,
                                                    LastName = clientPersonalDetail.LastName,
                                                    Address = clientPersonalDetail.Address,
                                                    PhoneNumber = clientPersonalDetail.PhoneNumber,
                                                    Email = clientPersonalDetail.Email,
                                                    ProductCategoryId = clientProductCategory.Id,
                                                    ProductCategoryName = productCategory.DisplayName,
                                                }).FirstOrDefault();

                Logger.LogInformation($"GetClientDetailById responded for User: {CurrentUser.Id}");

                return clientPersonalDetailData;
            }
            catch (Exception)
            {
                Logger.LogError(nameof(GetClientDetailById));
                throw;
            }
        }

        public async Task<ResponseDto<ClientDetailResponseDto>> UpdateClientDetail(Guid clientPersonalDetailId, UpdateClientDetailDto input)
        {
            try
            {
                Logger.LogInformation($"UpdateClientDetail requested by User: {CurrentUser.Id}");
                Logger.LogDebug($"UpdateClientDetail requested for User: {(CurrentUser.Id, input)}");

                var clientPersonalDetailQuery = await _clientPersonalDetailRepository.GetQueryableAsync();
                var productCategoryQuery = await _productCategoryRepository.GetQueryableAsync();
                var clientProductCategoryQuery = await _clientProductCategory.GetQueryableAsync();

                var clientPersonalDetail = clientPersonalDetailQuery.Where(x => x.Id == clientPersonalDetailId).FirstOrDefault()
                    ?? throw new UserFriendlyException("Client Personal Detail not found", code: "400");


                clientPersonalDetail.FirstName = input.FirstName;
                clientPersonalDetail.MiddleName = input.MiddleName;
                clientPersonalDetail.LastName = input.LastName;
                clientPersonalDetail.Address = input.Address;
                clientPersonalDetail.PhoneNumber = input.PhoneNumber;

                var clientProductCategoryToRemove = clientProductCategoryQuery.Where(x => x.ClientId == clientPersonalDetailId && !input.ProductCategoryId.Contains(x.ProductCategoryId)).ToList();

                foreach (var item in input.ProductCategoryId)
                {
                    if(!clientProductCategoryQuery.Where(x => x.ClientId == clientPersonalDetailId && x.ProductCategoryId == item).Any())
                    {
                        var clientProductCategory = new ClientProductCategory
                        {
                            ClientId = clientPersonalDetailId,
                            ProductCategoryId = item
                        };
                        await _clientProductCategory.InsertAsync(clientProductCategory);
                    }
                }

                if (clientProductCategoryToRemove != null)
                {
                    await _clientProductCategory.DeleteManyAsync(clientProductCategoryToRemove);
                }
                await _clientPersonalDetailRepository.UpdateAsync(clientPersonalDetail);

                var response = new ResponseDto<ClientDetailResponseDto>
                {
                    Success = true,
                    Code = 200,
                    Message = "New Client Details updated successfully",
                    Data = null
                };

                Logger.LogInformation($"UpdateClientDetail responded for User: {CurrentUser.Id}");

                return response;
            }
            catch (Exception)
            {
                Logger.LogError(nameof(UpdateClientDetail));
                throw;
            }
        }

        public async Task<PagedResultDto<ClientDetailResponseDto>> GetPagedAndSortedClientDetailAsync(PagedAndSortedClientDetailListDto input)
        {
            try
            {
                var clientPersonalDetailQuery = await _clientPersonalDetailRepository.GetQueryableAsync();
                var productCategoryQuery = await _productCategoryRepository.GetQueryableAsync();
                var clientProductCategoryQuery = await _clientProductCategory.GetQueryableAsync();


                if (input.Sorting.IsNullOrWhiteSpace())
                {
                    input.Sorting = "CreationTime";
                }

                input.Sorting = $"{input.Sorting} {input.SortOrder}";

                if (!string.IsNullOrWhiteSpace(input.SearchKeyword))
                {
                    clientPersonalDetailQuery = clientPersonalDetailQuery.Where(x =>
                        x.ClientId.ToString().Contains(input.SearchKeyword.ToLower()) ||
                        x.FirstName.ToLower().Contains(input.SearchKeyword.ToLower()) ||
                        x.MiddleName.ToLower().Contains(input.SearchKeyword.ToLower()) ||
                        x.LastName.ToLower().Contains(input.SearchKeyword.ToLower()) ||
                        x.Address.ToLower().Contains(input.SearchKeyword.ToLower()) ||
                        x.Email.ToLower().Contains(input.SearchKeyword.ToLower()));
                }

                var query = (from clientDetail in clientPersonalDetailQuery
                             join clientProductCategory in clientProductCategoryQuery on clientDetail.Id equals clientProductCategory.ClientId into clientProductCategoryLeft
                             from clientProductCategory in clientProductCategoryLeft.DefaultIfEmpty()
                             join productCategory in productCategoryQuery on clientProductCategory.ProductCategoryId equals productCategory.Id into productCategoryLeft
                             from productCategory in productCategoryLeft.DefaultIfEmpty()
                             select new ClientDetailResponseDto
                             {
                                 Id = clientDetail.Id,
                                 ClientId = clientDetail.ClientId,
                                 FirstName = clientDetail.FirstName,
                                 MiddleName = clientDetail.MiddleName,
                                 LastName = clientDetail.LastName,
                                 Address = clientDetail.Address,
                                 PhoneNumber = clientDetail.PhoneNumber,
                                 Email = clientDetail.Email,
                                 ProductCategoryId = clientProductCategory.ProductCategoryId,
                                 ProductCategoryName = productCategory.DisplayName,
                                 CreationTime = clientDetail.CreationTime
                             });

                if (input.ProductCategoryId != null)
                {
                    query = query.Where(x => x.ProductCategoryId == input.ProductCategoryId);
                }

                var result = query.OrderBy(input.Sorting)
                                  .Skip(input.SkipCount)
                                  .Take(input.MaxResultCount).ToList();

                var totalCount = query.Count();
                var response = new PagedResultDto<ClientDetailResponseDto>(totalCount, result);

                return response;
            }
            catch (Exception)
            {
                Logger.LogError(nameof(GetPagedAndSortedClientDetailAsync));
                throw;
            }
        }

        public async Task<ResponseDto<ClientDetailResponseDto>> DeleteClientDetailAsync(Guid clientPersonalDetailId)
        {
            try
            {
                Logger.LogInformation($"DeleteClientDetailAsync requested by User: {CurrentUser.Id}");
                Logger.LogDebug($"DeleteClientDetailAsync requested for User: {(CurrentUser.Id, clientPersonalDetailId)}");

                var clientPersonalDetailQuery = await _clientPersonalDetailRepository.GetQueryableAsync();
                var productCategoryQuery = await _productCategoryRepository.GetQueryableAsync();
                var clientProductCategoryQuery = await _clientProductCategory.GetQueryableAsync();

                var clientPersonalDetail = clientPersonalDetailQuery.Where(x => x.Id == clientPersonalDetailId).FirstOrDefault()
                    ?? throw new UserFriendlyException("Client Personal Detail not found", code: "400");

                var clientProductCategory = clientProductCategoryQuery.Where(x => x.ClientId == clientPersonalDetailId).ToList();

                await _clientProductCategory.DeleteManyAsync(clientProductCategory);
                await _clientPersonalDetailRepository.DeleteAsync(clientPersonalDetail);

                Logger.LogInformation($"DeleteClientDetailAsync responded for User: {CurrentUser.Id}");

                var response = new ResponseDto<ClientDetailResponseDto>
                {
                    Success = true,
                    Message = "Client personal detail deleted successfully",
                    Code = 200,
                    Data = null
                };
                return response;
            }
            catch (Exception)
            {
                Logger.LogError(nameof(DeleteClientDetailAsync));
                throw;
            }
        }

        public async Task<ExportClientDetailDto> ExportAllClientDetail()
        {
            try
            {
                Logger.LogInformation($"ExportAllClientDetail requested by User: {CurrentUser.Id}");

                var clientPersonalDetailQuery = await _clientPersonalDetailRepository.GetQueryableAsync();
                var productCategoryQuery = await _productCategoryRepository.GetQueryableAsync();
                var clientProductCategoryQuery = await _clientProductCategory.GetQueryableAsync();

                var query = (from clientDetail in clientPersonalDetailQuery
                             join clientProductCategory in clientProductCategoryQuery on clientDetail.Id equals clientProductCategory.ClientId into clientProductCategoryLeft
                             from clientProductCategory in clientProductCategoryLeft.DefaultIfEmpty()
                             join productCategory in productCategoryQuery on clientProductCategory.ProductCategoryId equals productCategory.Id into productCategoryLeft
                             from productCategory in productCategoryLeft.DefaultIfEmpty()
                             select new ClientDetailResponseDto
                             {
                                 Id = clientDetail.Id,
                                 ClientId = clientDetail.ClientId,
                                 FirstName = clientDetail.FirstName,
                                 MiddleName = clientDetail.MiddleName,
                                 LastName = clientDetail.LastName,
                                 Address = clientDetail.Address,
                                 PhoneNumber = clientDetail.PhoneNumber,
                                 Email = clientDetail.Email,
                                 ProductCategoryId = clientProductCategory.ProductCategoryId,
                                 ProductCategoryName = productCategory.DisplayName,
                                 CreationTime = clientDetail.CreationTime
                             });

                var clientDetailList = query.ToList();
                var stream = new MemoryStream();

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var excelPackage = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Employee Personal Details");

                    worksheet.Row(1).Height = 20;
                    worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Row(1).Style.Font.Bold = true;

                    var headingRowIndex = 1;
                    var headingColumnIndex = 1;

                    worksheet.Cells[headingRowIndex, headingColumnIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = "Client Id";
                    worksheet.Cells[headingRowIndex, headingColumnIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = "Full Name";
                    worksheet.Cells[headingRowIndex, headingColumnIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = "Address";
                    worksheet.Cells[headingRowIndex, headingColumnIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = "Email";
                    worksheet.Cells[headingRowIndex, headingColumnIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = "Phone Number";
                    worksheet.Cells[headingRowIndex, headingColumnIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = "Product";
                    worksheet.Cells[headingRowIndex, headingColumnIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = "Product Category";
                    headingRowIndex++;
                    var sNo = 1;

                    foreach (var rowData in clientDetailList)
                    {
                        headingColumnIndex = 1;
                        worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = rowData.ClientId;
                        worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = $"{rowData.FirstName} {rowData.MiddleName} {rowData.LastName}";
                        worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = rowData.Address;
                        worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = rowData.Email;
                        worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = rowData.PhoneNumber;
                        worksheet.Cells[headingRowIndex, headingColumnIndex++].Value = rowData.ProductCategoryName;
                        headingRowIndex++;
                        sNo++;
                    }
                    excelPackage.Save();

                    var excelFileStream = await Task.FromResult(stream);

                    Logger.LogInformation($"ExportAllClientDetail responded for User: {CurrentUser.Id}");

                    return new ExportClientDetailDto
                    {
                        Name = $"ClientDetails-{DateTime.Now:yyyy-MM-dd-mm-ss}",
                        content = excelFileStream.ToArray()
                    };
                }
            }
            catch (Exception)
            {
                Logger.LogError(nameof(ExportAllClientDetail));
                throw;
            }
        }

        public async Task<List<ProductCategoryResponseDto>> GetProductCategoryListAsync()
        {
            try
            {
                var peoductCategoryQuery = await _productCategoryRepository.GetQueryableAsync();

                var result = peoductCategoryQuery.Select(x => new ProductCategoryResponseDto
                {
                    ProductCategoryId = x.Id,
                    ProductCategoryName = x.DisplayName
                }).ToList();

                return result;
            }
            catch (Exception)
            {
                Logger.LogError(nameof(GetProductCategoryListAsync));
                throw;
            }
        }
    }
}
