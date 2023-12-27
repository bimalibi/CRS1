using YSJU.ClientRegistrationSystem.Localization;
using Volo.Abp.AspNetCore.Mvc;

namespace YSJU.ClientRegistrationSystem.Controllers;

/* Inherit your controllers from this class.
 */
public abstract class ClientRegistrationSystemController : AbpControllerBase
{
    protected ClientRegistrationSystemController()
    {
        LocalizationResource = typeof(ClientRegistrationSystemResource);
    }
}
