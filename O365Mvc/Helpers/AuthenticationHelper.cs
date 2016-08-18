using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.SharePoint.CoreServices;
using O365Mvc.Models;
using O365Mvc.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;

namespace O365Mvc.Helpers
{
    internal class AuthenticationHelper
    {
        //internal static async Task<string> GetTokenAsync()
        //{

        //    var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

        //    var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

        //    AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.Authority, new ADALTokenCache(signInUserId));

        //    var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId,
        //                                                   new ClientCredential(SettingsHelper.ClientId,
        //                                                                        SettingsHelper.ClientSecret),
        //                                                   new UserIdentifier(userObjectId,
        //                                                                      UserIdentifierType.UniqueId));

        //    return authResult.AccessToken;
        //}

        //OutlookServicesClient
        internal static async Task<OutlookServicesClient> EnsureOutlookServicesClientCreatedAsync(string capabilityName)
        {

            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            //AuthenticationContext(authority,tokenCache)
            AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.Authority, new ADALTokenCache(signInUserId));

            try
            {
                //DiscoveryClient(DiscoveryServiceEndpointUri,AccessToken)
                DiscoveryClient discClient = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                    async () =>
                    {
                        //AcquireTokenSilentAsync(DiscoveryServiceResourceId,ClientCredential,UserIdentifier)
                        var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId,
                                                                                   new ClientCredential(SettingsHelper.ClientId,
                                                                                                        SettingsHelper.ClientSecret),
                                                                                   new UserIdentifier(userObjectId,
                                                                                                      UserIdentifierType.UniqueId));

                        return authResult.AccessToken;
                    });

                var dcr = await discClient.DiscoverCapabilityAsync(capabilityName);

                return new OutlookServicesClient(dcr.ServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId,
                                                                                   new ClientCredential(SettingsHelper.ClientId,
                                                                                                        SettingsHelper.ClientSecret),
                                                                                   new UserIdentifier(userObjectId,
                                                                                                      UserIdentifierType.UniqueId));

                        return authResult.AccessToken;
                    });
            }
            catch (AdalException exception)
            {
                //Handle token acquisition failure
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    authContext.TokenCache.Clear();
                    throw exception;
                }
                return null;
            }
        }

        internal static async Task<SharePointClient> EnsureSharePointClientCreatedAsync(string capabilityName)
        {

            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.Authority, new ADALTokenCache(signInUserId));

            try
            {
                DiscoveryClient discClient = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId,
                                                                                   new ClientCredential(SettingsHelper.ClientId,
                                                                                                        SettingsHelper.ClientSecret),
                                                                                   new UserIdentifier(userObjectId,
                                                                                                      UserIdentifierType.UniqueId));
                        return authResult.AccessToken;
                    });

                var dcr = await discClient.DiscoverCapabilityAsync(capabilityName);

                return new SharePointClient(dcr.ServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId,
                                                                                   new ClientCredential(SettingsHelper.ClientId,
                                                                                                        SettingsHelper.ClientSecret),
                                                                                   new UserIdentifier(userObjectId,
                                                                                                      UserIdentifierType.UniqueId));

                        return authResult.AccessToken;
                    });
            }
            catch (AdalException exception)
            {
                //Partially handle token acquisition failure here and bubble it up to the controller
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    authContext.TokenCache.Clear();
                    throw exception;
                }
                return null;
            }
        }

    }
    //Capabilities
    public enum O365Capabilities
    {
        Mail,
        Calendar,
        Contacts,
        MyFiles,
    }

}