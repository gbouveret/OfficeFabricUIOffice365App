using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using OfficeFabricUIAppForO365.Models;
using OfficeFabricUIAppForO365.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace OfficeFabricUIAppForO365.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        public async Task<ActionResult> Index()        
        {
            List<EmailMessage> myMessages = new List<EmailMessage>();

            if (User.Identity.IsAuthenticated)
            {
                var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
                var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

                AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.Authority, new ADALTokenCache(signInUserId));

                try
                {
                    DiscoveryClient discClient = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                        async () =>
                        {
                            var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId, new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                            return authResult.AccessToken;
                        });

                    var dcr = await discClient.DiscoverCapabilityAsync("Mail");

                    OutlookServicesClient exClient = new OutlookServicesClient(dcr.ServiceEndpointUri,
                        async () =>
                        {
                            var authResult = await authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId, new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                            return authResult.AccessToken;
                        });

                    var messagesResult = await exClient.Me.Folders.GetById("Inbox").Messages.Take(20).ExecuteAsync();

                    //do
                    //{
                        var msgs = messagesResult.CurrentPage;
                        foreach (var m in msgs)
                        {
                            myMessages.Add(new EmailMessage { 
                                Subject = m.Subject,
                                Received = m.DateTimeReceived.Value.DateTime,
                                FromName = m.From.EmailAddress.Name,
                                FromEmail = m.From.EmailAddress.Address,
                                HasAttachments = m.HasAttachments.GetValueOrDefault(false),
                                Importance = m.Importance.ToString(),
                                Preview = m.BodyPreview,
                                IsRead = m.IsRead.GetValueOrDefault(false)
                            });
                        }

                        //messagesResult = await messagesResult.GetNextPageAsync();
                        //messagesResult = null;

                    //} while (messagesResult != null);
                }
                catch (AdalException exception)
                {
                    //handle token acquisition failure
                    if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                    {
                        authContext.TokenCache.Clear();

                        //handle token acquisition failure
                    }
                }
            }

            return View(myMessages);
        }

        public async Task<ActionResult> Contacts()
        {
            List<Persona> myContacts = new List<Persona>();

            if (User.Identity.IsAuthenticated)
            {
                var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
                var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

                AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.Authority, new ADALTokenCache(signInUserId));

                try
                {
                    DiscoveryClient discClient = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                        async () =>
                        {
                            var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId, new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                            return authResult.AccessToken;
                        });

                    var dcr = await discClient.DiscoverCapabilityAsync("Contacts");

                    OutlookServicesClient exClient = new OutlookServicesClient(dcr.ServiceEndpointUri,
                        async () =>
                        {
                            var authResult = await authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId, new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                            return authResult.AccessToken;
                        });

                    var contactsResult = await exClient.Me.Contacts.Take(20).ExecuteAsync();

                    //do
                    //{
                        var contacts = contactsResult.CurrentPage;
                        foreach (var c in contacts)
                        {
                            var firstEmail = c.EmailAddresses.FirstOrDefault();
                            string email = "";
                            if (firstEmail != null) email = firstEmail.Address;

                            myContacts.Add(new Persona
                            {
                                DisplayName = c.DisplayName,
                                Email = email,
                                JobTitle = c.JobTitle,
                                CompanyName = c.CompanyName,
                                Im = c.ImAddresses.FirstOrDefault(),
                                OfficeLocation = c.OfficeLocation,
                                Phone = c.BusinessPhones.FirstOrDefault()
                            });
                        }

                    //    contactsResult = await contactsResult.GetNextPageAsync();

                    //} while (contactsResult != null);
                }
                catch (AdalException exception)
                {
                    //handle token acquisition failure
                    if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                    {
                        authContext.TokenCache.Clear();

                        //handle token acquisition failure
                    }
                }
            }

            return View(myContacts);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Hey, this is a sample ! Follow me on Twitter @gbouveret";

            return View();
        }
    }
}