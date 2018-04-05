using Microsoft.Bot.Connector;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace CRMBot
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Login(string channelId, string userId, string extraQueryParams)
        {
            // CRM Url
            Session["channelId"] = channelId;
            Session["userId"] = userId;
            ChatState state = ChatState.RetrieveChatState(Session["channelId"].ToString(), Session["userId"].ToString());

            AuthenticationContext authContext = new AuthenticationContext(ConfigurationManager.AppSettings["CrmAuthority"]);
            var authUri = authContext.GetAuthorizationRequestUrlAsync(state.OrganizationUrl, ConfigurationManager.AppSettings["CrmClientId"],
            new Uri(ConfigurationManager.AppSettings["CrmRedirectUrl"]), UserIdentifier.AnyUser, extraQueryParams);
            return Redirect(authUri.Result.ToString());
        }

        public ActionResult Authorize(string code)
        {
            AuthenticationContext authContext = new AuthenticationContext(ConfigurationManager.AppSettings["CrmAuthority"]);
            var authResult = authContext.AcquireTokenByAuthorizationCodeAsync(
            code, new Uri(ConfigurationManager.AppSettings["CrmRedirectUrl"]),
            new ClientCredential(ConfigurationManager.AppSettings["CrmClientId"],
            ConfigurationManager.AppSettings["CrmClientSecret"]));

            // Saving token in Bot State
            var botCredentials = new MicrosoftAppCredentials(ConfigurationManager.AppSettings["MicrosoftAppId"],
            ConfigurationManager.AppSettings["MicrosoftAppPassword"]);
            ChatState state = ChatState.RetrieveChatState(Session["channelId"].ToString(), Session["userId"].ToString());
            state.AccessToken = authResult.Result.AccessToken;

            ViewBag.Message = $"Your Token - {authResult.Result.AccessToken} Channel Id - {Session["channelId"].ToString()} User Id - {Session["userId"].ToString()}";
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}