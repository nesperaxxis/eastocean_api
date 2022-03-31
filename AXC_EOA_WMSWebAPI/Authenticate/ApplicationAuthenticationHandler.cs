using System;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Http;
using System.Web.Http.Filters;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using System.Net.Http.Formatting;
using System.Security.Principal;
using System.Web.Http;
using System.Web.Http.Controllers;
using System.Web.Http.Results;
using Newtonsoft.Json;
using SBOCLASS.Class;
using System.Data.SqlClient;
using System.Data;


namespace AXC_EOA_WMSWebAPI.Authenticate
{
public class IdentityBasicAuthentication : Attribute, IAuthenticationFilter
{
   //private readonly IAuthenticator _authenticator;
   // private string data;
    public bool AllowMultiple
    {
        get { return true; }
    }

    public async Task AuthenticateAsync(HttpAuthenticationContext context, CancellationToken cancellationToken)
    {
        
        await Task.Factory.StartNew(() =>
        {
            HttpRequestMessage request = context.Request;
            AuthenticationHeaderValue authorization = request.Headers.Authorization;
            SBOCLASS.Models.ResponseResult response;

            if (authorization == null)
            {
                context.ErrorResult = new AuthenticationFailureResult("Missing authorization header", request);
                return;
            }

            if (authorization.Scheme != "Basic")
            {
                context.ErrorResult = new AuthenticationFailureResult("Authorization scheme not supported", request);
                return;
            }

            if (string.IsNullOrEmpty(authorization.Parameter))
            {
                context.ErrorResult = new AuthenticationFailureResult("Missing credentials", request);
                return;
            }

            Tuple<string, string> userNameAndPasword = ExtractUserNameAndPassword(authorization.Parameter);
            if (userNameAndPasword == null)
            {
                context.ErrorResult = new AuthenticationFailureResult("Invalid credentials", request);
            }
            else
            {
                string userName = userNameAndPasword.Item1;
                string password = userNameAndPasword.Item2;
                DateTime now = DateTime.Now;
                //you may need to decide here how to verify the user. if you have saved in db, then check in db

                string connStr = System.Configuration.ConfigurationManager.ConnectionStrings["SAPConnString"].ToString();

                AuthenticationUserPass auth = new AuthenticationUserPass();
                var returnValue = ValidateAuthenticateUser(userName,password);//auth.AuthenticateUserPass(connStr, userName, password);

                if (returnValue == 0)
                {
                    var identity = new GenericIdentity(userName, "Basic");
                    //if you need authorization as well, then fetch the roles and add it here
                    // context.Principal = new GenericPrincipal(identity, _authenticator.GetRoles());
                    string[] dat = new string[2];
                    dat[0] = "Admin";
                    dat[1] = "Admin";

                    context.Principal = new GenericPrincipal(identity, dat);
                }

                else
                {
                    if (returnValue == 2) 
                    {
                        response = new SBOCLASS.Models.ResponseResult();
                        response.RecordStatus = "false";
                        response.ErrorDescription = "UnAuthorised Error! " + " Authenticated User expires";
                        context.ErrorResult = new AuthenticationFailureResult(response, request);
                        return;
                    }
                    else
                    {
                        var identity = new GenericIdentity("ERROR", "ERROR");
                        string[] dat = new string[2];
                        dat[0] = "ERROR";
                        dat[1] = "UnAuthorised Error! " + "Username and Password not match";
                        response = new SBOCLASS.Models.ResponseResult();
                        response.RecordStatus = "false";
                        response.ErrorDescription = "UnAuthorised Error! " + "Authenticated User and Password not match";
                        context.Principal = new GenericPrincipal(identity, dat);
                        //context.ErrorResult = new UnauthorizedResult(new AuthenticationHeaderValue[0], context.Request);
                        context.ErrorResult = new AuthenticationFailureResult(response, request);
                    }
                }
            }
        });
    }

    #region " ValidateAuthenticateData "
    int ValidateAuthenticateUser(string RequserName, string ReqPassword)
    {
        try
        {
            int intReturn;
            string USER = System.Configuration.ConfigurationManager.AppSettings["USER"].ToString();
            string PWORD = System.Configuration.ConfigurationManager.AppSettings["PW"].ToString();

            if (USER == RequserName & PWORD == ReqPassword)
            {
                intReturn = 0;
            }
            else
            {
                intReturn = - 1;
            }
            return intReturn;
        }
        catch (Exception ex)
        {
            throw ex;
        }

    }
    #endregion
    public Task ChallengeAsync(HttpAuthenticationChallengeContext context, CancellationToken cancellationToken)
    {
        var host = context.Request.RequestUri.DnsSafeHost;
        var challenge = new AuthenticationHeaderValue("Basic");
        context.Result = new AddChallengeOnUnauthorizedResult(challenge, context.Result);
        return Task.FromResult(0);
    }

    private static Tuple<string, string> ExtractUserNameAndPassword(string authorizationParameter)
    {
        byte[] credentialBytes;

        try
        {
            credentialBytes = Convert.FromBase64String(authorizationParameter);
        }
        catch (FormatException)
        {
            return null;
        }

        // The currently approved HTTP 1.1 specification says characters here are ISO-8859-1.
        // However, the current draft updated specification for HTTP 1.1 indicates this encoding is infrequently
        // used in practice and defines behavior only for ASCII.
        Encoding encoding = Encoding.ASCII;
        // Make a writable copy of the encoding to enable setting a decoder fallback.
        encoding = (Encoding)encoding.Clone();
        // Fail on invalid bytes rather than silently replacing and continuing.
        encoding.DecoderFallback = DecoderFallback.ExceptionFallback;
        string decodedCredentials;

        try
        {
            decodedCredentials = encoding.GetString(credentialBytes);
        }
        catch (DecoderFallbackException)
        {
            return null;
        }

        if (String.IsNullOrEmpty(decodedCredentials))
        {
            return null;
        }

        int colonIndex = decodedCredentials.IndexOf(':');

        if (colonIndex == -1)
        {
            return null;
        }

        string userName = decodedCredentials.Substring(0, colonIndex);
        string password = decodedCredentials.Substring(colonIndex + 1);
        return new Tuple<string, string>(userName, password);
    }
    }

    public class AddChallengeOnUnauthorizedResult : IHttpActionResult
{
    public AddChallengeOnUnauthorizedResult(AuthenticationHeaderValue challenge, IHttpActionResult innerResult)
    {
        Challenge = challenge;
        InnerResult = innerResult;
    }

    public AuthenticationHeaderValue Challenge { get; private set; }

    public IHttpActionResult InnerResult { get; private set; }

    public async Task<HttpResponseMessage> ExecuteAsync(CancellationToken cancellationToken)
    {
        HttpResponseMessage response = await InnerResult.ExecuteAsync(cancellationToken);

        if (response.StatusCode == HttpStatusCode.Unauthorized)
        {
            // Only add one challenge per authentication scheme.
            if (!response.Headers.WwwAuthenticate.Any((h) => h.Scheme == Challenge.Scheme))
            {
                response.Headers.WwwAuthenticate.Add(Challenge);
            }
        }

        return response;
    }
}

    public class RequireAdminAttribute : AuthorizeAttribute
    {
        protected override bool IsAuthorized(HttpActionContext context)
        {
            var principal = context.Request.GetRequestContext().Principal as GenericPrincipal;
            return principal.IsInRole("Admin");
        }
    }

    public class AuthenticationFailureResult : IHttpActionResult
    {
        public AuthenticationFailureResult(object jsonContent, HttpRequestMessage request)
        {
            JsonContent = jsonContent;
            Request = request;
        }

        public HttpRequestMessage Request { get; private set; }

        public Object JsonContent { get; private set; }

        public Task<HttpResponseMessage> ExecuteAsync(CancellationToken cancellationToken)
        {
            return Task.FromResult(Execute());
        }

        private HttpResponseMessage Execute()
        {
            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.Unauthorized);
            response.RequestMessage = Request;
            response.Content = new ObjectContent(JsonContent.GetType(), JsonContent, new JsonMediaTypeFormatter());
            return response;
        }
    }

  

}
