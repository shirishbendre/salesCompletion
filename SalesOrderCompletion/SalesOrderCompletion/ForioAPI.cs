using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;

/// <summary>
/// Summary description for ForioAPI
/// </summary>
public class API
{
    protected string _token = null;
    protected string _path;
    protected string _login;
    protected string _password;
    public int method;
    public string strForioAPIUrl = System.Configuration.ConfigurationManager.AppSettings["ForioAPIUrl"].ToString();

    GetConnection objGetConnection = new GetConnection();

    public API(string path, string login, string password)
    {
        _path = path;
        _login = login;
        _password = password;
    }
    protected dynamic get(string method)
    {
        return get(method, "");
    }
    protected dynamic get(string method, string data)
    {
        return get(method, data, "");
    }
    protected dynamic get(string method, string data, string append)
    {
        return call("GET", method, data, append);
    }
    protected dynamic post(string method, string postdata)
    {
        return post(method, postdata, "");
    }
    protected dynamic post(string method, string postdata, string append)
    {
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "ForioAPI:post(method,  postdata,  append)",
    " post(" + method + ", " + postdata + ", " + append + ")", null, objGetConnection.strPurchaseModuleName);
        return call("POST", method, postdata, append);
    }
    protected dynamic put(string method, string postdata)
    {
        return put(method, postdata, "");
    }
    protected dynamic put(string method, string postdata, string append)
    {
        return call("PUT", method, postdata, append);
    }
    public dynamic createSession(string sessionKey, string details)
    {
        sessionKey = HttpUtility.UrlEncode(sessionKey);
        return put("usergroup", details, sessionKey);
    }
    public dynamic getSession(string sessionKey)
    {
        dynamic response = get("usergroup", "", sessionKey);
        if (response["status_code"].ToString() == "200" && response["group"]["name"] == "SIMULATION_ROOT")
        {
            response["status_code"] = 404;
            response["group"] = null;
            response["message"] = "Cannot find a session with key '" + sessionKey + "'";
        }
        return response;
    }
    public dynamic setSession(string sessionKey)
    {
        string parameters = "user_action=changeGroup&userGroup=" + sessionKey;
        return post("authentication", parameters);
    }
    public dynamic assignTeamToAdmin(string groupId, string email, string team)
    {
        string details =
            "action=autoAssignAll&dryRun=false&team=" + team + "&role=single" +
            "content=" + HttpUtility.UrlEncode(email);

        return post("usergroup", details, groupId);
    }
    public dynamic assignUserToSessionDummy(string groupId, string email, string firstName, string lastName, string team)
    {

        // create the user
        string details =
            "action=addUsers&dryRun=false&" +
            "content=" + HttpUtility.UrlEncode(email) + "," +
                        HttpUtility.UrlEncode(firstName) + "," +
                        HttpUtility.UrlEncode(lastName) + ",," + team;

        return post("usergroup", details, groupId);
    }
    public dynamic assignUserToSession(string groupId, string email, string firstName, string lastName)
    {
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "ForioAPI:assignUserToSession(groupId,email,firstName,lastName)",
           "Parameter: groupId=" + groupId + ",email=" + email + ",firstName=" + firstName + ",lastName=" + lastName, null, objGetConnection.strPurchaseModuleName);

        // create the user
        string details =
            "action=addUsers&dryRun=false&" +
            "content=" + HttpUtility.UrlEncode(email) + "," +
                        HttpUtility.UrlEncode(firstName) + "," +
                        HttpUtility.UrlEncode(lastName);

        var result = post("usergroup", details, groupId);
        return result;
    }
    public dynamic assignAdminToSession(string groupId, string email, string firstName, string lastName)
    {

        assignUserToSession(groupId, email, firstName, lastName);

        // create the user
        string details =
            "action=setPermission&permission=FACILITATOR&" +
            "email=" + HttpUtility.UrlEncode(email);

        return post("usergroup", details, groupId);
    }
    private dynamic login()
    {
        string parameters = "email=" + _login + "&password=" + _password + "&user_action=login&token=true";
        dynamic response = post("authentication", parameters);
        _token = response["token"];
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "ForioAPI:login",
        "_token" + _token, null, objGetConnection.strPurchaseModuleName);

        return response;
    }
    private dynamic call(string callMethod, string apiMethod, string data)
    {
        return call(callMethod, apiMethod, data, "");
    }
    private dynamic call(string callMethod, string apiMethod, string data, string append)
    {
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "ForioAPI:call(callMethod,  apiMethod,  data,append)",
  " call(" + callMethod + ", " + apiMethod + ", " + data + ", " + append + ")", null, objGetConnection.strPurchaseModuleName);
        // log in if necessary
        if (data.IndexOf("user_action=login") == -1 && String.IsNullOrEmpty(_token))
        {
            login();
        }

        if (!String.IsNullOrEmpty(append))
        {
            append = "/" + append;
        }

        String path = apiMethod == "user" ? "" : "/" + _path;

        string url = strForioAPIUrl + apiMethod + path + append + "?token=" + _token;

        // add any get parametersi
        if (callMethod == "GET" && !String.IsNullOrEmpty(data))
        {
            url += "&" + data;
        }

        HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
        req.KeepAlive = true;
        req.ProtocolVersion = HttpVersion.Version10;
        req.Method = callMethod == "PUT" ? "POST" : callMethod;

        if (callMethod == "PUT")
        {
            data = "method=put&" + data;
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "ForioAPI:call(callmethod,apimethod,data,append)",
       "callMethod =" + callMethod, null, objGetConnection.strPurchaseModuleName);
        }

        // prepare post data
        if (callMethod == "POST" || callMethod == "PUT")
        {
            try
            {
                req.ContentType = "application/x-www-form-urlencoded";
                byte[] postBytes = Encoding.ASCII.GetBytes(data);
                req.ContentLength = postBytes.Length;
                Stream requestStream = req.GetRequestStream();
                requestStream.Write(postBytes, 0, postBytes.Length);
                requestStream.Close();
            }
            catch (Exception ex)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "ForioAPI:call(callmethod,apimethod,data,append)",
            " IF(callMethod == POST || callMethod == PUT)." + ex.Message, ex, objGetConnection.strPurchaseModuleName);
            }
        }

        string responseText = "";
        try
        {
            HttpWebResponse response = (HttpWebResponse)req.GetResponse();
            Stream resStream = response.GetResponseStream();
            var sr = new StreamReader(response.GetResponseStream());
            responseText = sr.ReadToEnd();

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "ForioAPI:call(callmethod,apimethod,data,append)",
          "responseText" + responseText, null, objGetConnection.strPurchaseModuleName);
        }
        catch (WebException e)
        {
            WebResponse response = e.Response;
            Stream resStream = response.GetResponseStream();

            var sr = new StreamReader(response.GetResponseStream());
            responseText = sr.ReadToEnd();
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "ForioAPI:call(callmethod,apimethod,data,append)",
        "Inside Web Exception :responseText" + responseText, null, objGetConnection.strPurchaseModuleName);
        }

        dynamic responseObject = (new JavaScriptSerializer()).DeserializeObject(responseText);

        return responseObject;
    }
    public dynamic SendLogins(string email)
    {
        string parameters = "email=" + email + "&user_action=emailUserPassword&token=false";
        dynamic response = post("authentication", parameters);
        //_token = response["token"];
        return response;
    }
}