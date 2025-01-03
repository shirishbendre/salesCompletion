using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.Collections.Specialized;
using System.Text;
using System.IO;
using System.Web.Script.Serialization;
using System.Text.RegularExpressions;
using System.Net.Http;
using System.Net.Http.Headers;
//using System.Runtime.Serialization;

namespace ForioEpicenter
{
    public class EpicenterAPI
    {
        string responseText = "";
        public string _publicKey = "";
        public string _secretkey = "";

        public string strForioEpicenterOauthUrl = System.Configuration.ConfigurationManager.AppSettings["ForioEpicenterOauthUrl"].ToString();
        public string strForioEpicenterProjectUrl = System.Configuration.ConfigurationManager.AppSettings["ForioEpicenterProjectUrl"].ToString();
        public string strForioEpicenterGroupUrl = System.Configuration.ConfigurationManager.AppSettings["ForioEpicenterGroupUrl"].ToString();
        public string strForioEpicenterUserUrl = System.Configuration.ConfigurationManager.AppSettings["ForioEpicenterUserUrl"].ToString();
        public string strForioEpicenterPassRecoveryUrl = System.Configuration.ConfigurationManager.AppSettings["ForioEpicenterPassRecoveryUrl"].ToString();

        public string strForioEpicenterAddMemberUrl = System.Configuration.ConfigurationManager.AppSettings["ForioEpicenterAddMemberUrl"].ToString();


        GetConnection objGetConnection = new GetConnection();
        public EpicenterAPI(string publickey, string secretkey)
        {
            _publicKey = publickey;
            _secretkey = secretkey;
        }

        public dynamic createoauthtoken()
        {
            int statuscode = 0;
            var _Encode64 = _publicKey + ":" + _secretkey;

            var data = "grant_type = client_credentials";
            var plainTextBytes = Encoding.UTF8.GetBytes(_Encode64);

            var encodedText = Convert.ToBase64String(plainTextBytes);

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(strForioEpicenterOauthUrl);
            req.KeepAlive = true;
            req.ProtocolVersion = HttpVersion.Version10;
            try
            {
                req.ContentType = "application/x-www-form-urlencoded";
                req.Method = "POST";
                req.Headers.Add("Authorization", "Basic " + encodedText);
                byte[] postBytes = Encoding.ASCII.GetBytes(data);
                req.ContentLength = postBytes.Length;
                Stream requestStream = req.GetRequestStream();
                requestStream.Write(postBytes, 0, postBytes.Length);
                requestStream.Close();
            }
            catch (Exception ex)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "EpicenterAPI:call()",
                " Error for HttpWebRequest method  " + ex.Message, ex, objGetConnection.strPurchaseModuleName);
            }

            try
            {
                HttpWebResponse response = (HttpWebResponse)req.GetResponse();
                statuscode = (int)response.StatusCode;
                Stream resStream = response.GetResponseStream();
                var sr = new StreamReader(response.GetResponseStream());
                responseText = sr.ReadToEnd();


                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "EpicenterAPI:call()",
                "responseText : " + responseText, null, objGetConnection.strPurchaseModuleName);
            }
            catch (WebException e)
            {
                HttpWebResponse response = (HttpWebResponse)e.Response;
                Console.Write("Errorcode: {0}", (int)response.StatusCode);

                statuscode = (int)response.StatusCode;

                Stream resStream = response.GetResponseStream();

                var sr = new StreamReader(response.GetResponseStream());
                responseText = sr.ReadToEnd();

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "EpicenterAPI:call()",
                "Inside Web Exception : responseText " + responseText, null, objGetConnection.strPurchaseModuleName);
            }

            dynamic responseObject = (new JavaScriptSerializer()).DeserializeObject(responseText);
            return responseObject;
        }        

        public dynamic CreateGroup(string Token, string strGroupName, string strEpicenterAcoount, string strForioName, string maxUserCount)
        {
            var creategroupdetails = new
            {
                name = strGroupName,
                account = strEpicenterAcoount,
                project = strForioName,
                maxUsers = maxUserCount
            };
            
            string creategroupdetailsjson = new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(creategroupdetails);

            return post("group", creategroupdetailsjson, Token, strForioEpicenterGroupUrl);
        }

        public dynamic CreateUser(string Token, string strEmailId, string strEpicenterAccount, string strPassword, string strFirstName, string strLastName)
        {
            var createuser = new
            {
                userName = strEmailId,
                account = strEpicenterAccount,
                password = strPassword,
                firstName = strFirstName,
                lastName = strLastName
            };

            string createuserjson = new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(createuser);

            return post("user", createuserjson, Token, strForioEpicenterUserUrl);
        }

        public dynamic PasswordRecoveryAPI(string strEmailId, string strEpicenterAccount, string strProjectFullName, string simulationUrl)
        {
            var passrecovery = new
            {
                userName = strEmailId,
                account = strEpicenterAccount,
                projectFullName = strProjectFullName,
                redirectUrl = simulationUrl
            };
            
            string passrecoveryjson = new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(passrecovery);

            return post("password", passrecoveryjson, "", strForioEpicenterPassRecoveryUrl);
        }

        public dynamic AddMemberToGroup(string Token, string strUserId, string strGroupId, string strUserRole)
        {
            strForioEpicenterAddMemberUrl = strForioEpicenterAddMemberUrl + strGroupId;
            var addmember = new
            {
                userId = strUserId,
                role = strUserRole
            };
                       
            string Addmemberjson = new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(addmember);

            return post("addmember", Addmemberjson, Token, strForioEpicenterAddMemberUrl);
        }

        public dynamic SearchEndUser(string Token, string strEpicenterAccount, string strEmailId)
        {
            string strForioEpicenterSearchEndUser = strForioEpicenterUserUrl + "?" + "account=" + strEpicenterAccount + "&userName=" + strEmailId;

            return get("searchuser", "", Token, strForioEpicenterSearchEndUser);
        }

        protected dynamic get(string method, string data, string append, string baseUrl)
        {
            return call("GET", method, data, append, baseUrl);
        }

        protected dynamic post(string method, string postdata, string append, string baseUrl)
        {
            return call("POST", method, postdata, append, baseUrl);
        }

        public dynamic call(string callMethod, string apiMethod, string data, string append, string baseUrl)
        {
            dynamic responseObject = null;
            int statuscode = 0;

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(baseUrl);
            req.KeepAlive = true;
            req.ProtocolVersion = HttpVersion.Version10;

            try
            {
                req.ContentType = "application/json";
                req.Method = callMethod;
                if (!string.IsNullOrEmpty(append))
                {
                    req.Headers.Add("Authorization", "Bearer " + append);
                }

                if (!string.IsNullOrEmpty(data))
                {
                    byte[] postBytes = Encoding.ASCII.GetBytes(data);
                    req.ContentLength = postBytes.Length;
                    Stream requestStream = req.GetRequestStream();
                    requestStream.Write(postBytes, 0, postBytes.Length);
                    requestStream.Close();
                }
            }
            catch (Exception ex)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "EpicenterAPI:call()",
                " Error for HttpwebRequest method " + ex.Message, ex, objGetConnection.strPurchaseModuleName);
            }

            try
            {
                HttpWebResponse response = (HttpWebResponse)req.GetResponse();
                statuscode = (int)response.StatusCode;
                Stream resStream = response.GetResponseStream();
                var sr = new StreamReader(response.GetResponseStream());
                responseText = sr.ReadToEnd();

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "EpicenterAPI:call()",
                "responseText" + responseText, null, objGetConnection.strPurchaseModuleName);
            }
            catch (WebException e)
            {
                HttpWebResponse response = (HttpWebResponse)e.Response;
                Console.Write("Errorcode: {0}", (int)response.StatusCode);
                statuscode = (int)response.StatusCode;
                Stream resStream = response.GetResponseStream();
                var sr = new StreamReader(response.GetResponseStream());
                responseText = sr.ReadToEnd();

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "EpicenterAPI:call()",
               "Inside Web Exception :responseText" + responseText, null, objGetConnection.strPurchaseModuleName);
            }
            if (callMethod == "GET"){ responseObject = (new JavaScriptSerializer()).DeserializeObject(responseText); }  
            if (callMethod == "POST"){responseObject = (new JavaScriptSerializer()).DeserializeObject(responseText); }

            return responseObject;
        }
    }
}
