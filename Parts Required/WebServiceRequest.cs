using System;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace Parts_Required
{
    public class WebServiceRequest : WebClient
    {
        public int Timeout { get; set; }
        public string Method { get; set; }
        public string _userName;
        public string _password;
        RightNowConnectService _rnConnectService;
        public WebServiceRequest(string method)
        {
            Timeout = 900000;
            Method = method;
            _rnConnectService = RightNowConnectService.GetService();
            string webServiceLoginValue = _rnConnectService.GetConfigValue("CUSTOM_CFG_EBS_WS_LOGIN");
            if (webServiceLoginValue != null)
            {

                var s = new JavaScriptSerializer();

                var loginValue = s.Deserialize<WebServiceLoginCredential>(webServiceLoginValue);
                _userName = loginValue.UserName;
                _password = loginValue.Password;
            }
        }
        /// <summary>
        /// Override GetWebRequest function of webclient to inculde get/set handshake session
        /// Webclient always call this function internally before making any web request
        /// </summary>
        /// <param name="address">URL</param>
        /// <returns>WebRequest instance</returns>
        protected override WebRequest GetWebRequest(Uri address)
        {
            HttpWebRequest request = base.GetWebRequest(address) as HttpWebRequest;
            request.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
            request.Method = Method;
            request.Timeout = Timeout;
            request.KeepAlive = false;
            request.CookieContainer = new CookieContainer();
            request.PreAuthenticate = true;
            NetworkCredential networkCredential = new NetworkCredential(_userName,_password);
            request.Credentials = networkCredential;
            return request;
        }
        /// <summary>
        /// Get the response from EBS webservice
        /// </summary>
        /// <param name="url">url where need to post data</param>
        /// <param name="content">data in jSon string format</param>
        /// <param name="lstHeaders">http header required to make request to EBS</param>
        /// <param name="Method">Type of request ..POST in this case</param>
        /// <returns>Response in jSon string</returns>
        public static string Get(string url, string content, string Method)
        {
            string strResponse = "";
            try
            {
                using (var request = new WebServiceRequest(Method))
                {
                    request.Encoding = Encoding.UTF8;
                    request.Headers.Add("Accept", "application/json");
                    request.Headers.Add("Content-Type", "application/json");
                    // request.Headers.Add("connection", "close");
                    if (content == "")
                    {
                        strResponse = request.DownloadString(url);
                    }
                    else
                    {
                        strResponse = request.UploadString(url, content);
                    }
                }
            }
            catch (WebException w)
            {
              
                System.Windows.Forms.MessageBox.Show(w.Message);

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
            return strResponse;
        }

        /// <summary>
        /// DeSerialize rest response string to dynamic object
        /// </summary>
        /// <param name="jsonText">string to deserialize</param>
        /// <returns>deserialized object</returns>
        public static object JsonDeSerialize(string jsonText)
        {
            var jss = new JavaScriptSerializer();
            var json = jss.Deserialize<dynamic>(jsonText);

            return json;
        }

        /// <summary>
        /// Serialize object into json string
        /// </summary>
        /// <param name="obj">object to serialize</param>
        /// <returns>serialized json string</returns>
        public static string JsonSerialize(object obj)
        {
            var json = new JavaScriptSerializer().Serialize(obj);
            return json;
        }
    }
}
