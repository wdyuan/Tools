using System;
using System.IO;
using System.Net;
using System.Text;

namespace tang.cdt_ec_order
{
    /// <summary>
    /// HttpRequest辅助类
    /// </summary>
    public class HttpRequestHelper
    {
        /// <summary>
        /// 执行请求
        /// </summary>
        /// <param name="url">请求URL地址，如："http://www.123.com/GetUsers"</param>
        /// <param name="requestData">请求传递的参数数据,如："{\"id\":\"123\"}"</param>
        /// <param name="requestMethod">请求方法，如：GET,POST。默认POST。</param>
        /// <param name="encoding">字符编码方式，如：Encoding.UTF8。默认Encoding.UTF8</param>
        /// <returns>返回请求响应的字符串</returns>
        public static string ExecRequest(string url, string requestData, string requestMethod = null, string cookie = null, Encoding encoding = null)
        {
            if (encoding == null) encoding = Encoding.UTF8;
            if (string.IsNullOrEmpty(requestMethod)) requestMethod = "POST";

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = requestMethod;

            if (!string.IsNullOrWhiteSpace(cookie))
            {
                request.Headers.Add("Cookie", cookie);
            }

            try
            {
                if (requestMethod.ToUpper() == "POST")
                {
                    request.ContentType = "application/x-www-form-urlencoded";

                    if (!string.IsNullOrWhiteSpace(requestData))
                        request.ContentLength = encoding.GetByteCount(requestData);

                    if (requestData != null)
                    {
                        byte[] buffer = encoding.GetBytes(requestData);

                        request.ContentLength = buffer.Length;

                        using (Stream myRequestStream = request.GetRequestStream())
                        {
                            myRequestStream.Write(buffer, 0, buffer.Length);
                            myRequestStream.Close();
                        }
                    }
                }

                var response = (HttpWebResponse)request.GetResponse();

                string retString = string.Empty;

                using (Stream myResponseStream = response.GetResponseStream())
                {
                    if (myResponseStream != null)
                    {
                        using (var myStreamReader = new StreamReader(myResponseStream, encoding))
                        {
                            retString = myStreamReader.ReadToEnd();
                            myStreamReader.Close();
                            myResponseStream.Close();

                            return retString;
                        }
                    }
                }

                return retString;
            }
            catch (Exception exception)
            {
                return exception.Message;
            }
        }

        /// <summary>
        /// POST方式执行请求
        /// </summary>
        /// <param name="url">请求URL地址，如："http://www.123.com/GetUsers"</param>
        /// <param name="postData">请求传递的参数数据,如："{\"id\":\"123\"}"</param>
        /// <param name="cookie"></param>
        /// <param name="encoding">字符编码方式，如：Encoding.UTF8。默认Encoding.UTF8</param>
        /// <returns>返回请求响应的字符串</returns>
        public static string Post(string url, string postData, string cookie = null, Encoding encoding = null)
        {
            return ExecRequest(url, postData, "POST", cookie, encoding);
        }
    }
}
