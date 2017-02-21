using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    /// <summary>
    /// Static class full of helper methods to make HTTP requests
    /// </summary>
    public static class HttpHelper
    {
        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <returns>The String value of the result</returns>
        public async static Task<String> MakeGetRequestForStringAsync(String requestUrl,
            String accessToken = null)
        {
            HttpResponseHeaders responseHeaders = null;
            return (await MakeHttpRequestAsync<String>("GET",
                requestUrl,
                responseHeaders,
                accessToken: accessToken,
                resultPredicate: async (r) => await r.Content.ReadAsStringAsync()));
        }

        /// <summary>
        /// This helper method makes an HTTP HEAD request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <returns>The String value of the result</returns>
        public async static Task<String> MakeHeadRequestAsync(String requestUrl,
            String accessToken = null)
        {
            HttpResponseHeaders responseHeaders = null;
            return (await MakeHttpRequestAsync<String>("HEAD",
                requestUrl,
                responseHeaders,
                accessToken: accessToken));
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The accept header for the response</param>
        /// <returns>The Stream  of the result</returns>
        public async static Task<System.IO.Stream> MakeGetRequestForStreamAsync(String requestUrl,
            String accept,
            String accessToken = null,
            String referer = null)
        {
            HttpResponseHeaders responseHeaders = null;
            return (await MakeHttpRequestAsync<System.IO.Stream>("GET",
                requestUrl,
                responseHeaders,
                accessToken: accessToken,
                referer: referer,
                resultPredicate: async (r) => await r.Content.ReadAsStreamAsync()));
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="responseHeaders">The response headers of the HTTP request (output argument)</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The accept header for the response</param>
        /// <returns>The Stream  of the result</returns>
        public async static Task<System.IO.Stream> MakeGetRequestForStreamWithResponseHeadersAsync(String requestUrl,
            String accept,
            HttpResponseHeaders responseHeaders,
            String accessToken = null)
        {
            return (await MakeHttpRequestAsync<System.IO.Stream>("GET",
                requestUrl,
                responseHeaders,
                accessToken,
                resultPredicate: async (r) => await r.Content.ReadAsStreamAsync()));
        }

        /// <summary>
        /// This helper method makes an HTTP POST request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        public async static Task MakePostRequestAsync(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            HttpResponseHeaders responseHeaders = null;
            await MakeHttpRequestAsync<String>("POST",
                requestUrl,
                responseHeaders,
                accessToken: accessToken,
                content: content,
                contentType: contentType);
        }

        /// <summary>
        /// This helper method makes an HTTP POST request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <returns>The String value of the result</returns>
        public async static Task<String> MakePostRequestForStringAsync(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            HttpResponseHeaders responseHeaders = null;
            return (await MakeHttpRequestAsync<String>("POST",
                requestUrl,
                responseHeaders,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: async (r) => await r.Content.ReadAsStringAsync()));
        }

        /// <summary>
        /// This helper method makes an HTTP PUT request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        public async static Task MakePutRequestAsync(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            HttpResponseHeaders responseHeaders = null;
            await MakeHttpRequestAsync<String>("PUT",
                requestUrl,
                responseHeaders,
                accessToken: accessToken,
                content: content,
                contentType: contentType);
        }

        /// <summary>
        /// This helper method makes an HTTP PUT request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <returns>The String value of the result</returns>
        public async static Task<String> MakePutRequestForStringAsync(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            HttpResponseHeaders responseHeaders = null;
            return (await MakeHttpRequestAsync<String>("PUT",
                requestUrl,
                responseHeaders,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: async (r) => await r.Content.ReadAsStringAsync()));
        }

        /// <summary>
        /// This helper method makes an HTTP PATCH request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <returns>The String value of the result</returns>
        public async static Task<String> MakePatchRequestForStringAsync(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            HttpResponseHeaders responseHeaders = null;
            return (await MakeHttpRequestAsync<String>("PATCH",
                requestUrl,
                responseHeaders,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: async (r) => await r.Content.ReadAsStringAsync()));
        }

        /// <summary>
        /// This helper method makes an HTTP DELETE request
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <returns>The String value of the result</returns>
        public async static Task MakeDeleteRequestAsync(String requestUrl,
            String accessToken = null)
        {
            HttpResponseHeaders responseHeaders = null;
            await MakeHttpRequestAsync<String>("DELETE", requestUrl, responseHeaders, accessToken);
        }

        /// <summary>
        /// This helper method makes an HTTP request and eventually returns a result
        /// </summary>
        /// <param name="httpMethod">The HTTP method for the request</param>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The content type of the accepted response</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content  type of the request</param>
        /// <param name="resultPredicate">The predicate to retrieve the result, if any</param>
        /// <typeparam name="TResult">The type of the result, if any</typeparam>
        /// <returns>The value of the result, if any</returns>
        private async static Task<TResult> MakeHttpRequestAsync<TResult>(
            String httpMethod,
            String requestUrl,
            String accessToken = null,
            String accept = null,
            Object content = null,
            String contentType = null,
            String referer = null,
            Func<HttpResponseMessage, Task<TResult>> resultPredicate = null)
        {
            HttpResponseHeaders responseHeaders = null;
            return (await MakeHttpRequestAsync<TResult>(httpMethod,
                requestUrl,
                responseHeaders,
                accessToken,
                accept,
                content,
                contentType,
                referer,
                resultPredicate));
        }

        /// <summary>
        /// This helper method makes an HTTP request and eventually returns a result
        /// </summary>
        /// <param name="httpMethod">The HTTP method for the request</param>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="responseHeaders">The response headers of the HTTP request (output argument)</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The content type of the accepted response</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content  type of the request</param>
        /// <param name="resultPredicate">The predicate to retrieve the result, if any</param>
        /// <typeparam name="TResult">The type of the result, if any</typeparam>
        /// <returns>The value of the result, if any</returns>
        private async static Task<TResult> MakeHttpRequestAsync<TResult>(
            String httpMethod,
            String requestUrl,
            HttpResponseHeaders responseHeaders = null,
            String accessToken = null,
            String accept = null,
            Object content = null,
            String contentType = null,
            String referer = null,
            Func<HttpResponseMessage, Task<TResult>> resultPredicate = null)
        {
            // Prepare the variable to hold the result, if any
            TResult result = default(TResult);
            responseHeaders = null;

            Uri requestUri = new Uri(requestUrl);

            // If we have the token, then handle the HTTP request
            HttpClientHandler handler = new HttpClientHandler();
            handler.AllowAutoRedirect = true;
            HttpClient httpClient = new HttpClient(handler, true);

            // Set the Authorization Bearer token
            if (!String.IsNullOrEmpty(accessToken))
            {
                httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", accessToken);
            }

            if (!String.IsNullOrEmpty(referer))
            {
                httpClient.DefaultRequestHeaders.Referrer = new Uri(referer);
            }

            // If there is an accept argument, set the corresponding HTTP header
            if (!String.IsNullOrEmpty(accept))
            {
                httpClient.DefaultRequestHeaders.Accept.Clear();
                httpClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue(accept));
            }

            // Prepare the content of the request, if any
            HttpContent requestContent = null;
            System.IO.Stream streamContent = content as System.IO.Stream;
            if (streamContent != null)
            {
                requestContent = new StreamContent(streamContent);
                requestContent.Headers.ContentType = new MediaTypeHeaderValue(contentType);
            }
            else
            {
                requestContent =
                    (content != null) ?
                    new StringContent(JsonConvert.SerializeObject(content,
                        Formatting.None,
                        new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            ContractResolver = new CamelCasePropertyNamesContractResolver(),
                        }),
                    Encoding.UTF8, contentType) :
                    null;
            }

            // Prepare the HTTP request message with the proper HTTP method
            HttpRequestMessage request = new HttpRequestMessage(
                new HttpMethod(httpMethod), requestUrl);

            // Set the request content, if any
            if (requestContent != null)
            {
                request.Content = requestContent;
            }

            // Fire the HTTP request
            HttpResponseMessage response = await httpClient.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                // If the response is Success and there is a
                // predicate to retrieve the result, invoke it
                if (resultPredicate != null)
                {
                    result = await resultPredicate(response);
                }

                // Get any response header and put it in the answer
                responseHeaders = response.Headers;
            }
            else
            {
                throw new ApplicationException(
                    String.Format("Exception while invoking endpoint {0}.", requestUrl),
                    new HttpException(
                        (Int32)response.StatusCode,
                        response.Content.ReadAsStringAsync().Result));
            }

            return (result);
        }
    }
}
