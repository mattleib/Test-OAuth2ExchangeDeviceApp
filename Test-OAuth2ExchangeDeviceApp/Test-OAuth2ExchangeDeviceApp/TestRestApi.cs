//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
//
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Test_OAuth2ExchangeDeviceApp
{
    class TestRestApi
    {
        private string endpoint;
        private string token;
        private bool verbose;

        public TestRestApi(
            string restEndpoint, 
            string accessToken,
            bool verbose)
        {
            this.endpoint = restEndpoint;
            this.token = accessToken;
            this.verbose = verbose;
        }

        public void Start()
        {
            try
            {
                this.DoWork();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private void DoWork()
        {
            string meApi = this.endpoint + "/me"; // users('smtpaddress')
            //string meApi = this.endpoint + "/users(" + mailbox + ")/folders/inbox/messages";

            Func<HttpRequestMessage> requestCreator = () =>
            {
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, meApi);
                request.Headers.TryAddWithoutValidation("Content-Type", "application/json");
                return request;
            };

            Console.WriteLine("DOING: Rest API call");
            Console.WriteLine("");

            string json = HttpRequestHelper.MakeHttpRequest(requestCreator, this.token, Constants.UserAgent);
            dynamic parsedJson = JsonConvert.DeserializeObject(json);
            string formattedJson = JsonConvert.SerializeObject(parsedJson, Formatting.Indented);

            Console.WriteLine(formattedJson);
            Console.WriteLine("");
            
            Console.WriteLine("DONE: Rest API call");
            Console.WriteLine("");
        }
    }
}
// MIT License: 

// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 

// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.