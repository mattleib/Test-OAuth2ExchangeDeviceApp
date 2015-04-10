//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
//
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_OAuth2ExchangeDeviceApp
{
    public class TestAutodiscover
    {
        private string mailbox;
        private string endpoint;
        private string token;
        private bool verbose;
        private GetUserSettingsResponse lastAutodiscoverResponse;
        public GetUserSettingsResponse LastAutodiscoverResponse {
            get { return lastAutodiscoverResponse; }
        }

        public TestAutodiscover(
            string mailboxSmtpAddress, 
            string autodiscoverEndpoint, 
            string accessToken,
            bool verbose)
        {
            this.mailbox = mailboxSmtpAddress;
            this.endpoint = autodiscoverEndpoint;
            this.token = accessToken;
            this.verbose = verbose;
            this.lastAutodiscoverResponse = null;
        }

        private void DoWork()
        {
            lastAutodiscoverResponse = null;

            AutodiscoverService autodiscoverService = new AutodiscoverService(ExchangeVersion.Exchange2013);

            if (this.verbose)
            {
                autodiscoverService.TraceEnabled = true;
                autodiscoverService.TraceFlags = TraceFlags.All;
            }

            autodiscoverService.Url = new Uri(this.endpoint);
            autodiscoverService.Credentials = new OAuthCredentials(this.token);
            autodiscoverService.UserAgent = Constants.UserAgent;
            autodiscoverService.ClientRequestId = Guid.NewGuid().ToString();
            autodiscoverService.ReturnClientRequestId = true;

            Console.WriteLine("Doing Autodiscover API call");
            Console.WriteLine("");

            lastAutodiscoverResponse = autodiscoverService.GetUserSettings(mailbox, UserSettingName.ExternalEwsUrl);

            Console.WriteLine("");
            Console.WriteLine("SUCCESS: Autodiscover API call");
            Console.WriteLine("");
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