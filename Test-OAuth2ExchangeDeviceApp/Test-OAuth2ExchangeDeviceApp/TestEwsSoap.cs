//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
//
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_OAuth2ExchangeDeviceApp
{
    class TestEwsSoap
    {
        private string endpoint;
        private string token;
        private bool verbose;

        public TestEwsSoap(
            string ewsEndpoint, 
            string accessToken,
            bool verbose)
        {
            this.endpoint = ewsEndpoint;
            this.token = accessToken;
            this.verbose = verbose;
        }

        private void DoWork()
        {
            ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2013);

            if (this.verbose)
            {
                exchangeService.TraceEnabled = true;
                exchangeService.TraceFlags = TraceFlags.All;
            }

            exchangeService.Url = new Uri(this.endpoint);
            exchangeService.Credentials = new OAuthCredentials(this.token);
            exchangeService.UserAgent = Constants.UserAgent;
            exchangeService.ClientRequestId = Guid.NewGuid().ToString();
            exchangeService.ReturnClientRequestId = true;
            //exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailbox);
            //exchangeService.HttpHeaders.Add("X-AnchorMailbox", mailbox);     

            Console.WriteLine("Doing EWS Soap API call");
            Console.WriteLine("");

            exchangeService.FindFolders(WellKnownFolderName.Inbox, new FolderView(10));

            Console.WriteLine("");
            Console.WriteLine("SUCCESS: EWS Soap API call");
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