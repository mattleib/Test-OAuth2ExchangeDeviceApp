//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_OAuth2ExchangeDeviceApp
{
    public class Constants
    {
        public const string OfficeClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c";
        public const string DefaultInstalledClientRedirectUri = "urn:ietf:wg:oauth:2.0:oob";

        public const string AuthorityPPE = "https://login.windows-ppe.net/common/authorize";
        public const string AuthorityProd = "https://login.windows.net/common/authorize";

        public const string ResourceProd = "https://outlook.office365.com";
        public const string AutodiscoverEndpointProd = Constants.ResourceProd + "/autodiscover/autodiscover.svc";
        public const string EWSApiEndpointProd = Constants.ResourceProd + "/ews/exchange.asmx";
        public const string RestApiEndpointProd = Constants.ResourceProd + "/api/v1.0";

        public const string ResourcePPE = "https://sdfpilot.outlook.com";
        public const string AutodiscoverEndpointPPE = Constants.ResourcePPE + "/autodiscover/autodiscover.svc";
        public const string EWSApiEndpointPPE = Constants.ResourcePPE + "/ews/exchange.asmx";
        public const string RestApiEndpointPPE = Constants.ResourcePPE + "/api/v1.0";

        public const string ToBeDiscoverd = "{..will be discoverd..}";

        public const string UserAgent = "Test-OAuth2ExchangeDeviceApp/1.0";
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