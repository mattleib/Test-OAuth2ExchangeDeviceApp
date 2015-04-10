//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
//
using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_OAuth2ExchangeDeviceApp
{
    class CommandLineArgs
    {
        [Option('m', "MailboxSmtpAddress", Required = true,
            HelpText = "SMTP address of mailbox to access.")]
        public string MailboxSmtpAddress { get; set; }

        [Option('s', "DiscoverEndpointsFromSmtpDomain", DefaultValue = false,
            HelpText = "Use SMTP address to autodiscover endpoints.")]
        public bool DiscoverEndpointsFromSmtpDomain { get; set; }

        [Option('c', "ClientId", DefaultValue = Constants.OfficeClientId,
            HelpText = "ClientId of the application as registered in AAD.")]
        public string ClientId { get; set; }

        [Option('u', "ClientRedirectUri", DefaultValue = Constants.DefaultInstalledClientRedirectUri,
            HelpText = "Client reply address of the application as registered in AAD. Usually for installed applications 'urn:ietf:wg:oauth:2.0:oob'.")]
        public string ClientRedirectUri { get; set; }

        [Option('a', "Authority", DefaultValue = Constants.AuthorityProd,
            HelpText = "Authority address of the Authorization Server.")]
        public string Authority { get; set; }

        [Option('d', "AutodiscoverEndpointAddress", DefaultValue = Constants.AutodiscoverEndpointProd,
            HelpText = "Autodiscover endpoint address of the Exchange Service.")]
        public string AutodiscoverEndpointAddress { get; set; }

        [Option('e', "EwsApiEndpointAddress", DefaultValue = Constants.EWSApiEndpointProd,
            HelpText = "EWS Soap endpoint address of the Exchange Service.")]
        public string EwsApiEndpointAddress { get; set; }

        [Option('r', "RestApiEndpointAddress", DefaultValue = Constants.RestApiEndpointProd,
            HelpText = "Rest API endpoint address of the Exchange Service.")]
        public string RestApiEndpointAddress { get; set; }

        [Option('o', "OverrideUsePPE", DefaultValue = false,
            HelpText = "Override all commandline parameters and use PPE defaults.")]
        public bool OverrideUsePPE { get; set; }

        [Option('v', "Verbose", DefaultValue = false,
            HelpText = "Print details.")]
        public bool Verbose { get; set; }

        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
                (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
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
