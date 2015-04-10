//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
//
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Test_OAuth2ExchangeDeviceApp
{
    class Program
    {
        static void Main(string[] args)
        {
            CommandLineArgs commandLineArgs = new CommandLineArgs();

            if (!CommandLine.Parser.Default.ParseArguments(args, commandLineArgs))
                return;

            if (!Helpers.IsValidSmtp(commandLineArgs.MailboxSmtpAddress))
            {
                Console.WriteLine("ERROR: Invalid SMTP address.");
                return;
            }

            var t = new Thread(Run);
            t.SetApartmentState(ApartmentState.STA);
            t.Start(commandLineArgs);
            t.Join();
        }

        static void Run(object obj)
        {
            CommandLineArgs commandLineArgs = obj as CommandLineArgs;

            if (commandLineArgs.OverrideUsePPE)
            {
                commandLineArgs.RestApiEndpointAddress = Constants.RestApiEndpointPPE;
                commandLineArgs.AutodiscoverEndpointAddress = Constants.AutodiscoverEndpointPPE;
                commandLineArgs.EwsApiEndpointAddress = Constants.EWSApiEndpointPPE;
                commandLineArgs.Authority = Constants.AuthorityPPE;
            }

            if (commandLineArgs.DiscoverEndpointsFromSmtpDomain)
            { 
                string domain = commandLineArgs.MailboxSmtpAddress.Split('@')[1];

                commandLineArgs.AutodiscoverEndpointAddress = string.Format("https://autodiscover.{0}/autodiscover/autodiscover.svc", domain);
                commandLineArgs.EwsApiEndpointAddress = Constants.ToBeDiscoverd;
                commandLineArgs.RestApiEndpointAddress = Constants.ToBeDiscoverd;
            }

            Console.WriteLine("");
            Console.WriteLine("Starting Tests with following parameters:");
            Console.WriteLine("=========================================");
            Console.WriteLine("MailboxSmtpAddress: " + commandLineArgs.MailboxSmtpAddress);
            Console.WriteLine("DiscoverEndpointsFromSmtpDomain: " + commandLineArgs.DiscoverEndpointsFromSmtpDomain.ToString());
            Console.WriteLine("ClientId: " + commandLineArgs.ClientId);
            Console.WriteLine("ClientRedirectUri: " + commandLineArgs.ClientRedirectUri);
            Console.WriteLine("Authority: " + commandLineArgs.Authority);
            Console.WriteLine("AutodiscoverEndpointAddress: " + commandLineArgs.AutodiscoverEndpointAddress);
            Console.WriteLine("EwsApiEndpointAddress: " + commandLineArgs.EwsApiEndpointAddress);
            Console.WriteLine("RestApiEndpointAddress: " + commandLineArgs.RestApiEndpointAddress);
            Console.WriteLine("Verbose: " + commandLineArgs.Verbose.ToString());
            Console.WriteLine("=========================================");
            Console.WriteLine("");


            // Initializing the Authorization context
            AuthenticationContext authenticationContext = new AuthenticationContext(commandLineArgs.Authority, false);

            //-------------------------------------------------------
            //--- Testing AutoDiscover 
            //-------------------------------------------------------
            Helpers.WriteStep(string.Format("Step1: Autodiscover mailbox [{0}] against endpoint [{1}]",
                    commandLineArgs.MailboxSmtpAddress,    
                    commandLineArgs.AutodiscoverEndpointAddress
                ));

            string accessToken = null;

            accessToken = Helpers.GetAccessToken(
                authenticationContext,
                commandLineArgs.AutodiscoverEndpointAddress,
                commandLineArgs.ClientId,
                commandLineArgs.ClientRedirectUri,
                commandLineArgs.Verbose
                );

            TestAutodiscover autodiscover = new TestAutodiscover(
                commandLineArgs.MailboxSmtpAddress,
                commandLineArgs.AutodiscoverEndpointAddress,
                accessToken,
                commandLineArgs.Verbose
                );

            autodiscover.Start();

            if (commandLineArgs.DiscoverEndpointsFromSmtpDomain)
            {
                if (autodiscover.LastAutodiscoverResponse == null)
                {
                    Console.WriteLine("ERROR: Autodiscover could not discover the EWS or Rest API endpoint for mailbox [{0}].",
                        commandLineArgs.MailboxSmtpAddress);
                    Console.WriteLine("");

                    return;
                }

                string endpoint = autodiscover.LastAutodiscoverResponse.Settings[UserSettingName.ExternalEwsUrl] as string;
                Uri endpointUri = new Uri(endpoint);

                commandLineArgs.EwsApiEndpointAddress = endpoint;
                commandLineArgs.RestApiEndpointAddress = endpointUri.Scheme + "://" + endpointUri.Authority + "/api/v1.0";

                Console.WriteLine("Using following discovered API endpoints:");
                Console.WriteLine("=========================================");
                Console.WriteLine("EwsApiEndpointAddress: " + commandLineArgs.EwsApiEndpointAddress);
                Console.WriteLine("RestApiEndpointAddress: " + commandLineArgs.RestApiEndpointAddress);
                Console.WriteLine("");
            }

            Helpers.WriteStep("Step1: complete.");

            //-------------------------------------------------------
            //--- Testing EWS Soap 
            //-------------------------------------------------------
            Helpers.WriteStep(string.Format("Step2: EWS Soap API call for mailbox [{0}] against endpoint [{1}]",
                    commandLineArgs.MailboxSmtpAddress,
                    commandLineArgs.EwsApiEndpointAddress
            ));

            accessToken = Helpers.GetAccessToken(
                authenticationContext,
                commandLineArgs.EwsApiEndpointAddress,
                commandLineArgs.ClientId,
                commandLineArgs.ClientRedirectUri,
                commandLineArgs.Verbose
                );


            TestEwsSoap ews = new TestEwsSoap(
                commandLineArgs.EwsApiEndpointAddress,
                accessToken,
                commandLineArgs.Verbose
                );

            ews.Start();

            Helpers.WriteStep("Step2: complete.");

            //-------------------------------------------------------
            //--- Testing Rest APIs 
            //-------------------------------------------------------


            Helpers.WriteStep(string.Format("Step3: Rest API call for mailbox [{0}] against endpoint [{1}]",
                commandLineArgs.MailboxSmtpAddress,
                commandLineArgs.RestApiEndpointAddress
            ));

            accessToken = Helpers.GetAccessToken(
                authenticationContext,
                commandLineArgs.RestApiEndpointAddress,
                commandLineArgs.ClientId,
                commandLineArgs.ClientRedirectUri,
                commandLineArgs.Verbose
                );

            TestRestApi restApi = new TestRestApi(
                commandLineArgs.RestApiEndpointAddress,
                accessToken,
                commandLineArgs.Verbose
                );

            restApi.Start();

            Helpers.WriteStep("Step3: complete.");

            //--- Cleaning up...
            authenticationContext.TokenCache.Clear();
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