//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
//
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Net.Http;
using System.Web;

namespace Test_OAuth2ExchangeDeviceApp
{
    public static class Helpers
    {
        public static bool IsValidSmtp(string smtp)
        {
            if (string.IsNullOrEmpty(smtp))
                return false;

            // Copied from: http://www.regular-expressions.info/email.html
            string smtpRegexPattern = @"^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}$";
            Regex rgx = new Regex(smtpRegexPattern, RegexOptions.IgnoreCase);
            MatchCollection matches = rgx.Matches(smtp);
            if (matches.Count == 0)
                return false;

            return true;
        }

        private static List<string> SplitString(string value, int len)
        {
            List<string> r = new List<string>();
            if (value.Length <= len)
            {
                r.Add(value);
                return r;
            }

            while (value.Length >= len)
            {
                string s = value.Substring(0, len-1);
                r.Add(s);
                value = value.Substring(len);
            }
            if (value.Length > 0)
            { 
                r.Add(value); 
            }            

            return r;
        }

        public static void WriteStep(string stepInfo)
        {
            Console.WriteLine("");
            Console.WriteLine("********************************************************************************");
            List<string> stepList = Helpers.SplitString(stepInfo, 75);
            foreach(string s in stepList) 
            {
                Console.WriteLine(" " + s);
            }
            Console.WriteLine("********************************************************************************");
            Console.WriteLine("");
        }

        private static string Base64UrlDecodeJwtTokenPayload(string base64UrlEncodedJwtToken)
        {
            string payload = base64UrlEncodedJwtToken.Split('.')[1];

            return Base64UrlEncoder.Decode(payload);
        }

        public static string GetAccessToken(AuthenticationContext context, 
            string apiEndpoint, string clientId, string clientRedirectUri, bool verbose)
        {
            try
            {
                Uri endpoint = new Uri(apiEndpoint);
                string resource = endpoint.Scheme + "://" + endpoint.Authority + "/";
                Uri clientUri = new Uri(clientRedirectUri);

                Console.WriteLine("GetAccessToken for resource [{0}]", resource);
                Console.WriteLine("");

                AuthenticationResult result = context.AcquireToken(
                    resource,
                    clientId,
                    clientUri
                    );

                Console.WriteLine("SUCCESS: GetAccessToken for resource [{0}]", resource);
                Console.WriteLine("");

                if (verbose)
                {
                    Console.WriteLine("AccessToken: [{0}]", result.AccessToken);
                    Console.WriteLine("");

                    Console.WriteLine("Peek AccessToken: [http://jwt.calebb.net/#jwt={0}]", result.AccessToken);
                    Console.WriteLine("");

                    string tokenJsonPayload = Base64UrlDecodeJwtTokenPayload(result.AccessToken);
                    dynamic parsedJson = JsonConvert.DeserializeObject(tokenJsonPayload);
                    string formattedJson = JsonConvert.SerializeObject(parsedJson, Formatting.Indented);
                    Console.WriteLine("AccessToken Payload:");
                    Console.WriteLine("====================");
                    Console.WriteLine(formattedJson);
                    Console.WriteLine("");
                }

                return result.AccessToken;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return String.Empty;
            }

        } // GetAccessToken

    } // class Helpers
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