//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
//
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_OAuth2ExchangeDeviceApp
{
    // From: Jason Johnston@https://github.com/jasonjoh/office365-azure-guides/blob/master/code/parse-token.cs
    public static class Base64UrlEncoder
    {
        static char Base64PadCharacter = '=';
        static string DoubleBase64PadCharacter = String.Format(CultureInfo.InvariantCulture, "{0}{0}", Base64PadCharacter);
        static char Base64Character62 = '+';
        static char Base64Character63 = '/';
        static char Base64UrlCharacter62 = '-';
        static char Base64UrlCharacter63 = '_';

        public static byte[] DecodeBytes(string arg)
        {
            string s = arg;
            s = s.Replace(Base64UrlCharacter62, Base64Character62); // 62nd char of encoding
            s = s.Replace(Base64UrlCharacter63, Base64Character63); // 63rd char of encoding
            switch (s.Length % 4) // Pad 
            {
                case 0:
                    break; // No pad chars in this case
                case 2:
                    s += DoubleBase64PadCharacter; break; // Two pad chars
                case 3:
                    s += Base64PadCharacter; break; // One pad char
                default:
                    throw new ArgumentException("Illegal base64url string!", arg);
            }
            return Convert.FromBase64String(s); // Standard base64 decoder
        }

        public static string Decode(string arg)
        {
            return Encoding.UTF8.GetString(DecodeBytes(arg));
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