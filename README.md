# Test-OAuth2ExchangeDeviceApp

Example: Test-OAuth2ExchangeDeviceApp.exe

Test-OAuth2ExchangeDeviceApp 1.0.0.0
Copyright c  2015

  -m, --MailboxSmtpAddress                 Required. SMTP address of mailbox to
                                           access.

  -s, --DiscoverEndpointsFromSmtpDomain    (Default: False) Use SMTP address to
                                           autodiscover endpoints.

  -c, --ClientId                           (Default:
                                           d3590ed6-52b3-4102-aeff-aad2292ab01c)
                                           ClientId of the application as
                                           registered in AAD.

  -u, --ClientRedirectUri                  (Default: urn:ietf:wg:oauth:2.0:oob)
                                           Client reply address of the
                                           application as registered in AAD.
                                           Usually for installed applications
                                           'urn:ietf:wg:oauth:2.0:oob'.

  -a, --Authority                          (Default:
                                           https://login.windows.net/common/auth
                                           orize) Authority address of the
                                           Authorization Server.

  -d, --AutodiscoverEndpointAddress        (Default:
                                           https://outlook.office365.com/autodis
                                           cover/autodiscover.svc) Autodiscover
                                           endpoint address of the Exchange
                                           Service.

  -e, --EwsApiEndpointAddress              (Default:
                                           https://outlook.office365.com/ews/exc
                                           hange.asmx) EWS Soap endpoint
                                           address of the Exchange Service.

  -r, --RestApiEndpointAddress             (Default:
                                           https://outlook.office365.com/api/v1.
                                           0) Rest API endpoint address of the
                                           Exchange Service.

  -o, --OverrideUsePPE                     (Default: False) Override all
                                           commandline parameters and use PPE
                                           defaults.

  -v, --Verbose                            (Default: False) Print details.

  --help                                   Display this help screen.
