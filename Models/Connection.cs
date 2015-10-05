using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net;
using System.Security.Cryptography;
using System.Text;

using log4net;
using log4net.Config;

using PagedList;
using EWS = Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using Mail_WebArchiveView.Models;

namespace Mail_WebArchiveView.Models
{
    public class Connection : Controller
    {

        public static ExchangeVersion ExVersion
        {
            get
            {
                string ev =  ConfigurationManager.AppSettings["ExchangeVersion"];

                switch (ev)
                {
                    case "Exchange2007_SP1":
                        return ExchangeVersion.Exchange2007_SP1;
                    case "Exchange2010":
                        return ExchangeVersion.Exchange2010;
                    case "Exchange2010_SP1":
                        return ExchangeVersion.Exchange2010_SP1;
                    case "Exchange2010_SP2":
                        return ExchangeVersion.Exchange2010_SP2;
                    case "Exchange2013":
                        return ExchangeVersion.Exchange2013;
                }
                return ExchangeVersion.Exchange2013;
            }
        }
        
        public static string ExUser
        {
            get
            {
                return ConfigurationManager.AppSettings["JournalUser"];
            }
        }

        public static string ExCred
        {
            get
            {
                return ConfigurationManager.AppSettings["JournalPassword"];
            }
        }

        public static string ExAccount
        {
            get
            {
                return ConfigurationManager.AppSettings["JournalAcct"];
            }
        }

        public static string ExAutoDiscover
        {
            get
            {
                return ConfigurationManager.AppSettings["AutoDiscoverUrl"];
            }
        }

        public static int ExPageSize
        {
            get
            {
                return Convert.ToInt32(ConfigurationManager.AppSettings["ExPageSize"]);
            }
        }

        public static int ExOffset
        {
            get
            {
                return Convert.ToInt32(ConfigurationManager.AppSettings["ExOffset"]);
            }
        }

        public static ExchangeService ConnectEWS()
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            ExchangeService service = new ExchangeService(Connection.ExVersion);


            if (!string.IsNullOrWhiteSpace(Connection.ExUser) && !string.IsNullOrWhiteSpace(Connection.ExCred))
            {
                service.Credentials = new NetworkCredential(Connection.ExUser, Connection.ExCred);
            }
            else
            {
                service.UseDefaultCredentials = true;
            }

            if (!string.IsNullOrWhiteSpace(Connection.ExAutoDiscover))
            {
                Uri manualAutoUrl = new Uri(Connection.ExAutoDiscover);
                service.Url = manualAutoUrl;
            }
            else
            {
                EWS.Autodiscover.AutodiscoverService autoDiscover = new EWS.Autodiscover.AutodiscoverService(Connection.ExVersion);
                EWS.Autodiscover.GetUserSettingsResponse response = autoDiscover.GetUserSettings(Connection.ExAccount, new EWS.Autodiscover.UserSettingName[] { EWS.Autodiscover.UserSettingName.InternalEwsUrl, EWS.Autodiscover.UserSettingName.UserDeploymentId });

                Uri url = new Uri(response.Settings[EWS.Autodiscover.UserSettingName.InternalEwsUrl].ToString());

                service.Url = url;
            }

            return service;

        }

        public static bool CertificateValidationCallBack(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }

    }

}
