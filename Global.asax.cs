using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using Mail_WebArchiveView.Models;
using log4net;

namespace Mail_WebArchiveView
{
    // Note: For instructions on enabling IIS6 or IIS7 classic mode, 
    // visit http://go.microsoft.com/?LinkId=9394801

    public class MvcApplication : System.Web.HttpApplication
    {
        
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();

            WebApiConfig.Register(GlobalConfiguration.Configuration);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
            AuthConfig.RegisterAuth();
            ValueProviderFactories.Factories.Add(new JsonValueProviderFactory());  
            log4net.Config.XmlConfigurator.Configure();

        }

        void Session_Start(Object sender, EventArgs e)
        {
            Search search = null;
            HttpContext.Current.Session.Add("search", search);
        }
    }
}