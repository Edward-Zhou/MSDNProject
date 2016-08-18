using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Office365OAuth.Startup))]
namespace Office365OAuth
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
