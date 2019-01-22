using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Excel_Import_Customization.Startup))]
namespace Excel_Import_Customization
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
