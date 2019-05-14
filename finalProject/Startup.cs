using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(finalProject.Startup))]
namespace finalProject
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
