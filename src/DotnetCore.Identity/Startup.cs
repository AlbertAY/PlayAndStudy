using IdentityServer4;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DotnetCore.Identity
{
    public class Startup
    {
        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940


        private readonly LoggerFactory _loggerFactory;


        public  Startup()
        {
           // _loggerFactory = loggerFactory;
        }


        public void ConfigureServices(IServiceCollection services)
        {
            services.AddIdentityServer()
                            .AddDeveloperSigningCredential()
            //配置身份资源
            .AddInMemoryIdentityResources(Config.GetIdentityResources())
            //配置API资源
            .AddInMemoryApiResources(Config.GetApiResources())
            //预置允许验证的Client
            .AddInMemoryClients(Config.GetClients())
            .AddTestUsers(Config.GetUsers());
            services.AddAuthentication()
                //添加Google第三方身份认证服务（按需添加）
                //.AddGoogle("Google", options =>
                //{
                //    options.SignInScheme = IdentityServerConstants.ExternalCookieAuthenticationScheme;
                //    options.ClientId = "434483408261-55tc8n0cs4ff1fe21ea8df2o443v2iuc.apps.googleusercontent.com";
                //    options.ClientSecret = "3gcoTrEDPPJ0ukn_aYYT6PWo";
                //})
                //如果当前IdentityServer不提供身份认证服务，还可以添加其他身份认证服                务提供商
                .AddOpenIdConnect("oidc", "OpenID Connect", options =>
                {
                    options.SignInScheme = IdentityServerConstants.ExternalCookieAuthenticationScheme;
                    options.SignOutScheme = IdentityServerConstants.SignoutScheme;
                    options.Authority = "https://demo.identityserver.io/";
                    options.ClientId = "implicit";
                    options.TokenValidationParameters = new TokenValidationParameters
                    {
                        NameClaimType = "name",
                        RoleClaimType = "role"
                    };
                });
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {

            //_loggerFactory.a(LogLevel.Debug);


            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseIdentityServer();

            app.UseRouting();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapGet("/", async context =>
                {
                    await context.Response.WriteAsync("Hello World!");
                });
            });
        }
    }
}
