using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.OpenApi.Models;
using SharePointCamlQuery.Models;
using System.Net;
using System.Net.Http;

namespace SharePointCamlQuery
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();

            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo { Title = "Sharepoint Documents API", Version = "v1" });
            });
            services.AddHttpClient("SharepointClient", c =>
            {
                c.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
            }).ConfigurePrimaryHttpMessageHandler(c =>
            {
                var settings = Configuration.GetSection($"SharePointSettings").Get<SharePointSettings>();
                return new HttpClientHandler { Credentials = new NetworkCredential(settings.SPUserName, settings.SPPassword, settings.SPDomain) };
            });

        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseHttpsRedirection();

            app.UseRouting();

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
            app.UseSwagger();

            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "Sharepoint Documents API V1");
            });

        }
    }
}
