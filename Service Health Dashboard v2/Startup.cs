using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.OData;
using Microsoft.AspNetCore.OData.NewtonsoftJson;
using Microsoft.AspNetCore.SpaServices.ReactDevelopmentServer;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.ServiceHealthHub.Core;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;

namespace Service_Health_Dashboard_v2
{
    public class FeatureStatusInfo
    {
        public string FeatureDisplayName { get; set; }
        public string FeatureName { get; set; }
        public string FeatureServiceStatus { get; set; }
        public string FeatureServiceStatusDisplayName { get; set; }
    }

    public class WorkloadStatus
    {
        public List<FeatureStatusInfo> FeatureStatus { get; set; }
        public string Id { get; set; }
        public List<object> IncidentIds { get; set; }
        public string Status { get; set; }
        public string StatusDisplayName { get; set; }
        public DateTime StatusTime { get; set; }
        public string Workload { get; set; }
        public string WorkloadDisplayName { get; set; }
    }

    public class StatusInfo
    {
        public List<WorkloadStatus> value { get; set; }
    }

    public class ServiceMessage
    {
        public string MessageText { get; set; }
        public DateTime PublishedTime { get; set; }
    }
    public class ServiceCommunication
    {
        public List<object> AffectedWorkloadDisplayNames { get; set; }
        public List<object> AffectedWorkloadNames { get; set; }
        public string Status { get; set; }
        public string Workload { get; set; }
        public string WorkloadDisplayName { get; set; }
        public object ActionType { get; set; }
        public List<object> AdditionalDetails { get; set; }
        public int AffectedTenantCount { get; set; }
        public object AffectedUserCount { get; set; }
        public string Classification { get; set; }
        public DateTime? EndTime { get; set; }
        public string Feature { get; set; }
        public string FeatureDisplayName { get; set; }
        public string UserFunctionalImpact { get; set; }
        public string Id { get; set; }
        public string ImpactDescription { get; set; }
        public DateTime LastUpdatedTime { get; set; }
        public string MessageType { get; set; }
        public List<ServiceMessage> Messages { get; set; }
        public object PostIncidentDocumentUrl { get; set; }
        public string Severity { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime? ActionRequiredByDate { get; set; }
        public string Title { get; set; }
        public bool IsMajorChange { get; set; }
        public List<object> MessageTagNames { get; set; }
        public string SHD_WorkItemID { get; set; }
        public string SHD_WorkItemURL { get; set; }
    }

    public class ServiceCommunicationCollection
    {
        public List<ServiceCommunication> value { get; set; }
    }

    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
            GlobalConfiguration.LoadConfiguration(configuration);
            Cache.Initialize();
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {

            services.Configure<CookiePolicyOptions>(options => {
                options.CheckConsentNeeded = context => true;
                options.MinimumSameSitePolicy = SameSiteMode.None;
            });

            /* services.AddAuthentication(AzureADDefaults.AuthenticationScheme)
                .AddAzureAD(options => Configuration.Bind("AzureAd", options));

            services.Configure<OpenIdConnectOptions>(AzureADDefaults.OpenIdScheme, options =>
            {
                options.Authority = options.Authority + "/v2.0/";
                options.TokenValidationParameters.ValidateIssuer = false;
            }); */

            /* JwtSecurityTokenHandler.DefaultInboundClaimTypeMap.Clear();

            services.AddMicrosoftIdentityWebApiAuthentication(Configuration);

            services.AddControllers(options =>
            {
                var policy = new AuthorizationPolicyBuilder()
                    .RequireAuthenticatedUser()
                    .RequireClaim("email")
                    .Build();
                options.Filters.Add(new AuthorizeFilter(policy));
            }); */

            services
                .AddAuthentication(options => {
                   options.DefaultScheme = JwtBearerDefaults.AuthenticationScheme;
                })
                .AddJwtBearer(options => {
                    options.Authority = string.Format(
                        "https://login.microsoftonline.com/{0}/v2.0/",
                        string.IsNullOrWhiteSpace(GlobalConfiguration.Instance.ClientTenantDomain) ? 
                            GlobalConfiguration.Instance.TenantDomain :
                            GlobalConfiguration.Instance.ClientTenantDomain
                        );
                    options.Audience = GlobalConfiguration.Instance.ClientAppId;
                } );
            // .AddJwtBearer(options => Configuration.Bind("AzureAd", options)); 

            services.AddControllersWithViews().AddNewtonsoftJson();
            if (GlobalConfiguration.Instance.ApplicationInsightsInstrumentationKey != null)
                services.AddApplicationInsightsTelemetry(GlobalConfiguration.Instance.ApplicationInsightsInstrumentationKey);

            services.AddMemoryCache();
            services.AddControllers()
                    .AddOData(opt => opt.Count().Filter().Select().Expand().OrderBy())
                    .AddJsonOptions(options =>
                    {
                        options.JsonSerializerOptions.PropertyNamingPolicy = null;
                        options.JsonSerializerOptions.IncludeFields = true;
                    })
                    .AddNewtonsoftJson(options =>
                    {
                        options.SerializerSettings.ContractResolver = new CamelCasePropertyNamesContractResolver();
                        options.SerializerSettings.NullValueHandling = Newtonsoft.Json.NullValueHandling.Include;
                    })
                    .AddODataNewtonsoftJson();

            // In production, the React files will be served from this directory
            services.AddSpaStaticFiles(configuration =>
            {
                configuration.RootPath = "ClientApp/build";
            });
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            app.UseHttpsRedirection();

            app.UseStaticFiles();
            app.UseSpaStaticFiles();

            app.UseRouting();
            app.UseAuthentication();
            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                    name: "default",
                    pattern: "{controller}/{action=Index}/{id?}");
            });

            app.UseSpa(spa =>
            {
                spa.Options.SourcePath = "ClientApp";

                if (env.IsDevelopment())
                {
                    spa.UseReactDevelopmentServer(npmScript: "start");
                }
            });
        }
    }
}
