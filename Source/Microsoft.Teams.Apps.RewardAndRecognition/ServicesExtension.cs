// <copyright file="ServicesExtension.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition
{
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Localization;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Teams.Apps.RewardAndRecognition.Bot;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// Class to extend ServiceCollection.
    /// </summary>
    public static class ServicesExtension
    {
        /// <summary>
        /// Adds application configuration settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddConfigurationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            services.Configure<RewardAndRecognitionActivityHandlerOptions>(options =>
            {
                options.TenantId = configuration.GetValue<string>("Bot:TenantId");
                options.AppBaseUri = configuration.GetValue<string>("Bot:AppBaseUri");
            });
            services.Configure<StorageOptions>(options =>
            {
                options.ConnectionString = configuration.GetValue<string>("Storage:ConnectionString");
            });
            services.Configure<SearchServiceOptions>(options =>
            {
                options.SearchServiceName = configuration.GetValue<string>("Search:SearchServiceName");
                options.SearchServiceQueryApiKey = configuration.GetValue<string>("Search:SearchServiceQueryApiKey");
                options.SearchServiceAdminApiKey = configuration.GetValue<string>("Search:SearchServiceAdminApiKey");
            });
            services.Configure<TelemetryOptions>(options =>
            {
                options.InstrumentationKey = configuration.GetValue<string>("ApplicationInsights:InstrumentationKey");
            });
        }

        /// <summary>
        /// Adds providers to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        public static void AddProviders(this IServiceCollection services)
        {
            services.AddSingleton<IConfigureAdminStorageProvider, ConfigureAdminStorageProvider>();
            services.AddSingleton<IAwardNominationStorageProvider, AwardNominationStorageProvider>();
            services.AddSingleton<IRewardCycleStorageProvider, RewardCycleStorageProvider>();
            services.AddSingleton<IEndorsementsStorageProvider, EndorsementsStorageProvider>();
            services.AddSingleton<ITeamStorageProvider, TeamStorageProvider>();
            services.AddSingleton<IAwardsStorageProvider, AwardsStorageProvider>();
            services.AddSingleton<IAwardNominationSearchService, AwardNominationSearchService>();
        }

        /// <summary>
        /// Adds credential providers for authentication.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddCredentialProviders(this IServiceCollection services, IConfiguration configuration)
        {
            services
                .AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
            services.AddSingleton(new MicrosoftAppCredentials(configuration.GetValue<string>("MicrosoftAppId"), configuration.GetValue<string>("MicrosoftAppPassword")));
        }

        /// <summary>
        /// Adds localization settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddLocalizationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddLocalization(options => options.ResourcesPath = "Resources");
            services.Configure<RequestLocalizationOptions>(options =>
            {
                var defaultCulture = CultureInfo.GetCultureInfo(configuration.GetValue<string>("i18n:DefaultCulture"));
                var supportedCultures = configuration.GetValue<string>("i18n:SupportedCultures").Split(',')
                    .Select(culture => CultureInfo.GetCultureInfo(culture))
                    .ToList();

                options.DefaultRequestCulture = new RequestCulture(defaultCulture);
                options.SupportedCultures = supportedCultures;
                options.SupportedUICultures = supportedCultures;

                options.RequestCultureProviders = new List<IRequestCultureProvider>
                {
                    new BotLocalizationCultureProvider(),
                };
            });
        }

        /// <summary>
        /// Adds bot framework adapter to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        public static void AddBotFrameworkAdapter(this IServiceCollection services)
        {
            // Create the Bot Framework Adapter with error handling enabled.
            services.AddSingleton<IBotFrameworkHttpAdapter, RewardAndRecognitionAdapterWithErrorHandler>();

            // Create the Translation Middleware that will be added to the middleware pipeline in the AdapterWithErrorHandler
            services.AddSingleton<RewardAndRecognitionActivityMiddleware>();
            services.AddTransient<IBot, RewardAndRecognitionActivityHandler>();
            services.AddTransient(serviceProvider => (BotFrameworkAdapter)serviceProvider.GetRequiredService<IBotFrameworkHttpAdapter>());
        }
    }
}
