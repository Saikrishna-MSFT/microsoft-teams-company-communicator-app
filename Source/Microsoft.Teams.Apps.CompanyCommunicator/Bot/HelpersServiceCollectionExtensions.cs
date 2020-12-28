namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Teams.Apps.CompanyCommunicator.Helpers;
    using Microsoft.Teams.Apps.CompanyCommunicator.Interfaces;

    /// <summary>
    /// Extension class for registering helper services in DI container.
    /// </summary>
    public static class HelpersServiceCollectionExtensions
    {
        /// <summary>
        /// Extension method to register helper services in DI container.
        /// </summary>
        /// <param name="services">IServiceCollection instance.</param>
        public static void AddHelpers(this IServiceCollection services)
        {
            services.AddSingleton<ICard, CardHelper>();
        }
    }
}
