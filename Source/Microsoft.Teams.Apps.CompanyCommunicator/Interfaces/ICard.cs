namespace Microsoft.Teams.Apps.CompanyCommunicator.Interfaces
{
    using Microsoft.Bot.Schema;

    /// <summary>
    /// Card interface
    /// </summary>
    public interface ICard
    {
        /// <summary>
        /// Get welcome card
        /// </summary>
        /// <returns>Welcome card</returns>
        Attachment GetWelcomeCard();
    }
}
