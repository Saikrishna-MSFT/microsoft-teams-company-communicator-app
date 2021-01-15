// <copyright file="UserTeamsActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Interfaces;

    /// <summary>
    /// Company Communicator User Bot.
    /// Captures user data, team data.
    /// </summary>
    public class UserTeamsActivityHandler : TeamsActivityHandler
    {
        private static readonly string TeamRenamedEventType = "teamRenamed";

        private readonly TeamsDataCapture teamsDataCapture;

        private readonly ILogger<AuthorTeamsActivityHandler> logger;
        private readonly IConfiguration configuration;
        private readonly ICard cardHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserTeamsActivityHandler"/> class.
        /// </summary>
        /// <param name="teamsDataCapture">Teams data capture service.</param>
        /// <param name="logger">ILogger instance.</param>
        /// <param name="configuration">IConfiguration instance.</param>
        /// <param name="cardHelper">ICard instance.</param>
        public UserTeamsActivityHandler(
            TeamsDataCapture teamsDataCapture,
            ILogger<AuthorTeamsActivityHandler> logger,
            IConfiguration configuration,
            ICard cardHelper)
        {
            this.teamsDataCapture = teamsDataCapture ?? throw new ArgumentNullException(nameof(teamsDataCapture));
            this.logger = logger;
            this.configuration = configuration;
            this.cardHelper = cardHelper;
        }

        /// <summary>
        /// Invoked when a conversation update activity is received from the channel.
        /// </summary>
        /// <param name="turnContext">The context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnConversationUpdateActivityAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            // base.OnConversationUpdateActivityAsync is useful when it comes to responding to users being added to or removed from the conversation.
            // For example, a bot could respond to a user being added by greeting the user.
            // By default, base.OnConversationUpdateActivityAsync will call <see cref="OnMembersAddedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been added or <see cref="OnMembersRemovedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been removed. base.OnConversationUpdateActivityAsync checks the member ID so that it only responds to updates regarding members other than the bot itself.
            await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);

            var activity = turnContext.Activity;

            var isTeamRenamed = this.IsTeamInformationUpdated(activity);
            if (isTeamRenamed)
            {
                await this.teamsDataCapture.OnTeamInformationUpdatedAsync(activity);
            }

            if (activity.MembersAdded != null)
            {
                // await this.teamsDataCapture.OnBotAddedAsync(activity);
            }

            if (activity.MembersRemoved != null)
            {
                // await this.teamsDataCapture.OnBotRemovedAsync(activity);
            }
        }

        ///// <summary>
        ///// Gets called when when members other than the bot join the conversation.
        ///// </summary>
        ///// <param name="membersAdded">A list of all the members added to the conversation.</param>
        ///// <param name="turnContext">A strongly-typed context object for this turn.</param>
        ///// <param name="cancellationToken"> A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        ///// <returns>A task that represents the work queued to execute.</returns>
        // protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        // {
        //    this.logger.LogInformation("Inside OnMembersAddedAsync()");
        //    try
        //    {
        //        var credentials = new MicrosoftAppCredentials(this.configuration["AuthorAppId"], this.configuration["AuthorAppPassword"]);
        //        ConversationReference conversationReference = null;
        //        foreach (var member in membersAdded)
        //        {
        //            if (member.Id != turnContext.Activity.Recipient.Id)
        //            {
        //                var proactiveMessage = MessageFactory.Attachment(this.cardHelper.GetWelcomeCard());
        //                proactiveMessage.TeamsNotifyUser();
        //                var conversationParameters = new ConversationParameters
        //                {
        //                    IsGroup = false,
        //                    Bot = turnContext.Activity.Recipient,
        //                    Members = new ChannelAccount[] { member },
        //                    TenantId = turnContext.Activity.Conversation.TenantId,
        //                };
        //                await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
        //                    turnContext.Activity.ChannelId,
        //                    turnContext.Activity.ServiceUrl,
        //                    credentials,
        //                    conversationParameters,
        //                    async (t1, c1) =>
        //                    {
        //                        conversationReference = t1.Activity.GetConversationReference();
        //                        await ((BotFrameworkAdapter)turnContext.Adapter).ContinueConversationAsync(
        //                            this.configuration["AuthorAppId"],
        //                            conversationReference,
        //                            async (t2, c2) =>
        //                            {
        //                                await t2.SendActivityAsync(proactiveMessage, c2);
        //                            },
        //                            cancellationToken);
        //                    },
        //                    cancellationToken);
        //            }
        //            else
        //            {
        //                var connectorClient = new ConnectorClient(new Uri(turnContext.Activity.ServiceUrl), this.configuration["AuthorAppId"], this.configuration["AuthorAppPassword"]);
        //                var message = turnContext.Activity;
        //                var channelData = message.GetChannelData<TeamsChannelData>();
        //                var teamorConversationId = channelData.Team != null ? channelData.Team.Id : message.Conversation.Id;
        //                var members = await connectorClient.Conversations.GetConversationMembersAsync(teamorConversationId);
        //                foreach (var mem in members)
        //                {
        //                    var card = this.cardHelper.GetWelcomeCard();
        //                    var replyMessage = Activity.CreateMessageActivity();
        //                    var parameters = new ConversationParameters
        //                    {
        //                        Members = new ChannelAccount[] { new ChannelAccount(mem.Id) },
        //                        ChannelData = new TeamsChannelData
        //                        {
        //                            Tenant = channelData.Tenant,
        //                            Notification = new NotificationInfo() { Alert = true },
        //                        },
        //                    };

        // var conversationResource = await connectorClient.Conversations.CreateConversationAsync(parameters);
        //                    replyMessage.ChannelData = new TeamsChannelData() { Notification = new NotificationInfo(true) };
        //                    replyMessage.Conversation = new ConversationAccount(id: conversationResource.Id.ToString());
        //                    replyMessage.TextFormat = TextFormatTypes.Xml;
        //                    replyMessage.Attachments.Add(card);
        //                    await connectorClient.Conversations.SendToConversationAsync((Activity)replyMessage);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        this.logger.LogError("Exception OnMembersAddedAsync() : " + ex.ToString());
        //    }
        // }
        private bool IsTeamInformationUpdated(IConversationUpdateActivity activity)
        {
            if (activity == null)
            {
                return false;
            }

            var channelData = activity.GetChannelData<TeamsChannelData>();
            if (channelData == null)
            {
                return false;
            }

            return UserTeamsActivityHandler.TeamRenamedEventType.Equals(channelData.EventType, StringComparison.OrdinalIgnoreCase);
        }
    }
}