// <copyright file="CompanyCommunicatorBot.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
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
    /// Company Communicator Bot.
    /// </summary>
    public class CompanyCommunicatorBot : TeamsActivityHandler
    {
        private static readonly string TeamRenamedEventType = "teamRenamed";

        private readonly TeamsDataCapture teamsDataCapture;
        private readonly IConfiguration configuration;
        private readonly ICard cardHelper;
        private readonly ILogger<CompanyCommunicatorBot> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyCommunicatorBot"/> class.
        /// </summary>
        /// <param name="teamsDataCapture">Teams data capture service.</param>
        /// <param name="configuration">IConfiguration instance.</param>
        /// <param name="cardHelper">ICard instance.</param>
        /// <param name="logger">ILogger instance.</param>
        public CompanyCommunicatorBot(TeamsDataCapture teamsDataCapture, IConfiguration configuration, ICard cardHelper, ILogger<CompanyCommunicatorBot> logger)
        {
            this.teamsDataCapture = teamsDataCapture;
            this.configuration = configuration;
            this.cardHelper = cardHelper;
            this.logger = logger;
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
            var botId = activity.Recipient.Id;

            var isTeamRenamed = this.IsTeamInformationUpdated(activity);
            if (isTeamRenamed)
            {
                await this.teamsDataCapture.OnTeamInformationUpdatedAsync(activity);
            }

            // Take action if this event includes the bot being added
            if (activity.MembersAdded?.FirstOrDefault(p => p.Id == botId) != null)
            {
                await this.teamsDataCapture.OnBotAddedAsync(activity);
            }

            // Take action if this event includes the bot being removed
            if (activity.MembersRemoved?.FirstOrDefault(p => p.Id == botId) != null)
            {
                await this.teamsDataCapture.OnBotRemovedAsync(activity);
            }
        }

        /// <summary>
        /// Gets called when when members other than the bot join the conversation
        /// </summary>
        /// <param name="membersAdded">A list of all the members added to the conversation</param>
        /// <param name="turnContext">A strongly-typed context object for this turn.</param>
        /// <param name="cancellationToken"> A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            this.logger.LogWarning("Inside OnMembersAddedAsync()");
            try
            {
                var credentials = new MicrosoftAppCredentials(this.configuration["MicrosoftAppId"], this.configuration["MicrosoftAppPassword"]);
                ConversationReference conversationReference = null;
                foreach (var member in membersAdded)
                {
                    if (member.Id != turnContext.Activity.Recipient.Id)
                    {
                        var proactiveMessage = MessageFactory.Attachment(this.cardHelper.GetWelcomeCard());
                        proactiveMessage.TeamsNotifyUser();
                        var conversationParameters = new ConversationParameters
                        {
                            IsGroup = false,
                            Bot = turnContext.Activity.Recipient,
                            Members = new ChannelAccount[] { member },
                            TenantId = turnContext.Activity.Conversation.TenantId,
                        };
                        await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                            turnContext.Activity.ChannelId,
                            turnContext.Activity.ServiceUrl,
                            credentials,
                            conversationParameters,
                            async (t1, c1) =>
                            {
                                conversationReference = t1.Activity.GetConversationReference();
                                await ((BotFrameworkAdapter)turnContext.Adapter).ContinueConversationAsync(
                                    this.configuration["MicrosoftAppId"],
                                    conversationReference,
                                    async (t2, c2) =>
                                    {
                                        await t2.SendActivityAsync(proactiveMessage, c2);
                                    },
                                    cancellationToken);
                            },
                            cancellationToken);
                    }
                    else
                    {
                        var connectorClient = new ConnectorClient(new Uri(turnContext.Activity.ServiceUrl), this.configuration["MicrosoftAppId"], this.configuration["MicrosoftAppPassword"]);
                        var message = turnContext.Activity;
                        var channelData = message.GetChannelData<TeamsChannelData>();
                        var teamorConversationId = channelData.Team != null ? channelData.Team.Id : message.Conversation.Id;
                        var members = await connectorClient.Conversations.GetConversationMembersAsync(teamorConversationId);
                        foreach (var mem in members)
                        {
                            var card = this.cardHelper.GetWelcomeCard();
                            var replyMessage = Activity.CreateMessageActivity();
                            var parameters = new ConversationParameters
                            {
                                Members = new ChannelAccount[] { new ChannelAccount(mem.Id) },
                                ChannelData = new TeamsChannelData
                                {
                                    Tenant = channelData.Tenant,
                                    Notification = new NotificationInfo() { Alert = true },
                                },
                            };

                            var conversationResource = await connectorClient.Conversations.CreateConversationAsync(parameters);
                            replyMessage.ChannelData = new TeamsChannelData() { Notification = new NotificationInfo(true) };
                            replyMessage.Conversation = new ConversationAccount(id: conversationResource.Id.ToString());
                            replyMessage.TextFormat = TextFormatTypes.Xml;
                            replyMessage.Attachments.Add(card);
                            await connectorClient.Conversations.SendToConversationAsync((Activity)replyMessage);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError("Exception OnMembersAddedAsync() : " + ex.ToString());
            }
        }

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

            return CompanyCommunicatorBot.TeamRenamedEventType.Equals(channelData.EventType, StringComparison.OrdinalIgnoreCase);
        }
    }
}