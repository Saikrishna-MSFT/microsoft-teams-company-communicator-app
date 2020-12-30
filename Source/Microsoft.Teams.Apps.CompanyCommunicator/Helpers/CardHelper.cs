// <copyright file="CardHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Helpers
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CompanyCommunicator.Interfaces;

    /// <summary>
    /// Card Helper
    /// </summary>
    public class CardHelper : ICard
    {
        private readonly IConfiguration configuration;
        private readonly string welcomeText = "Appy is your official Microsoft Teams Platform assistant!";

        /// <summary>
        /// Initializes a new instance of the <see cref="CardHelper"/> class.
        /// </summary>
        /// <param name="configuration">IConfiguration instance.</param>
        public CardHelper(IConfiguration configuration)
        {
            this.configuration = configuration;
        }

        /// <summary>
        /// Get Welcome Card Attachment
        /// </summary>
        /// <returns>Welcome card</returns>
        public Attachment GetWelcomeCard()
        {
            var welcomeCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>()
                {
                    new AdaptiveContainer()
                    {
                        Items = new List<AdaptiveElement>()
                        {
                            new AdaptiveImage()
                            {
                                Url = new Uri(this.configuration["BaseUri"] + "/Images/WelcomeCard.png"),
                            },
                            new AdaptiveRichTextBlock()
                            {
                                Inlines = new List<IAdaptiveInline>()
                                {
                                     new AdaptiveTextRun()
                                     {
                                         Text = this.welcomeText,
                                         Weight = AdaptiveTextWeight.Bolder,
                                         Size = AdaptiveTextSize.Small,
                                     },
                                     new AdaptiveTextRun()
                                     {
                                         Text = " Use Appy to:​",
                                         Size = AdaptiveTextSize.Small,
                                     },
                                },
                            },
                            new AdaptiveColumnSet()
                            {
                                Columns = new List<AdaptiveColumn>()
                                {
                                    new AdaptiveColumn()
                                    {
                                         Width = AdaptiveColumnWidth.Auto,
                                         Items = new List<AdaptiveElement>()
                                         {
                                             new AdaptiveImage() { Url = new Uri(this.configuration["BaseUri"] + "/Images/EmailNotifications.png"), Size = AdaptiveImageSize.Small, Style = AdaptiveImageStyle.Default, SelectAction = new AdaptiveOpenUrlAction() { Url = new Uri(this.configuration["EmailNotifications"]), Title = "Email Notifications" }, HorizontalAlignment = AdaptiveHorizontalAlignment.Center, Spacing = AdaptiveSpacing.None },
                                         },
                                    },

                                    new AdaptiveColumn()
                                    {
                                         Width = AdaptiveColumnWidth.Auto,
                                         Items = new List<AdaptiveElement>()
                                         {
                                             new AdaptiveTextBlock() { Text = "Email Notifications", Color = AdaptiveTextColor.Accent, Size = AdaptiveTextSize.Medium, Spacing = AdaptiveSpacing.None, HorizontalAlignment = AdaptiveHorizontalAlignment.Center },
                                         },
                                         SelectAction = new AdaptiveOpenUrlAction()
                                         {
                                             Url = new Uri(this.configuration["EmailNotifications"]),
                                             Title = "Email Notifications",
                                         },
                                    },
                                },
                            },
                        },
                    },
                },
            };
            return new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = welcomeCard,
            };
        }
    }
}
