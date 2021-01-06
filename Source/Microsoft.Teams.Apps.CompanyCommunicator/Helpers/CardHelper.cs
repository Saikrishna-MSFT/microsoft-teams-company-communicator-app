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
    /// Card Helper.
    /// </summary>
    public class CardHelper : ICard
    {
        private readonly IConfiguration configuration;
        private readonly string welcomeText = "Appy is your official Microsoft Teams Platform assistant!";
        private readonly string welocmeDescription1 = "In this NDA community, you will get the most up to date information on all things Microsoft Teams, have an opportunity to join any of our Monthly Office Hours, connect and network with other Partners who are also interested in Microsoft Teams as well as interact with the Microsoft Teams Engineering team.";
        private readonly string welocmeDescription2 = "We are happy you are here!";

        /// <summary>
        /// Initializes a new instance of the <see cref="CardHelper"/> class.
        /// </summary>
        /// <param name="configuration">IConfiguration instance.</param>
        public CardHelper(IConfiguration configuration)
        {
            this.configuration = configuration;
        }

        /// <summary>
        /// Get Welcome Card Attachment.
        /// </summary>
        /// <returns>Welcome card.</returns>
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
                                },
                            },
                            new AdaptiveTextBlock()
                            {
                                Text = this.welocmeDescription1,
                                Wrap = true,
                                Size = AdaptiveTextSize.Small,
                            },
                            new AdaptiveTextBlock()
                            {
                                Text = this.welocmeDescription2,
                                Wrap = true,
                                Size = AdaptiveTextSize.Small,
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
                                             new AdaptiveTextBlock() { Text = "Release Updates", Color = AdaptiveTextColor.Accent, Size = AdaptiveTextSize.Medium, Spacing = AdaptiveSpacing.None, HorizontalAlignment = AdaptiveHorizontalAlignment.Center },
                                         },
                                         SelectAction = new AdaptiveOpenUrlAction()
                                         {
                                             Url = new Uri(this.configuration["ReleaseUpdates"]),
                                             Title = "Release Updates",
                                         },
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
                                             new AdaptiveTextBlock() { Text = "Platform and GraphAPI", Color = AdaptiveTextColor.Accent, Size = AdaptiveTextSize.Medium, Spacing = AdaptiveSpacing.None, HorizontalAlignment = AdaptiveHorizontalAlignment.Center },
                                         },
                                         SelectAction = new AdaptiveOpenUrlAction()
                                         {
                                             Url = new Uri(this.configuration["PlatformAndGraphAPI"]),
                                             Title = "Platform and GraphAPI",
                                         },
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
                                             new AdaptiveTextBlock() { Text = "Office Hours", Color = AdaptiveTextColor.Accent, Size = AdaptiveTextSize.Medium, Spacing = AdaptiveSpacing.None, HorizontalAlignment = AdaptiveHorizontalAlignment.Center },
                                         },
                                         SelectAction = new AdaptiveOpenUrlAction()
                                         {
                                             Url = new Uri(this.configuration["OfficeHours"]),
                                             Title = "Office Hours",
                                         },
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
                                             new AdaptiveTextBlock() { Text = "Industry - FLW and Healthcare", Color = AdaptiveTextColor.Accent, Size = AdaptiveTextSize.Medium, Spacing = AdaptiveSpacing.None, HorizontalAlignment = AdaptiveHorizontalAlignment.Center },
                                         },
                                         SelectAction = new AdaptiveOpenUrlAction()
                                         {
                                             Url = new Uri(this.configuration["IndustryFLWandHealthcare"]),
                                             Title = "Industry - FLW and Healthcare",
                                         },
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
                                             new AdaptiveTextBlock() { Text = "Events", Color = AdaptiveTextColor.Accent, Size = AdaptiveTextSize.Medium, Spacing = AdaptiveSpacing.None, HorizontalAlignment = AdaptiveHorizontalAlignment.Center },
                                         },
                                         SelectAction = new AdaptiveOpenUrlAction()
                                         {
                                             Url = new Uri(this.configuration["Events"]),
                                             Title = "Events",
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
