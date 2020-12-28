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
        private readonly string welocmeDescription1 = "Access on-demand information (assets and resources) pertaining to Microsoft Teams app platform";
        private readonly string welocmeDescription2 = "Influence Teams Platform Ecosystem group and contribute back to apps, initiatives and opportunities";
        private readonly string welocmeDescription3 = "Seek support on questions and opportunities to drive usage of apps on Microsoft Teams";
        private readonly string welcomeDesc4 = "➕ Add to your apps:";
        private readonly string welcomeDesc5 = " Tap ‘Add’ button on the top bar to install Appy!​";

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
                            new AdaptiveTextBlock()
                            {
                                Text = $"- {this.welocmeDescription1}",
                                Wrap = true,
                                Size = AdaptiveTextSize.Small,
                            },
                            new AdaptiveTextBlock()
                            {
                                Text = $"- {this.welocmeDescription2}",
                                Wrap = true,
                                Size = AdaptiveTextSize.Small,
                            },
                            new AdaptiveTextBlock()
                            {
                                Text = $"- {this.welocmeDescription3}",
                                Wrap = true,
                                Size = AdaptiveTextSize.Small,
                            },
                            new AdaptiveColumnSet()
                            {
                                Style = AdaptiveContainerStyle.Accent,
                                Bleed = true,
                                Columns = new List<AdaptiveColumn>()
                                {
                                     new AdaptiveColumn()
                                    {
                                        Width = "100",
                                        Items = new List<AdaptiveElement>()
                                        {
                                            new AdaptiveRichTextBlock()
                                            {
                                                Inlines = new List<IAdaptiveInline>()
                                                {
                                                    new AdaptiveTextRun()
                                                    {
                                                        Text = this.welcomeDesc4,
                                                        Weight = AdaptiveTextWeight.Bolder,
                                                        Size = AdaptiveTextSize.Small,
                                                    },
                                                    new AdaptiveTextRun()
                                                    {
                                                        Text = this.welcomeDesc5,
                                                        Size = AdaptiveTextSize.Small,
                                                    },
                                                },
                                            },
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
