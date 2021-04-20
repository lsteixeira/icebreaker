// <copyright file="PairUpNotificationAdaptiveCard.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Helpers.AdaptiveCards
{
    using System;
    using global::AdaptiveCards.Templating;
    using Icebreaker.Properties;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;

    /// <summary>
    /// Builder class for the pairup notification card
    /// </summary>
    public class PairUpNotificationAdaptiveCard : AdaptiveCardBase
    {
        /// <summary>
        /// Default marker string in the UPN that indicates a user is externally-authenticated
        /// </summary>
        private const string ExternallyAuthenticatedUpnMarker = "#ext#";

        private static readonly Lazy<AdaptiveCardTemplate> AdaptiveCardTemplate =
            new Lazy<AdaptiveCardTemplate>(() => CardTemplateHelper.GetAdaptiveCardTemplate(AdaptiveCardName.PairUpNotification));

        /// <summary>
        /// Creates the pairup notification card.
        /// </summary>
        /// <param name="teamName">The team name.</param>
        /// <param name="sender">The user who will be sending this card.</param>
        /// <param name="recipient">The user who will be receiving this card.</param>
        /// <param name="botDisplayName">The bot display name.</param>
        /// <returns>Pairup notification card</returns>
        public static Attachment GetCard(string teamName, TeamsChannelAccount sender, TeamsChannelAccount recipient1, TeamsChannelAccount recipient2, string botDisplayName)
        {
            // Guest users may not have their given name specified in AAD, so fall back to the full name if needed
            var senderGivenName = string.IsNullOrEmpty(sender.GivenName) ? sender.Name : sender.GivenName;
            var recipient1GivenName = string.IsNullOrEmpty(recipient1.GivenName) ? recipient1.Name : recipient1.GivenName;
            var recipient2GivenName = string.IsNullOrEmpty(recipient2.GivenName) ? recipient2.Name : recipient2.GivenName;

            // To start a chat with a guest user, use their external email, not the UPN
            var recipient1Upn = !IsGuestUser(recipient1) ? recipient1.UserPrincipalName : recipient1.Email;
            var recipient2Upn = !IsGuestUser(recipient2) ? recipient2.UserPrincipalName : recipient2.Email;

            var meetingTitle = string.Format(Resources.MeetupTitle, senderGivenName, recipient1GivenName, recipient2GivenName);
            var meetingContent = string.Format(Resources.MeetupContent, botDisplayName);
            var meetingLink = "https://teams.microsoft.com/l/meeting/new?subject=" + Uri.EscapeDataString(meetingTitle) + "&attendees=" + recipient1Upn + ";" + recipient2Upn + "&content=" + Uri.EscapeDataString(meetingContent);
            var personUpn = "" + recipient1Upn + ";" + recipient2Upn;

            var cardData = new
            {
                matchUpCardTitleContent = Resources.MatchUpCardTitleContent,
                matchUpCardMatchedText = string.Format(Resources.MatchUpCardMatchedText, recipient1.Name, recipient2.Name),
                matchUpCardContentPart1 = string.Format(Resources.MatchUpCardContentPart1, botDisplayName, teamName, recipient1.Name, recipient2.Name),
                matchUpCardContentPart2 = Resources.MatchUpCardContentPart2,
                chatWithMatchButtonText = string.Format(Resources.ChatWithMatchButtonText, recipient1GivenName, recipient2GivenName),
                chatWithMessageGreeting = Resources.ChatWithMessageGreeting,
                pauseMatchesButtonText = Resources.PausePairingsButtonText,
                proposeMeetupButtonText = Resources.ProposeMeetupButtonText,
                personUpn = personUpn,
                meetingLink,
            };

            return GetCard(AdaptiveCardTemplate.Value, cardData);
        }

        /// <summary>
        /// Checks whether or not the account is a guest user.
        /// </summary>
        /// <param name="account">The <see cref="TeamsChannelAccount"/> user to check.</param>
        /// <returns>True if the account is a guest user, false otherwise.</returns>
        private static bool IsGuestUser(TeamsChannelAccount account)
        {
            return account.UserPrincipalName.IndexOf(ExternallyAuthenticatedUpnMarker, StringComparison.InvariantCultureIgnoreCase) >= 0;
        }
    }
}