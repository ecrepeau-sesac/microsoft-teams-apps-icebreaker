// <copyright file="IcebreakerBot.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Icebreaker.Helpers;
    using Icebreaker.Helpers.AdaptiveCards;
    using Icebreaker.Interfaces;
    using Icebreaker.Properties;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;

    /// <summary>
    /// Implements the core logic for the Icebreaker bot.
    /// </summary>
    public class IcebreakerBot : TeamsActivityHandler
    {
        private readonly IBotDataProvider dataProvider;
        private readonly ConversationHelper conversationHelper;
        private readonly MicrosoftAppCredentials appCredentials;
        private readonly TelemetryClient telemetryClient;
        private readonly string botDisplayName;
        private readonly bool disableTenantFilter;
        private readonly HashSet<string> allowedTenantIds;

        /// <summary>
        /// Initializes a new instance of the <see cref="IcebreakerBot"/> class.
        /// </summary>
        public IcebreakerBot(IBotDataProvider dataProvider, ConversationHelper conversationHelper, MicrosoftAppCredentials appCredentials, TelemetryClient telemetryClient)
        {
            this.dataProvider = dataProvider;
            this.conversationHelper = conversationHelper;
            this.appCredentials = appCredentials;
            this.telemetryClient = telemetryClient;
            this.botDisplayName = ConfigurationManager.AppSettings["BotDisplayName"];
            this.disableTenantFilter = Convert.ToBoolean(ConfigurationManager.AppSettings["DisableTenantFilter"], CultureInfo.InvariantCulture);
            var allowedTenants = ConfigurationManager.AppSettings["AllowedTenants"];
            this.allowedTenantIds = allowedTenants
                ?.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                ?.Select(p => p.Trim())
                .ToHashSet();
        }

        /// <inheritdoc/>
        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default(CancellationToken))
        {
            try
            {
                this.LogActivityTelemetry(turnContext.Activity);

                if (!this.ValidateTenant(turnContext))
                {
                    return;
                }

                string locale = turnContext?.Activity.Entities?.FirstOrDefault(entity => entity.Type == "clientInfo")?.Properties["locale"]?.ToString();
                if (!string.IsNullOrEmpty(locale))
                {
                    CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo(locale);
                }

                await base.OnTurnAsync(turnContext, cancellationToken);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
            }
        }

        /// <inheritdoc/>
        protected override async Task OnConversationUpdateActivityAsync(ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var teamsChannelData = turnContext.Activity.GetChannelData<TeamsChannelData>();
            if (string.IsNullOrEmpty(teamsChannelData?.Team?.Id))
            {
                return;
            }

            await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);
        }

        /// <inheritdoc/>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            if (membersAdded?.Count() > 0)
            {
                var message = turnContext.Activity;
                string myBotId = message.Recipient.Id;
                string teamId = message.Conversation.Id;
                var teamsChannelData = message.GetChannelData<TeamsChannelData>();

                foreach (var member in membersAdded)
                {
                    if (member.Id == myBotId)
                    {
                        this.telemetryClient.TrackTrace($"Bot installed to team {teamId}");

                        var properties = new Dictionary<string, string>
                        {
                            { "Scope", message.Conversation?.ConversationType },
                            { "TeamId", teamId },
                            { "InstallerId", message.From.Id },
                        };
                        this.telemetryClient.TrackEvent("AppInstalled", properties);

                        var personThatAddedBot = (await this.conversationHelper.GetMemberAsync(turnContext, message.From.Id, cancellationToken))?.Name;
                        await this.SaveAddedToTeam(message.ServiceUrl, teamId, teamsChannelData.Tenant.Id, personThatAddedBot);
                        await this.WelcomeTeam(turnContext, personThatAddedBot, cancellationToken);
                    }
                    else
                    {
                        this.telemetryClient.TrackTrace($"New member {member.Id} added to team {teamsChannelData.Team.Id}");
                        await this.WelcomeUser(turnContext, member.Id, teamsChannelData.Tenant.Id, teamsChannelData.Team.Id, cancellationToken);
                    }
                }
            }

            await base.OnMembersAddedAsync(membersAdded, turnContext, cancellationToken);
        }

        /// <inheritdoc/>
        protected override async Task OnMembersRemovedAsync(IList<ChannelAccount> membersRemoved, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var message = turnContext.Activity;
            string myBotId = message.Recipient.Id;
            string teamId = message.Conversation.Id;
            var teamsChannelData = message.GetChannelData<TeamsChannelData>();

            if (message.MembersRemoved?.Any(x => x.Id == myBotId) == true)
            {
                this.telemetryClient.TrackTrace($"Bot removed from team {teamId}");

                var properties = new Dictionary<string, string>
                {
                    { "Scope", message.Conversation?.ConversationType },
                    { "TeamId", teamId },
                    { "UninstallerId", message.From.Id },
                };

                this.telemetryClient.TrackEvent("AppUninstalled", properties);

                await this.SaveRemoveFromTeam(teamId, teamsChannelData.Tenant.Id);
            }

            await base.OnMembersRemovedAsync(membersRemoved, turnContext, cancellationToken);
        }

        /// <inheritdoc/>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            await this.HandleMessageActivityAsync(turnContext, cancellationToken);
            await base.OnMessageActivityAsync(turnContext, cancellationToken);
        }

        private async Task HandleMessageActivityAsync(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            try
            {
                var activity = turnContext.Activity;
                var senderAadId = activity.From.AadObjectId;
                var tenantId = activity.GetChannelData<TeamsChannelData>().Tenant.Id;
                var serviceUrl = activity.ServiceUrl;
                var text = (activity.Text ?? string.Empty).Trim();

                if (string.Equals(text, MatchingActions.OptOut, StringComparison.InvariantCultureIgnoreCase))
                {
                    await this.HandleOptOutAsync(turnContext, tenantId, senderAadId, serviceUrl, cancellationToken);
                }
                else if (string.Equals(text, MatchingActions.OptIn, StringComparison.InvariantCultureIgnoreCase))
                {
                    await this.HandleOptInAsync(turnContext, tenantId, senderAadId, serviceUrl, cancellationToken);
                }
                else if (string.Equals(text, MatchingActions.FeedbackYes, StringComparison.InvariantCultureIgnoreCase))
                {
                    await this.HandleFeedbackAsync(turnContext, tenantId, senderAadId, serviceUrl, didMeet: true, cancellationToken);
                }
                else if (string.Equals(text, MatchingActions.FeedbackNo, StringComparison.InvariantCultureIgnoreCase))
                {
                    await this.HandleFeedbackAsync(turnContext, tenantId, senderAadId, serviceUrl, didMeet: false, cancellationToken);
                }
                else if (string.Equals(text, MatchingActions.FrequencyWeekly, StringComparison.InvariantCultureIgnoreCase))
                {
                    await this.HandleFrequencyChangeAsync(turnContext, tenantId, senderAadId, serviceUrl, PairingFrequency.Weekly, cancellationToken);
                }
                else if (string.Equals(text, MatchingActions.FrequencyBiweekly, StringComparison.InvariantCultureIgnoreCase))
                {
                    await this.HandleFrequencyChangeAsync(turnContext, tenantId, senderAadId, serviceUrl, PairingFrequency.Biweekly, cancellationToken);
                }
                else if (string.Equals(text, MatchingActions.FrequencyMonthly, StringComparison.InvariantCultureIgnoreCase))
                {
                    await this.HandleFrequencyChangeAsync(turnContext, tenantId, senderAadId, serviceUrl, PairingFrequency.Monthly, cancellationToken);
                }
                else
                {
                    this.telemetryClient.TrackTrace($"Cannot process the following: {text}");
                    var replyActivity = activity.CreateReply();
                    await this.SendUnrecognizedInputMessageAsync(turnContext, replyActivity, cancellationToken);
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Error while handling message activity: {ex.Message}", SeverityLevel.Warning);
                this.telemetryClient.TrackException(ex);
            }
        }

        private async Task HandleOptOutAsync(ITurnContext turnContext, string tenantId, string senderAadId, string serviceUrl, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"User {senderAadId} opted out");
            this.telemetryClient.TrackEvent("UserOptInStatusSet", new Dictionary<string, string> { { "UserAadId", senderAadId }, { "OptInStatus", "false" } });

            await this.dataProvider.SetUserInfoAsync(tenantId, senderAadId, false, serviceUrl);

            var reply = turnContext.Activity.CreateReply();
            reply.Attachments = new List<Attachment>
            {
                new HeroCard()
                {
                    Text = Resources.OptOutConfirmation,
                    Buttons = new List<CardAction>
                    {
                        new CardAction
                        {
                            Title = Resources.ResumePairingsButtonText,
                            DisplayText = Resources.ResumePairingsButtonText,
                            Type = ActionTypes.MessageBack,
                            Text = MatchingActions.OptIn,
                        },
                    },
                }.ToAttachment(),
            };

            await turnContext.SendActivityAsync(reply, cancellationToken).ConfigureAwait(false);
        }

        private async Task HandleOptInAsync(ITurnContext turnContext, string tenantId, string senderAadId, string serviceUrl, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"User {senderAadId} opted in");
            this.telemetryClient.TrackEvent("UserOptInStatusSet", new Dictionary<string, string> { { "UserAadId", senderAadId }, { "OptInStatus", "true" } });

            await this.dataProvider.SetUserInfoAsync(tenantId, senderAadId, true, serviceUrl);

            var reply = turnContext.Activity.CreateReply();
            reply.Attachments = new List<Attachment>
            {
                new HeroCard()
                {
                    Text = Resources.OptInConfirmation,
                    Buttons = new List<CardAction>
                    {
                        new CardAction
                        {
                            Title = Resources.PausePairingsButtonText,
                            DisplayText = Resources.PausePairingsButtonText,
                            Type = ActionTypes.MessageBack,
                            Text = MatchingActions.OptOut,
                        },
                    },
                }.ToAttachment(),
            };

            await turnContext.SendActivityAsync(reply, cancellationToken).ConfigureAwait(false);
        }

        private async Task HandleFeedbackAsync(ITurnContext turnContext, string tenantId, string senderAadId, string serviceUrl, bool didMeet, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"User {senderAadId} submitted feedback: didMeet={didMeet}");
            this.telemetryClient.TrackEvent("MeetupFeedbackReceived", new Dictionary<string, string>
            {
                { "UserAadId", senderAadId },
                { "DidMeet", didMeet.ToString() },
            });

            // The paired-with userId is passed as the value payload from the card action.
            var pairedWithUserId = turnContext.Activity.Value?.ToString();

            if (!string.IsNullOrEmpty(pairedWithUserId))
            {
                await this.dataProvider.RecordMeetupFeedbackAsync(tenantId, senderAadId, pairedWithUserId, didMeet, serviceUrl);
            }

            var confirmationText = didMeet ? Resources.FeedbackYesConfirmation : Resources.FeedbackNoConfirmation;
            await turnContext.SendActivityAsync(MessageFactory.Text(confirmationText), cancellationToken);
        }

        private async Task HandleFrequencyChangeAsync(ITurnContext turnContext, string tenantId, string senderAadId, string serviceUrl, PairingFrequency frequency, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"User {senderAadId} set frequency to {frequency}");
            this.telemetryClient.TrackEvent("UserFrequencyChanged", new Dictionary<string, string>
            {
                { "UserAadId", senderAadId },
                { "Frequency", frequency.ToString() },
            });

            await this.dataProvider.SetUserFrequencyAsync(tenantId, senderAadId, frequency, serviceUrl);

            string frequencyLabel;
            switch (frequency)
            {
                case PairingFrequency.Biweekly:
                    frequencyLabel = Resources.FrequencyBiweeklyLabel;
                    break;
                case PairingFrequency.Monthly:
                    frequencyLabel = Resources.FrequencyMonthlyLabel;
                    break;
                default:
                    frequencyLabel = Resources.FrequencyWeeklyLabel;
                    break;
            }

            var confirmationText = string.Format(Resources.FrequencySetConfirmation, frequencyLabel);
            await turnContext.SendActivityAsync(MessageFactory.Text(confirmationText), cancellationToken);
        }

        private Task<TeamInstallInfo> GetInstalledTeam(string teamId)
        {
            return this.dataProvider.GetInstalledTeamAsync(teamId);
        }

        private async Task WelcomeUser(ITurnContext turnContext, string memberAddedId, string tenantId, string teamId, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"Sending welcome message for user {memberAddedId}");

            var installedTeam = await this.GetInstalledTeam(teamId);
            var teamName = turnContext.Activity.TeamsGetTeamInfo().Name;
            ChannelAccount userThatJustJoined = await this.conversationHelper.GetMemberAsync(turnContext, memberAddedId, cancellationToken);

            if (userThatJustJoined != null)
            {
                var welcomeMessageCard = WelcomeNewMemberAdaptiveCard.GetCard(teamName, userThatJustJoined.Name, this.botDisplayName, installedTeam.InstallerName);
                await this.conversationHelper.NotifyUserAsync(turnContext, MessageFactory.Attachment(welcomeMessageCard), userThatJustJoined, tenantId, cancellationToken);
            }
            else
            {
                this.telemetryClient.TrackTrace($"Member {memberAddedId} was not found in team {teamId}, skipping welcome message.", SeverityLevel.Warning);
            }
        }

        private async Task WelcomeTeam(ITurnContext turnContext, string botInstaller, CancellationToken cancellationToken)
        {
            var teamId = turnContext.Activity.Conversation.Id;
            this.telemetryClient.TrackTrace($"Sending welcome message for team {teamId}");

            var teamName = turnContext.Activity.TeamsGetTeamInfo().Name;
            var welcomeTeamMessageCard = WelcomeTeamAdaptiveCard.GetCard(teamName, botInstaller);
            await this.NotifyTeamAsync(turnContext, MessageFactory.Attachment(welcomeTeamMessageCard), teamId, cancellationToken);
        }

        private async Task SendUnrecognizedInputMessageAsync(ITurnContext turnContext, Activity replyActivity, CancellationToken cancellationToken)
        {
            replyActivity.Attachments = new List<Attachment> { UnrecognizedInputAdaptiveCard.GetCard() };
            await turnContext.SendActivityAsync(replyActivity, cancellationToken);
        }

        private Task SaveAddedToTeam(string serviceUrl, string teamId, string tenantId, string botInstaller)
        {
            return this.dataProvider.UpdateTeamInstallStatusAsync(new TeamInstallInfo
            {
                ServiceUrl = serviceUrl,
                TeamId = teamId,
                TenantId = tenantId,
                InstallerName = botInstaller,
            }, true);
        }

        private Task SaveRemoveFromTeam(string teamId, string tenantId)
        {
            return this.dataProvider.UpdateTeamInstallStatusAsync(new TeamInstallInfo
            {
                TeamId = teamId,
                TenantId = tenantId,
            }, false);
        }

        private async Task NotifyTeamAsync(ITurnContext turnContext, IMessageActivity activity, string teamId, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"Sending notification to team {teamId}");

            try
            {
                activity.Conversation = new ConversationAccount { Id = teamId };

                var conversationParameters = new ConversationParameters
                {
                    Activity = (Activity)activity,
                    ChannelData = new TeamsChannelData { Channel = new ChannelInfo(teamId) },
                };

                await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                    null,
                    turnContext.Activity.ServiceUrl,
                    this.appCredentials,
                    conversationParameters,
                    null,
                    cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Error sending notification to team: {ex.Message}", SeverityLevel.Warning);
                this.telemetryClient.TrackException(ex);
            }
        }

        private void LogActivityTelemetry(Activity activity)
        {
            var fromObjectId = activity.From?.AadObjectId;
            var clientInfoEntity = activity.Entities?.Where(e => e.Type == "clientInfo")?.FirstOrDefault();
            var channelData = activity.GetChannelData<TeamsChannelData>();

            var properties = new Dictionary<string, string>
            {
                { "ActivityId", activity.Id },
                { "ActivityType", activity.Type },
                { "UserAadObjectId", fromObjectId },
                { "ConversationType", string.IsNullOrWhiteSpace(activity.Conversation?.ConversationType) ? "personal" : activity.Conversation.ConversationType },
                { "ConversationId", activity.Conversation?.Id },
                { "TeamId", channelData?.Team?.Id },
                { "Locale", clientInfoEntity?.Properties["locale"]?.ToString() },
                { "Platform", clientInfoEntity?.Properties["platform"]?.ToString() },
            };
            this.telemetryClient.TrackEvent("UserActivity", properties);
        }

        private bool ValidateTenant(ITurnContext turnContext)
        {
            if (this.disableTenantFilter)
            {
                return true;
            }

            if (this.allowedTenantIds == null || !this.allowedTenantIds.Any())
            {
                throw new ApplicationException("AllowedTenants setting is not set properly in the configuration file.");
            }

            var tenantId = turnContext?.Activity?.Conversation?.TenantId;
            return this.allowedTenantIds.Contains(tenantId);
        }
    }
}
