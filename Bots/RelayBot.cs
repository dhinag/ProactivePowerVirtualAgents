// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Bot.Builder;
using Microsoft.Bot.Connector.DirectLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DirectLineActivity = Microsoft.Bot.Connector.DirectLine.Activity;
using DirectLineActivityTypes = Microsoft.Bot.Connector.DirectLine.ActivityTypes;
using IConversationUpdateActivity = Microsoft.Bot.Schema.IConversationUpdateActivity;
using IMessageActivity = Microsoft.Bot.Schema.IMessageActivity;
using Microsoft.Bot.Schema;
using Newtonsoft.Json;
using RestSharp;


namespace Microsoft.PowerVirtualAgents.Samples.RelayBotSample.Bots
{
    /// <summary>
    /// This IBot implementation shows how to connect
    /// an external Azure Bot Service channel bot (external bot)
    /// to your Power Virtual Agent bot
    /// </summary>
    public class RelayBot : ActivityHandler
    {
        private const int WaitForBotResponseMaxMilSec = 5 * 1000;
        private const int PollForBotResponseIntervalMilSec = 1000;
        private static ConversationManager s_conversationManager = ConversationManager.Instance;
        private ResponseConverter _responseConverter;
        private IBotService _botService;

        public RelayBot(IBotService botService, ConversationManager conversationManager)
        {
            _botService = botService;
            _responseConverter = new ResponseConverter();
        }

        // Invoked when a conversation update activity is received from the external Azure Bot Service channel
        // Start a Power Virtual Agents bot conversation and store the mapping
        protected override async Task OnConversationUpdateActivityAsync(ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            AddConversationReference(turnContext.Activity as Bot.Schema.Activity);
            await s_conversationManager.GetOrCreateBotConversationAsync(turnContext.Activity.Conversation.Id, _botService);
        }

        private void AddConversationReference(Bot.Schema.Activity activity)
        {
            var conversationReference = activity.GetConversationReference();
            var convRefJson = JsonConvert.SerializeObject(conversationReference);
            var response = RunFlow("https://prod-44.westus.logic.azure.com:443/workflows/13b923e5f90/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=T1gSZplRsr-7BUY", convRefJson);

            //Call a flow that will add the Conversation Reference if not already present.
            //_conversationReferences.AddOrUpdate(conversationReference.User.Id, conversationReference, (key, newValue) => conversationReference);
        }

        // Invoked when a message activity is received from the user
        // Send the user message to Power Virtual Agent bot and get response
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var currentConversation = await s_conversationManager.GetOrCreateBotConversationAsync(turnContext.Activity.Conversation.Id, _botService);           

            using (DirectLineClient client = new DirectLineClient(currentConversation.Token))
            {
                // Send user message using directlineClient
                await client.Conversations.PostActivityAsync(currentConversation.ConversationtId, new DirectLineActivity()
                {
                    Type = DirectLineActivityTypes.Message,
                    From = new Bot.Connector.DirectLine.ChannelAccount { Id = turnContext.Activity.From.Id, Name = turnContext.Activity.From.Name },
                    Text = turnContext.Activity.Text,
                    TextFormat = turnContext.Activity.TextFormat,
                    Locale = turnContext.Activity.Locale,
                });

                await RespondPowerVirtualAgentsBotReplyAsync(client, currentConversation, turnContext);
            }

            // Update LastConversationUpdateTime for session management
            currentConversation.LastConversationUpdateTime = DateTime.Now;
        }

        private async Task RespondPowerVirtualAgentsBotReplyAsync(DirectLineClient client, RelayConversation currentConversation, ITurnContext<IMessageActivity> turnContext)
        {
            var retryMax = WaitForBotResponseMaxMilSec / PollForBotResponseIntervalMilSec;
            for (int retry = 0; retry < retryMax; retry++)
            {
                // Get bot response using directlineClient,
                // response contains whole conversation history including user & bot's message
                ActivitySet response = await client.Conversations.GetActivitiesAsync(currentConversation.ConversationtId, currentConversation.WaterMark);

                // Filter bot's reply message from response
                List<DirectLineActivity> botResponses = response?.Activities?.Where(x =>
                      x.Type == DirectLineActivityTypes.Message &&
                        string.Equals(x.From.Name, _botService.GetBotName(), StringComparison.Ordinal)).ToList();

                if (botResponses?.Count() > 0)
                {
                    if (int.Parse(response?.Watermark ?? "0") <= int.Parse(currentConversation.WaterMark ?? "0"))
                    {
                        // means user sends new message, should break previous response poll
                        return;
                    }

                    currentConversation.WaterMark = response.Watermark;
                    await turnContext.SendActivitiesAsync(_responseConverter.ConvertToBotSchemaActivities(botResponses).ToArray());
                }

                Thread.Sleep(PollForBotResponseIntervalMilSec);
            }
        }

        private IRestResponse RunFlow(string flowUrl, string parameter)
        {
            var client = new RestClient(flowUrl);
            var request = new RestRequest(Method.POST);

            request.AddHeader("Connection", "keep-alive");
            request.AddHeader("Accept-Encoding", "gzip, deflate");
            request.AddHeader("Cache-Control", "no-cache");
            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("test", parameter, ParameterType.RequestBody);
            var response = client.Execute(request);
            return response;
        }
    }
}
