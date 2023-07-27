import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionActionResponse,
  MessagingExtensionAction,
  Activity,
  MessageFactory,
} from "botbuilder";

function adaptiveCard(submit: boolean){
  return CardFactory.adaptiveCard({
    actions: submit ? [
      {
        type: "Action.Submit",
        title: "Submit",
        data: { submitLocation: "messagingExtensionSubmit" },
      },
    ] : [],
    body: [{ text: "Adaptive Card Test", type: "TextBlock", weight: "bolder" }],
    type: "AdaptiveCard",
    version: "1.0",
  });
}

export class TeamsME extends TeamsActivityHandler {
  protected handleTeamsMessagingExtensionFetchTask(
    _context: TurnContext,
    _action: MessagingExtensionAction
  ): Promise<MessagingExtensionActionResponse> {
    return Promise.resolve({
      task: {
        type: "continue",
        value: {
          card: adaptiveCard(true),
          height: "medium",
          width: "medium",
          title: "Adaptive Card",
        },
      },
    });
  }

  protected handleTeamsMessagingExtensionSubmitAction(_context: TurnContext, _action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
      return Promise.resolve({
        composeExtension: {
            type: "botMessagePreview",
            activityPreview: MessageFactory.attachment(adaptiveCard(false)) as any as Activity,
        }
      });
  }

  protected handleTeamsMessagingExtensionBotMessagePreviewEdit(_context: TurnContext, _action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
    return Promise.resolve({
        task: {
          type: "continue",
          value: {
            card: adaptiveCard(true),
            height: "medium",
            width: "medium",
            title: "Adaptive Card",
          },
        },
      });
  }

  protected async handleTeamsMessagingExtensionBotMessagePreviewSend(_context: TurnContext, _action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
    await _context.sendActivity(MessageFactory.attachment(adaptiveCard(false)));
    return {};
  }
}
