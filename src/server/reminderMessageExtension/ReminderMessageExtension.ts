import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult, TaskModuleRequest, TaskModuleContinueResponse } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/reminderMessageExtension/config.html")
@PreventIframe("/reminderMessageExtension/action.html")
export default class ReminderMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onFetchTask(context: TurnContext, value: MessagingExtensionQuery): Promise<MessagingExtensionResult | TaskModuleContinueResponse> {

        if (!value.state) { // TODO: implement logic when config is persisted
            return Promise.resolve<MessagingExtensionResult>({
                type: "config", // use "config" or "auth" here
                suggestedActions: {
                    actions: [
                        {
                            type: "openUrl",
                            value: `https://${process.env.PUBLIC_HOSTNAME}/reminderMessageExtension/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`,
                            title: "Configuration"
                        }
                    ]
                }
            });
        }

        return Promise.resolve<TaskModuleContinueResponse>({
            type: "continue",
            value: {
                title: "Input form",
                card: CardFactory.adaptiveCard({
                    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                    type: "AdaptiveCard",
                    version: "1.4",
                    body: [
                        {
                            type: "TextBlock",
                            text: "Please enter an e-mail address"
                        },
                        {
                            type: "Input.Text",
                            id: "email",
                            placeholder: "somemail@example.com",
                            style: "email"
                        },
                        {
                            type: "ActionSet",
                            actions: [
                                {
                                    type: "Action.Execute",
                                    title: "OK",
                                    data: { id: "unique-id" },
                                    fallback: "Action.Submit"
                                }
                            ]
                        }
                    ]
                })
            }
        });
    }

    // handle action response in here
    // See documentation for `MessagingExtensionResult` for details
    public async onSubmitAction(context: TurnContext, value: TaskModuleRequest): Promise<MessagingExtensionResult> {

        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: value.data.email
                    },
                    {
                        type: "Image",
                        url: `https://randomuser.me/api/portraits/thumb/women/${Math.round(Math.random() * 100)}.jpg`
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.4"
            });
        return Promise.resolve({
            type: "result",
            attachmentLayout: "list",
            attachments: [card]
        } as MessagingExtensionResult);
    }

    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
        return Promise.resolve({
            title: "Reminder Message Extension Configuration",
            value: `https://${process.env.PUBLIC_HOSTNAME}/reminderMessageExtension/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`
        });
    }

    public async onSettings(context: TurnContext): Promise<void> {
        // take care of the setting returned from the dialog, with the value stored in state
        const setting = context.activity.value.state;
        log(`New setting: ${setting}`);
        return Promise.resolve();
    }

}
