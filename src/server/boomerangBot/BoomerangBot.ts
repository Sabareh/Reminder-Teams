import { TurnContext, ActivityTypes, BotFrameworkAdapter, CardFactory, ConversationState, MemoryStorage, UserState } from "botbuilder";
import { AdaptiveCard } from "adaptivecards";

import { DialogBot } from "./dialogBot";
import { MainDialog } from "./dialogs/mainDialog";
import WelcomeCard from "./cards/welcomeCard";
import ReminderMessageExtension from "../reminderMessageExtension/ReminderMessageExtension";
import { MessageExtensionDeclaration, BotDeclaration, PreventIframe } from "express-msteams-host";

@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD
)
@PreventIframe("/boomerangBot/aboutBoomerang.html")
@MessageExtensionDeclaration("reminderMessageExtension")
export class BoomerangBot extends DialogBot {
    private adapter: BotFrameworkAdapter;
    private reminderMessageExtension: ReminderMessageExtension;

    constructor(conversationState: ConversationState, userState: UserState) {
        super(conversationState, userState, new MainDialog());
        this.adapter = new BotFrameworkAdapter();
        this.reminderMessageExtension = new ReminderMessageExtension();

        this.onMessage(async (context: TurnContext) => {
            const messageId = context.activity.id;
            const conversationId = context.activity.conversation.id;

            const replied = await this.checkMessageReplied(messageId);
            const replyLater = await this.checkMessageReplyLater(context);

            if (!replied && !replyLater) {
                await this.sendReminder(context, messageId, conversationId);
            }

            await super.onMessage(context);
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            if (membersAdded && membersAdded.length > 0) {
                for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                    if (membersAdded[cnt].id !== context.activity.recipient.id) {
                        await this.sendWelcomeCard(context);
                    }
                }
            }
            await next();
        });
    }

    private async checkMessageReplied(messageId: string): Promise<boolean> {
    // Implement logic to check if the message has been replied to
    // You can use external storage or database to track message replies

        // For demonstration purposes, assume the message has been replied to
        return true;
    }

    private async checkMessageReplyLater(context: TurnContext): Promise<boolean> {
    // Implement logic to check if the message has been marked for "reply later"
    // You can use user or conversation state to track the "reply later" status

        // For demonstration purposes, assume the message has been marked for "reply later"
        return true;
    }

    private async sendReminder(context: TurnContext, messageId: string, conversationId: string) {
        const reminderMessage = `You have an unread message: ${messageId}`;

        // Create an Adaptive Card for the reminder message
        const adaptiveCard = new AdaptiveCard();
        adaptiveCard.addTextBlock(reminderMessage);

        // Convert the Adaptive Card to an attachment
        const cardAttachment = CardFactory.adaptiveCard(adaptiveCard);

        // Send the reminder message to the user or channel
        const reminderActivity = {
            type: ActivityTypes.Message,
            attachments: [cardAttachment],
            conversation: { id: conversationId },
            recipient: { id: context.activity.from.id }
        };

        await this.adapter.continueConversation(conversationId, async (turnContext) => {
            await turnContext.sendActivity(reminderActivity);
        });
    }

    public async sendWelcomeCard(context: TurnContext): Promise<void> {
        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
        await context.sendActivity({ attachments: [welcomeCard] });
    }
}
