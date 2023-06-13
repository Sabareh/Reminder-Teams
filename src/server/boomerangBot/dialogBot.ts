import {
    ConversationState,
    UserState,
    TeamsActivityHandler,
    TurnContext,
    StatePropertyAccessor
} from "botbuilder";
import { MainDialog } from "./dialogs/mainDialog";
import { DialogSet } from "botbuilder-dialogs";

export class DialogBot extends TeamsActivityHandler {
    private conversationState: ConversationState;
    private userState: UserState;
    private dialog: MainDialog;
    private dialogState: StatePropertyAccessor<any>;

    constructor(conversationState: ConversationState, userState: UserState, dialog: MainDialog) {
        super();
        this.conversationState = conversationState;
        this.userState = userState;
        this.dialog = dialog;
        this.dialogState = this.conversationState.createProperty("DialogState");

        this.onMessage(async (context: TurnContext, next: () => Promise<void>) => {
            const dialogContext = await this.dialog.createContext(context);

            // Continue the dialog if it's not done
            if (!dialogContext.context.responded) {
                await dialogContext.continueDialog();
            }

            // Start the dialog if it hasn't been started yet
            if (!context.responded) {
                await dialogContext.beginDialog(this.dialog.id);
            }

            await next();
        });
    }

    public async run(context: TurnContext): Promise<void> {
        await super.run(context);
        await this.conversationState.saveChanges(context, false);
        await this.userState.saveChanges(context, false);
    }
}
