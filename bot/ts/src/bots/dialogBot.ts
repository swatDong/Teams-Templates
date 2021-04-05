import { TeamsActivityHandler, BotState, TurnContext } from "botbuilder";
import { MainDialog } from "../dialogs/mainDialog";

export class DialogBot extends TeamsActivityHandler {
  conversationState: BotState;
  userState: BotState;
  dialog: MainDialog;
  dialogState: any;

  /**
   *
   * @param {ConversationState} conversationState
   * @param {UserState} userState
   * @param {Dialog} dialog
   */
  constructor(conversationState: BotState, userState: BotState, dialog: MainDialog) {
    super();
    if (!conversationState)
      throw new Error("[DialogBot]: Missing parameter. conversationState is required");
    if (!userState) throw new Error("[DialogBot]: Missing parameter. userState is required");
    if (!dialog) throw new Error("[DialogBot]: Missing parameter. dialog is required");

    this.conversationState = conversationState;
    this.userState = userState;
    this.dialog = dialog;
    this.dialogState = this.conversationState.createProperty("DialogState");

    this.onMessage(async (context, next) => {
      console.log("Running dialog with Message Activity.");

      // Run the Dialog with the new message Activity.
      await this.dialog.run(context, this.dialogState);

      await next();
    });
  }

  /**
   * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
   */
  async run(context: TurnContext) {
    await super.run(context);

    // Save any state changes. The load happened during the execution of the Dialog.
    await this.conversationState.saveChanges(context, false);
    await this.userState.saveChanges(context, false);
  }
}
