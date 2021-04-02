import { BotState, SigninStateVerificationQuery, TurnContext } from "botbuilder";
import { MainDialog } from "../dialogs/mainDialog";
import { DialogBot } from "./dialogBot";

export class TeamsBot extends DialogBot {
  /**
   *
   * @param {ConversationState} conversationState
   * @param {UserState} userState
   * @param {Dialog} dialog
   */
  constructor(conversationState: BotState, userState: BotState, dialog: MainDialog) {
    super(conversationState, userState, dialog);

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      if (membersAdded) {
        for (let cnt = 0; cnt < membersAdded.length; cnt++) {
          if (membersAdded[cnt].id !== context.activity.recipient.id) {
            await context.sendActivity(
              "Welcome to TeamsBot. Type anything to get logged in. Type 'logout' to sign-out."
            );
          }
        }
      }

      await next();
    });
  }

  async handleTeamsSigninVerifyState(context: TurnContext, query: SigninStateVerificationQuery) {
    await this.dialog.run(context, this.dialogState);
  }

  async handleTeamsSigninTokenExchange(
    context: TurnContext, // really need this?
    query: SigninStateVerificationQuery
  ) {
    await this.dialog.run(context, this.dialogState);
  }
}
