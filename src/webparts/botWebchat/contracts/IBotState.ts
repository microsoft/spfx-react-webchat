import { IBotToken } from "./IBotToken";
export interface IBotState {
  botToken: IBotToken;
  isInitializing: boolean;
  isWelcomeEventPosted: boolean;
  isOpenPanel: boolean;
  uniqueId: number;
}