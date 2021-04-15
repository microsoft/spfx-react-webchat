export interface IBotWebPartProps {    
    botButtonText: string;
    chatWindowHeaderTitle: string;
    description: string;
    botAuthenticationType: string;
    botDirectLineSecret: string;
    botDirectLineTokenApiUrl: string;
    botTokenApiResourceId: string;    
    botTokenApiUrl: string;
    botAvatarUrl: string;
    avatarSize: number;
    botAvatarInitials: string;
    backgroundColor: string;
    bubbleBackground: string;
    bubbleBorderRadius: number;
    bubbleFromUserBackground: string;
    bubbleFromUserBorderRadius: string;
    bubbleFromUserTextColor: string;
    suggestedActionBackground: string;
    suggestedActionTextColor: string;
    sendBoxTextWrap: boolean;
    hideScrollToEndButton: boolean;
  }