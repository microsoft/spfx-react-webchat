import * as React from 'react';
import { DirectLine } from 'botframework-directlinejs';
import ReactWebChat from 'botframework-webchat';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { escape } from '@microsoft/sp-lodash-subset';
import { AadHttpClient, HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

import styles from './botWebchat.module.scss';
import { IBotProps } from '../contracts/IBotProps';
import { IBotState } from '../contracts/IBotState';
import { IBotToken } from '../contracts/IBotToken';

export default class BotWebchat extends React.Component<IBotProps, IBotState> {
  private directLine: DirectLine;
  constructor(props: IBotProps, state: IBotState) {
    super(props);
    this.setDefaultState();
  }

  public async componentDidMount() {
  }

  public dismissPanel() {
    this.setState({ isOpenPanel: false });
  }

  public render(): React.ReactElement<IBotProps> {    
    return (
      <div className={styles.botWebChat}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <button type="button" className={'btn btn-primary ' + styles.btnPimary} onClick={() => { this.initiateOrRestartBot(); }}>
                 <span>{escape(this.props.botButtonText)}</span>
              </button>
            </div>
            <Panel isOpen={this.state.isOpenPanel} hasCloseButton={false} type={PanelType.custom} isLightDismiss={true} customWidth='480px'>
              <div className={styles.modalDialog} role='document'>
                  <div className={styles.modalContent}>
                    { this.getBotWindowHeader() }
                    <div className={'modalBody ' + styles.modalBody}>
                      {this.state.isInitializing == false && this.directLine != null &&
                        <ReactWebChat directLine={this.directLine} user={{
                            id: this.state.uniqueId.toString(),
                            name: this.state.uniqueId.toString()
                          }}
                          styleOptions={{
                            avatarSize: this.props.avatarSize,
                            botAvatarInitials: this.props.botAvatarInitials,
                            botAvatarImage: escape(this.props.botAvatarUrl),                            
                            backgroundColor: this.props.backgroundColor,
                            bubbleBackground: this.props.bubbleBackground,
                            bubbleBorderRadius: this.props.bubbleBorderRadius,
                            bubbleFromUserBackground: this.props.bubbleFromUserBackground,
                            bubbleFromUserBorderRadius: this.props.bubbleFromUserBorderRadius,
                            bubbleFromUserTextColor: this.props.bubbleFromUserTextColor,
                            suggestedActionBackground: this.props.suggestedActionBackground,
                            suggestedActionTextColor: this.props.suggestedActionTextColor,
                            sendBoxTextWrap:this.props.sendBoxTextWrap,
                            hideScrollToEndButton: this.props.hideScrollToEndButton
                          }}
                        />
                      }
                      {this.state.isInitializing == true && 
                        <div className={styles.spinner}>
                            <p>Loading Bot...</p>
                        </div>
                      }
                    </div>
                  </div>
                </div>
            </Panel>
          </div>
        </div>
      </div>
    );
  }

  private getBotWindowHeader(): JSX.Element {
    return (
        <div className={styles.modalHeader}>
          <span>{escape(this.props.chatWindowHeaderTitle)}</span>
          <button type='button' className={styles.headerIcon} title='Close' onClick={() => this.dismissPanel() } aria-label='Close'>
              <Icon iconName='Cancel' /> 
          </button>
          <button type='button' className={styles.headerIcon} title='Start Over' onClick={() => { this.initiateOrRestartBot(); }} aria-label='Start Over'>
              <Icon iconName='Refresh' /> 
          </button>
          <button type='button' className={styles.headerIcon} title='Feedback' onClick={() => { this.initiateFeedback(); }} aria-label='Feedback'>
              <Icon iconName='Feedback' /> 
          </button>
          <button type='button' className={styles.headerIcon} title='Help' onClick={() => { this.initiateHelp(); }} aria-label='Help'>
              <Icon iconName='Help' />
          </button>
        </div>
      );
  }  

  private setDefaultState() {
    this.state = { botToken:  null, isInitializing: true, isOpenPanel: false, uniqueId: (Math.random()*100), isWelcomeEventPosted: false };
  }

  private async getBotToken(): Promise<IBotToken> {
    const httpClientOptions: IHttpClientOptions = {};
    var client = await this.props.context.aadHttpClientFactory.getClient(this.props.botTokenApiResourceId);
    var response: HttpClientResponse = await client.post(this.props.botTokenApiUrl, AadHttpClient.configurations.v1, httpClientOptions);
    return await response.json();
  }

  private async getBotDirectLineToken(): Promise<IBotToken> {
    const httpClientOptions: IHttpClientOptions = {
      headers: new Headers({ "Authorization" : "Bearer " + this.props.botDirectLineSecret }),
      method: "POST"
    };
    var response: HttpClientResponse = await this.props.context.httpClient.post(this.props.botDirectLineTokenApiUrl, HttpClient.configurations.v1, httpClientOptions);
    return await response.json();
  }

  private async launchBot() {
    var tokenResponse = this.props.botAuthenticationType == "DL-API" ? await this.getBotDirectLineToken() :  this.props.botAuthenticationType == "Custom-API" ? await this.getBotToken() : null;
    this.directLine = this.props.botAuthenticationType == "DL-Secret" ? new DirectLine({ secret: this.props.botDirectLineSecret}) : new DirectLine({
      token: tokenResponse.token,
      webSocket: true
    });
    this.setState({ isInitializing: false, isOpenPanel: true, botToken: tokenResponse, isWelcomeEventPosted: false });
    await this.initiateChatActivity();
  }

  private async initiateOrRestartBot() {
    this.setState({ isInitializing: true, isWelcomeEventPosted: false });
    await this.launchBot();
  }

  private async initiateChatActivity() {
    this.clearConversation(false);
    if (this.state.isInitializing || this.directLine == null) {
      return;
    }

    this.directLine.postActivity({
      from: { id: 'dl_' + this.state.uniqueId.toString() },
      type: 'event',
      name: 'WelcomeEvent',
      value: {
        platform: 'sharepoint',
        pageName: location.href
      }
    })
    .subscribe(
      (id: any) => {
        console.log(`Posted activity, assigned id ${id}`);
        this.setState({ isInitializing: false, isWelcomeEventPosted: true });
        this.clearConversation(true);
      },
      (exception: any) => {
        console.log(`Error posting activity ${exception}`);
      }
    );
  }

  private initiateFeedback() {
    if (this.state.isInitializing || this.directLine == null) {
      return;
    }
    this.directLine.postActivity({
      from: { id: 'dl_' + this.state.uniqueId.toString() },
      type: 'message',
      text: 'Feedback'
    })
    .subscribe(
      (id: any) => {
        console.log(`Posted feedback activity, assigned id ${id}`);
      },
      (exception: any) => {
        console.log(`Error posting feedback activity ${exception}`);
      }
    );
 }

 private initiateHelp() {
  if (this.state.isInitializing || this.directLine == null) {
    return;
  }
  this.directLine.postActivity({
    from: { id: 'dl_' + this.state.uniqueId.toString() },
    type: 'message',
    text: 'help'
  })
  .subscribe(
    (id: any) => {
      console.log(`Posted help activity, assigned id ${id}`);
    },
    (exception: any) => {
      console.log(`Error posting help activity ${exception}`);
    }
  );
}

private clearConversation(keepLastNode: boolean) {
    const chatList = document.querySelectorAll('ul[aria-roledescription="transcript"] > li');
    const lastNode = chatList.length > 0 ? chatList[chatList.length - 1] : null;
    const lastPrevNode = chatList.length > 1 ? chatList[chatList.length - 2] : null;
    chatList.forEach(x => {
      if (keepLastNode && x == lastNode) {
        return;
      }
      if (keepLastNode && x == lastPrevNode) {
        return;
      }
      x.parentNode.removeChild(x);
    });
  }
}
