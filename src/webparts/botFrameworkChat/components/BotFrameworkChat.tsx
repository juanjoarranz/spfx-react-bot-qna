import * as React from 'react';
import { css, hiddenContentStyle } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import styles from './BotFrameworkChat.module.scss';
import { IBotFrameworkChatProps } from './IBotFrameworkChatProps';
import * as showdown from 'showdown';
import './botstyles.css';

declare function require( path: string ): any;


const BotSpinner = ( props: { show: boolean } ) => {
  return (
    <div style={{ display: props.show ? 'block' : 'none' }}>
      <div style={{ content: "", clear: 'both', display: 'table' }}></div>
      <Spinner size={SpinnerSize.xSmall} style={{ position: 'absolute', bottom: -15, right: 25 }} />
    </div>
  );
};

export default class BotFrameworkChat extends React.Component<IBotFrameworkChatProps, any> {

  private pollInterval = 1000;
  private directLineClient;
  private clientSwagger;
  private conversationId;
  private messagesHtml;
  private currentMessageToSend;
  private sendAsUserName;
  private displaySpinner = false;
  private conversationUpdateEventText = 'conversationUpdate event detected';
  private botInitialized: boolean = false;

  constructor( props: IBotFrameworkChatProps, context ) {
    super( props );
    this.state = {
      resolvedError  : false,
      resolvedSuccess: false,
      message        : null,
      error          : ''
    };

    this.sendMessage = this.sendMessage.bind( this );
    this.restart = this.restart.bind( this );
  }

  public render(): JSX.Element {

    if ( this.state.resolvedError ) {
      return (
        <div className={styles.botFrameworkChat}>
          <div className={styles.container}>
            <div className={css( 'ms-Grid-rowZ ms-font-xl', styles.chatHeader )} style={{ backgroundColor: this.props.titleBarBackgroundColor }} >
              {this.props.title}
            </div>

            <div className={css( 'ms-Grid-rowZ' )}>
              <h3 style={{color: 'red'}}>Error: Invalid Direct Line Secret in the web part properties</h3>
            </div>
          </div>
        </div>
      );
    }

    if ( this.state.resolvedSuccess ) {

      let displayRefreshIcon = !this.messagesHtml || this.messagesHtml === '' ? 'none' : 'block';

      return (
        <div className={styles.botFrameworkChat}>
          <div className={styles.container} style={{ borderColor: this.props.titleBarBackgroundColor }}>
            <div className={css( 'ms-font-xl', styles.chatHeader )} style={{ backgroundColor: this.props.titleBarBackgroundColor, position: 'relative' }} >
              {this.props.title}
              <div className="refresh-icon" onClick={this.restart} style={{ display: displayRefreshIcon }} >
                <TooltipHost content="Restart" id="refresh-icon-tooltip" calloutProps={{ gapSpace: 0 }}>
                  <Icon iconName="Refresh"/>
                </TooltipHost>
              </div>
            </div>

            <div className={styles.messagesRow} style={{ height: this.props.messagesRowHeight }}>
              <div style={{position: 'relative'}}>
                <div className='ms-Grid-col ms-u-sm12' ref='messageHistoryDiv' dangerouslySetInnerHTML={{ __html: this.getMessagesHtml() }}></div>
                <BotSpinner show={this.displaySpinner}/>
              </div>
            </div>

            <div className={css( 'bot-inputbox-row' )}>
              <TextField id='MessageBox' onKeyUp={( e ) => this.tbKeyUp( e )} onKeyDown={( e ) => this.tbKeyDown( e )}
                value={this.currentMessageToSend} placeholder={this.props.placeholderText} className={css( 'ms-fontSize-m', styles.messageBox )} />
              <div className={css( 'bot-sendbutton')} onClick={this.sendMessage}>
                <svg height="28" viewBox="0 0 45.7 33.8" width="28"><path d="M8.55 25.25l21.67-7.25H11zm2.41-9.47h19.26l-21.67-7.23zm-6 13l4-11.9L5 5l35.7 11.9z" fill="#8e8d8c" clip-rule="evenodd"></path></svg>
              </div>
            </div>
          </div>
        </div>
      );
    }

    if ( this.props.directLineSecret === '' ) {
      return (
        <div className={styles.botFrameworkChat}>
          <div className={styles.container}>
            <div className={css( 'ms-Grid-rowZ ms-font-xl', styles.chatHeader )} style={{ backgroundColor: this.props.titleBarBackgroundColor }} >
              {this.props.title}
            </div>
            <div className={css( 'ms-Grid-rowZ' )}>
              <h3>Enter the Bot Direct Line Secret in the web part properties</h3>
            </div>
          </div>
        </div>
      );
    }

    return <h6></h6>;
  }

  public componentDidMount(): void {
    console.log( 'component mounted' );
    if ( this.props.directLineSecret ) {
      if ( !this.directLineClient ) {
        this._initClientSwagger();

      } else {
        this.setState( {
          resolvedSuccess: true
        } );
      }
    }
  }

  public componentDidUpdate( prevProps: IBotFrameworkChatProps, prevState: {}, prevContext: any ): void {
    console.log( 'component updated' );
    console.log( 'previous prevProps.directLineSecret', prevProps.directLineSecret );
    console.log( 'curren directLineSecret', this.props.directLineSecret );

    if ( this.props.directLineSecret && this.props.directLineSecret !== prevProps.directLineSecret) {
        this._initClientSwagger();
    }
  }

  private _initClientSwagger() {
    this._getClientSwagger()
      .then( client => {
        client.Conversations.Conversations_NewConversation()
          .then( ( response ) => response.obj.conversationId )
          .then( ( conversationId ) => {
            this.conversationId = conversationId;
            this.pollMessages( client, conversationId );
            this.directLineClient = client;

            this.sendAsUserName = this.props.context.pageContext.user.loginName;
            this.printMessage = this.printMessage.bind( this );

            this.setState( {
              resolvedError: false,
              resolvedSuccess: true
            } );

          } )
          .catch( error => {
            console.log( 'Client Swagger Creation Error', error );
            this.directLineClient = null;
            this.setState( {
              resolvedSuccess: false,
              resolvedError: true,
              error: error.obj.error.message
            } );
          } );
      } )
      .catch( error => { } );
  }

  private _getClientSwagger(): Promise<any> {
    var Swagger = require( 'swagger-client' );
    var directLineSpec = require( './directline-swagger.json' );

    return new Promise( ( resolve, reject ) => {
      this.clientSwagger = new Swagger( {
          spec: directLineSpec,
          usePromise: true,
        } )
        .then( ( client ) => {
          client.clientAuthorizations.add( 'AuthorizationBotConnector', new Swagger.ApiKeyAuthorization( 'Authorization', 'BotConnector ' + this.props.directLineSecret, 'header' ) );
          console.log( 'DirectLine client generated' );
          resolve( client );
        } )
        .catch( ( err ) => {
          console.error( 'Error initializing DirectLine client', err );
          reject( err );
      } );
    } );
  }

  public getMessagesHtml() {
    return this.messagesHtml;
  }

  public tbKeyUp( e ) {
    this.currentMessageToSend = e.target.value;
    this.forceMessagesContainerScroll();
  }

  public tbKeyDown( e ) {
    if ( e.keyCode === 13 ) {
      this.sendMessage();
    }
  }

  private sendMessage() {
    if ( this.currentMessageToSend ) {
      this.displaySpinner = true;
      let messageToSend: string = this.currentMessageToSend;

      this.currentMessageToSend = '';

      this.setState( {
        message: ''
      } );

      if ( !this.messagesHtml ) {
        this.messagesHtml = '';
      }

      let chatTimeHtml = this.props.displayChatTime ?
        `<div class="${styles.timestamp}" style="right:0; text-align: right">You at ${new Date().toLocaleTimeString().replace( /:\d{2}$/, '' )}</div>` : '';

      this.messagesHtml +=
        `<span class = "${ styles.message } ${ styles.fromUser } ms-fontSize-mPlus"
          style="background-color: ${this.props.userMessagesBackgroundColor }; color: ${ this.props.userMessagesForegroundColor }">
            ${messageToSend }
            <div class="${styles.calloutArrow }" style="right: -6px; top: 12px; background-color: ${ this.props.userMessagesBackgroundColor }"></div>
            ${chatTimeHtml}
        </span>`;

      this.directLineClient.Conversations.Conversations_PostMessage( {
        conversationId: this.conversationId,
        message: {
          from: this.sendAsUserName,
          text: messageToSend
        } } )
        .catch( ( err ) => console.error( 'Error sending message:', err ) );
    }
  }

  protected pollMessages( client, conversationId ) {
    console.log( 'Starting polling message for conversationId: ' + conversationId );
    var watermark = null;
    setInterval( () => {
      client.Conversations.Conversations_GetMessages( { conversationId: conversationId, watermark: watermark } )
        .then( ( response ) => {
          watermark = response.obj.watermark;
          return response.obj.messages;
        } )
        .then( ( messages ) => this.printMessages( messages ) );
    }, this.pollInterval );
  }

  protected printMessages( messages ) {
    if ( messages && messages.length ) {
      messages = messages.filter( ( m ) => m.from !== this.sendAsUserName );
      if ( messages.length ) {
        messages.forEach( this.printMessage );
      }
    }
  }

  protected printMessage( message ) {

    if ( !this.botInitialized && this.props.botWelcomeMessage !== undefined && this.props.botWelcomeMessage !== '' && message.text === this.conversationUpdateEventText) {
      this.displayReceivedMessage( this.props.botWelcomeMessage );
    } else {
      if ( message.text && message.text !== this.conversationUpdateEventText ) {
        this.displayReceivedMessage( message.text );
      }
    }
  }

  private displayReceivedMessage( messageText ) {
    this.setState( {
      message: messageText
    } );

    if ( !this.messagesHtml ) {
      this.messagesHtml = '';
    }

    let converter = new showdown.Converter();
    let answerHtml: string = converter.makeHtml( messageText );

    // get all links that are not html-ready yet and convert them to html
    let regex: RegExp = /[^"'](https:\/\/[^\s]+)/g;
    if ( regex.test( answerHtml ) ) {
      let linksToHtml = answerHtml.match( regex ).map( s => s.replace( /[^h]*/, '' ) );
      linksToHtml.forEach( link => answerHtml = answerHtml.replace( link, `<a href="${ link }">${ link }</a>` ) );
    }

    let chatTimeHtml = this.props.displayChatTime ?
      `<div class="${ styles.timestamp }" style="left:0">${ this.props.title } at ${ new Date().toLocaleTimeString().replace( /:\d{2}$/, '' ) }</div>` : '';

    this.messagesHtml +=
      `<span class = "${ styles.message } ${ styles.fromBot } ms-fontSize-m"
        style="background-color: ${this.props.botMessagesBackgroundColor }; color: ${ this.props.botMessagesForegroundColor }">
          ${answerHtml }
          <div class="${styles.calloutArrow }" style="left: -8px; top: 15px; background-color: ${ this.props.botMessagesBackgroundColor }"></div>
          ${chatTimeHtml}
      </span>`;

    this.displaySpinner = false;
    this.botInitialized = true;
    this.forceUpdate();
    this.forceMessagesContainerScroll();
  }

  protected forceMessagesContainerScroll() {
    var messagesRowClass = '.' + styles.messagesRow;
    var messagesDivElement = document.querySelector( messagesRowClass );
    messagesDivElement.scrollTop = messagesDivElement.scrollHeight;
  }

  private restart() {
    this.messagesHtml = '';
    if ( this.props.botWelcomeMessage && this.props.botWelcomeMessage !== '' ) {
      this.displayReceivedMessage( this.props.botWelcomeMessage );
    } else {
      this.forceUpdate();
      this.forceMessagesContainerScroll();
    }
  }
}
