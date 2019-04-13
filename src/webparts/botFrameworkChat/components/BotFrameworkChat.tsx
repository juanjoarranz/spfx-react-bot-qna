import * as React from 'react';
import { css, hiddenContentStyle } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react';
import styles from './BotFrameworkChat.module.scss';
import { IBotFrameworkChatProps } from './IBotFrameworkChatProps';
import * as showdown from 'showdown';
import './botstyles.css';

declare function require( path: string ): any;

export default class BotFrameworkChat extends React.Component<IBotFrameworkChatProps, any> {

  private pollInterval = 1000;
  private directLineClient;
  private conversationId;
  private clientSwagger;
  private messagesHtml;
  private currentMessageToSend;
  private sendAsUserName;
  private conversationUpdateEventText = 'conversationUpdate event detected';

  constructor( props, context ) {
    super( props );
    this.state = {
      resolvedError  : false,
      resolvedSuccess: false,
      message        : null,
      error          : ''
    };
  }

  public render(): JSX.Element {

    if ( this.state.resolveError ) {
      return (
        <div className={styles.botFrameworkChat}>
          <div className={styles.container}>
            <div className={css( 'ms-Grid-rowZ ms-font-xl', styles.chatHeader )} style={{ backgroundColor: this.props.titleBarBackgroundColor }} >
              {this.props.title}
            </div>

            <div className={css( 'ms-Grid-rowZ' )}>
              <h3>Error Encountered:{this.state.error} </h3>
            </div>
          </div>
        </div>
      );

    }

    if ( this.state.resolvedSuccess ) {
      return (
        <div className={styles.botFrameworkChat}>
          <div className={styles.container} style={{ borderColor: this.props.titleBarBackgroundColor }}>
            <div className={css( 'ms-Grid-rowZ ms-font-xl', styles.chatHeader )} style={{ backgroundColor: this.props.titleBarBackgroundColor }} >
              {this.props.title}
            </div>
            <div className={css( 'ms-Grid-rowZ', styles.messagesRow )} style={{ height: this.props.messagesRowHeight }}>
              <div className='ms-Grid-col ms-u-sm12' ref='messageHistoryDiv' dangerouslySetInnerHTML={{ __html: this.getMessagesHtml() }}>
              </div>
            </div>
            <div className={css( 'bot-inputbox-row' )} style={{position: 'relative', borderTopColor: '#d8d8d8'}}>
              <TextField id='MessageBox' onKeyUp={( e ) => this.tbKeyUp( e )} onKeyDown={( e ) => this.tbKeyDown( e )}
                value={this.currentMessageToSend} placeholder={this.props.placeholderText} className={css( 'ms-fontSize-m', styles.messageBox )} />
              <div style={{ height: 28, width: 28, overflow: 'hidden', padding: 6, position: 'absolute', top: -4, right: 0 }}>
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
      if ( !this.clientSwagger ) {
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
    if ( this.props.directLineSecret ) {
      if ( !this.clientSwagger ) {
        this._initClientSwagger();
      }
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
          } );

        this.sendAsUserName = this.props.context.pageContext.user.loginName;
        this.printMessage = this.printMessage.bind( this );

        this.clientSwagger = client;
        this.setState( {
          resolvedSuccess: true
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
      var messageToSend = this.currentMessageToSend;

      this.currentMessageToSend = '';

      this.setState( {
        message: '',
      } );

      if ( !this.messagesHtml ) {
        this.messagesHtml = '';
      }

      this.messagesHtml = this.messagesHtml + ' <span class="' + styles.message
        + ' ' + styles.fromUser + '  ms-fontSize-mPlus" style="background-color:' + this.props.userMessagesBackgroundColor
        + '; color:' + this.props.userMessagesForegroundColor + '">' + e.target.value + '</span> ';

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
    // console.log( 'message.text = ', message.text );
    // console.log( 'this.state.message = ', this.state.message );
    // console.log( 'received message', message );

    if ( message.text && message.text !== this.conversationUpdateEventText ) {
      this.setState( {
        message: message.text
      } );

      if ( !this.messagesHtml ) {
        this.messagesHtml = '';
      }

      let answerHtml: string = message.text.replace( /\n/g, '<br/>' );
      let converter = new showdown.Converter();
      answerHtml = converter.makeHtml( answerHtml );

      // get all links that are not html-ready yet and convert them to html
      let regex: RegExp = /[^"'](https:\/\/[^\s]+)/g;
      if ( regex.test( answerHtml ) ) {
        let linksToHtml = answerHtml.match( regex ).map( s => s.replace( /[^h]*/, '' ) );
        linksToHtml.forEach( link => answerHtml = answerHtml.replace( link, `<a href="${ link }">${ link }</a>` ) );
      }

      this.messagesHtml +=
        `<span class = "${ styles.message } ${ styles.fromBot } ms-fontSize-m"
          style="background-color: ${this.props.botMessagesBackgroundColor }; color: ${ this.props.botMessagesForegroundColor }">
            ${answerHtml }
        </span>`;

      this.forceUpdate();

      this.forceMessagesContainerScroll();
    }
  }

  protected forceMessagesContainerScroll() {
    var messagesRowClass = '.' + styles.messagesRow;
    var messagesDivElement = document.querySelector( messagesRowClass );
    messagesDivElement.scrollTop = messagesDivElement.scrollHeight;
  }

}
