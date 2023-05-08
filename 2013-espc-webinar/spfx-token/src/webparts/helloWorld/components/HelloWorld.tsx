import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps, IHelloWorldState, ISite } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  AadHttpClient, HttpClientResponse, IHttpClientOptions
} from '@microsoft/sp-http';

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {

  // hardcoded variables for demo purposes only...
  private apiUrl = "https://func-speurope-demo-001.azurewebsites.net/api/GetListItems";
  private aadAuthenticationAppId = "fb43a954-86a4-4f24-b3f6-bdb32346cd3f";

  constructor(props: IHelloWorldProps, state: IHelloWorldState) {
    super(props);
    this.state = {
      items: [],
    };
  }

  private async LetListItems(): Promise<ISite[]> {
    try {
      const sites: ISite[] = [];

       // get the tenant url
      const resourceEndpoint = new URL(this.props.spfxContext.pageContext.site.absoluteUrl);
      const tenantUrl = resourceEndpoint.origin

      // get the token for the current user
      const aadTokenProvider = await this.props.spfxContext.aadTokenProviderFactory.getTokenProvider();
      const token = await aadTokenProvider.getToken(tenantUrl);


      const options: IHttpClientOptions = {
        headers: {
          "Content-Type": "application/json;odata=verbose",
          Accept: "application/json;odata=verbose",
        },
        body: JSON.stringify({
          UserToken: token // this is the token that will be used in the Azure Function
        })
      };

      // get the Azure AD HTTP client to call the Azure Function, passing the Azure AD application registration used to secure the Azure Function
      const client: AadHttpClient = await this.props.spfxContext.aadHttpClientFactory.getClient(this.aadAuthenticationAppId);
      // retrieve the list items from the Azure Function
      const response: HttpClientResponse = await client.post(this.apiUrl, AadHttpClient.configurations.v1, options);
      const responseData: any[] = await response.json();
      
      console.log(responseData);

      if (responseData && responseData?.length >= 0) {
        sites.push(...responseData.map((site) => site as ISite));
      }

      return sites;
    }
    catch (err) {
      console.error('Ups... Something went wrong while getting the hub sites.');
      return null;
    }
  }

  public async componentDidMount(): Promise<void> {
    const listItems = await this.LetListItems();
    this.setState({ items: listItems });
  }



  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
        </div>
        <div>
          {this.state.items.map((item, index) => {
            return (
              <p key={index}>{item.Title}</p>
            );
          })
          }
        </div>
      </section>
    );
  }



}
