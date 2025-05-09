import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField
  } from '@microsoft/sp-webpart-base';
  
  export interface ICognitoFormEmbedWebPartProps {
    formUrl: string; // URL of the Cognito form to embed
  }
  
  export default class CognitoFormEmbedWebPart extends BaseClientSideWebPart<ICognitoFormEmbedWebPartProps> {
  
    public render(): void {
  
      // Getting User Email
      const userEmail = this.context.pageContext.user.email;
  
      // URL with pre-filling
      const formUrl = `${this.properties.formUrl}?entry={"UserEmail":"${userEmail}"}`;
  
      const iframe = document.createElement('iframe');
      iframe.src = formUrl;
      iframe.allow = 'payment';
      iframe.style.border = '0';
      iframe.style.width = '100%';
      iframe.height = '3410';
  
      this.domElement.innerHTML = '';
      this.domElement.appendChild(iframe);
  
      const script = document.createElement('script');
      script.src = 'https://www.cognitoforms.com/f/iframe.js';
      this.domElement.appendChild(script);
    }
  
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      return {
        pages: [
          {
            header: {
              description: "Cognito Form Embed Settings"
            },
            groups: [
              {
                groupName: "Form Configuration",
                groupFields: [
                  PropertyPaneTextField('formUrl', {
                    label: "Cognito Form Embed URL"
                  })
                ]
              }
            ]
          }
        ]
      };
    }
  }
  