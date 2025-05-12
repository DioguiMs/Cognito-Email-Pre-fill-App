import {
    BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import {
    Version
} from '@microsoft/sp-core-library';

import {
    MSGraphClientV3
} from '@microsoft/sp-http';

export interface ICognitoIntegrationAppWebPartProps {
    formUrl: string;    // URL of the Cognito form to embed
    emailFieldName?: string;    // name of the field in the form for email
    nameFieldName?: string;   // name of the field in the form for name
    locationFieldName?: string;  // name of the field in the form for location
    cssCode?: string;
}


export default class CognitoIntegrationAppWebPart extends BaseClientSideWebPart<ICognitoIntegrationAppWebPartProps> {

    public async render(): Promise<void> {
    
        // Getting User Email
        const userEmail = this.context.pageContext.user.email;
        const userName = this.context.pageContext.user.displayName;        
        const userLocation = await this.getUserLocation();
        
        const entryData: Record<string, string> = {};
        if (this.properties.emailFieldName?.trim() && userEmail) {
            entryData[this.properties.emailFieldName.trim()] = userEmail;
        }

        if (this.properties.nameFieldName?.trim() && userName) {
            entryData[this.properties.nameFieldName.trim()] = userName;
        }

        if (this.properties.locationFieldName?.trim() && userLocation) {
            entryData[this.properties.locationFieldName.trim()] = userLocation;
        }

        let formUrl = this.properties.formUrl;
        if (formUrl && Object.keys(entryData).length > 0) {
            const entryJson = JSON.stringify(entryData);
            formUrl += `?entry=${entryJson}`
        }
        console.log("Form URL: ", formUrl);
    
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

        if (this.properties.cssCode?.trim()) {
            const css = this.properties.cssCode.trim();
            const escapedCss = css.replace(/'/g, "\\'"); // Escape single quotes if any

            // Create the style element for the CSS
            const style = document.createElement('style');
            style.type = 'text/css';
            style.innerHTML = escapedCss;

            // Wait for the iframe to load
            const iframe = document.querySelector('iframe');
            if (iframe) {
                iframe.onload = () => {
                    // Access the iframe's document and append the style
                    const iframeDocument = iframe.contentWindow?.document;
                    if (iframeDocument) {
                        iframeDocument.head.appendChild(style);
                    }
                };
            }
    }
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
                        label: "Cognito Form URL"
                    }),
                    PropertyPaneTextField('emailFieldName', {
                        label: "Email field name in form (optional)"
                    }),
                    PropertyPaneTextField('nameFieldName', {
                        label: "Name field name in form (optional)"
                    }),
                    PropertyPaneTextField('locationFieldName', {
                        label: "Location field name in form (optional)"
                    }),
                    PropertyPaneTextField('cssCode', {
                        label: "CSS Code (optional)",
                        multiline: true,
                        description: "Custom CSS code to style the form. Use with caution."
                    })
                    ]
                }
                ]
            }
            ]
        };
    }

    public getUserLocation(): Promise<string | null> {
        return this.context.msGraphClientFactory
            .getClient('3')
            .then((client: MSGraphClientV3) => {
            return client
                .api('/me')
                .select('officeLocation') // lowercase 'officeLocation' is correct
                .get()
                .then((user) => {
                console.log('User office location:', user.officeLocation);
                return user.officeLocation || null;
                });
            })
            .catch((error) => {
            console.error('Error getting MSGraphClient or location:', error);
            return null;
        });
    }

    
    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }
}
