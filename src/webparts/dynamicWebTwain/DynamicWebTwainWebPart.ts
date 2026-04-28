import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './DynamicWebTwainWebPart.module.scss';
import * as strings from 'DynamicWebTwainWebPartStrings';

declare let Dynamsoft: any;

export interface IDynamicWebTwainWebPartProps {
  description: string;
}

export default class DynamicWebTwainWebPart extends BaseClientSideWebPart<IDynamicWebTwainWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _DWTObject: any = null;
  private _dwtReady: boolean = false;
  private _domReady: boolean = false;

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  public render(): void {
    if (this._domReady) {
      return;
    }

    const containerId = `dwt-${this.instanceId}`;

    this.domElement.innerHTML = `
    <section class="${styles.dynamicWebTwain} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.controls}">
        <input type="button" value="Scan" id="scan-${this.instanceId}" class="${styles.button}" disabled />
        <input type="button" value="Upload" id="upload-${this.instanceId}" class="${styles.button}" disabled />
        <input type="button" value="Convert to binary image" id="binarize-${this.instanceId}" class="${styles.button}" disabled />
        <input type="button" value="Rotate clockwise" id="rotateCW-${this.instanceId}" class="${styles.button}" disabled />
        <input type="button" value="Rotate counter-clockwise" id="rotateCCW-${this.instanceId}" class="${styles.button}" disabled />
      </div>
      <div id="${containerId}" class="${styles.container}"></div>
    </section>`;

    this._bindEvents();
    this._initDWT(containerId).catch(err => console.error(err));
    this._domReady = true;
  }

  private _setButtonsEnabled(enabled: boolean): void {
    const buttons = this.domElement.querySelectorAll<HTMLInputElement>('input[type="button"]');
    buttons.forEach(btn => { btn.disabled = !enabled; });
  }

  private _bindEvents(): void {
    this.domElement.querySelector(`#scan-${this.instanceId}`)
      ?.addEventListener('click', () => this._acquireImage());
    this.domElement.querySelector(`#upload-${this.instanceId}`)
      ?.addEventListener('click', () => this._upload());
    this.domElement.querySelector(`#binarize-${this.instanceId}`)
      ?.addEventListener('click', () => this._binarizeImage());
    this.domElement.querySelector(`#rotateCW-${this.instanceId}`)
      ?.addEventListener('click', () => this._rotateCW());
    this.domElement.querySelector(`#rotateCCW-${this.instanceId}`)
      ?.addEventListener('click', () => this._rotateCCW());
  }

  private _loadDWTScript(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      if ((window as any).Dynamsoft) {
        resolve();
        return;
      }
      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/dwt/dist/dynamsoft.webtwain.min.js';
      script.onload = () => resolve();
      script.onerror = () => reject(new Error('Failed to load Dynamic Web TWAIN script'));
      document.head.appendChild(script);
    });
  }

  private async _initDWT(containerId: string): Promise<void> {
    if (this._dwtReady) {
      return;
    }

    try {
      await this._loadDWTScript();

      Dynamsoft.DWT.AutoLoad = false;
      Dynamsoft.DWT.IfCheckCssFiles = false;
      Dynamsoft.DWT.Containers = [{
        ContainerId: containerId,
        Width: '100%',
        Height: '400px'
      }];
      Dynamsoft.DWT.ProductKey = 'DLS2eyJvcmdhbml6YXRpb25JRCI6IjIwMDAwMSJ9';
      Dynamsoft.DWT.ResourcesPath = 'https://cdn.jsdelivr.net/npm/dwt/dist';
      Dynamsoft.DWT.ServiceInstallerLocation = 'https://demo.dynamsoft.com/DWT/Resources/dist/';

      Dynamsoft.DWT.RegisterEvent('OnWebTwainReady', () => {
        this._DWTObject = Dynamsoft.DWT.GetWebTwain(containerId);
        this._dwtReady = true;
        this._setButtonsEnabled(true);
      });

      Dynamsoft.DWT.Load();
    } catch (error) {
      console.error('Failed to initialize Dynamic Web TWAIN:', error);
    }
  }

  private _acquireImage(): void {
    if (this._DWTObject) {
      this._DWTObject.SelectSourceAsync()
        .then(() => {
          return this._DWTObject.AcquireImageAsync({
            IfCloseSourceAfterAcquire: true,
            IfShowUI: false,
            PixelType: Dynamsoft.DWT.EnumDWT_PixelType.TWPT_GRAY,
            Resolution: 150,
          });
        })
        .catch((exp: any) => {
          alert(exp.message);
        });
    }
  }

  private _upload(): void {
    if (this._DWTObject && this._DWTObject.HowManyImagesInBuffer > 0) {
      const strUrl = 'https://demo.dynamsoft.com/sample-uploads/';
      const imgAry = [this._DWTObject.CurrentImageIndexInBuffer];
      this._DWTObject.HTTPUpload(
        strUrl,
        imgAry,
        Dynamsoft.DWT.EnumDWT_ImageType.IT_PNG,
        Dynamsoft.DWT.EnumDWT_UploadDataFormat.Binary,
        'WebTWAINImage.png',
        () => { alert('Upload successful'); },
        (_errorCode: any, errorString: string, sHttpResponse: string) => {
          alert(sHttpResponse.length > 0 ? sHttpResponse : errorString);
        }
      );
    } else {
      alert('There is no image in buffer.');
    }
  }

  private _binarizeImage(): void {
    if (this._DWTObject) {
      this._DWTObject.ConvertToBW(this._DWTObject.CurrentImageIndexInBuffer);
    }
  }

  private _rotateCW(): void {
    if (this._DWTObject) {
      this._DWTObject.RotateRight(this._DWTObject.CurrentImageIndexInBuffer);
    }
  }

  private _rotateCCW(): void {
    if (this._DWTObject) {
      this._DWTObject.RotateLeft(this._DWTObject.CurrentImageIndexInBuffer);
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
