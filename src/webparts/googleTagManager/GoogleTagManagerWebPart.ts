import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'GoogleTagManagerWebPartStrings';
import { IGoogleTagManagerWebPartProps } from './IGoogleTagManagerWebPartProps';

// explitcitly declare Window or it'll error on gtm
declare global {
  interface Window { dataLayer: any; _spPageContextInfo: any;}
}

var getEmployeeId = function () {
  if (window && window._spPageContextInfo) {
      return window._spPageContextInfo.userLoginName.split('@')[0];
  }
  return null;
}

// Google Tag Manager dataLayer
window.dataLayer = window.dataLayer || [];
window.dataLayer.push({
  'userId': getEmployeeId()
});

var getEnv = function () {

    if (window.location.href.indexOf("bahdev.sharepoint.com") !== -1) {
        return "dev";
    }
    if (window.location.href.indexOf("bahtest.sharepoint.com") !== -1) {
        return "test";
    } else {
        return "prod";
    }
}
  // Google Tag Manager environment
  var env = getEnv();
  var gtm_auth, gtm_preview;
  switch (env) {
      case ('dev'):
          gtm_auth = '99tPzWvHXnuCa-8Z3QYMhg';
          gtm_preview = 'env-5';
          break;
      case ('test'):
          gtm_auth = 'chnucPjxDu7lXnSAKWy8Uw';
          gtm_preview = 'env-6';
          break;
      default:
          gtm_auth = 'kzJPJUywQJ2EbQF6h6Ej_Q';
          gtm_preview = 'env-7';
          break;
  }

  // Google Tag Manager script
  (function (w, d, s, l, i, a, p) {
      w[l] = w[l] || []; w[l].push({
          'gtm.start':
          new Date().getTime(), event: 'gtm.js'
      }); var f = d.getElementsByTagName(s)[0],
      // TS is typesafe, so you'll get an error that src is not a property of the HTMLElement type
      j = <HTMLScriptElement>d.createElement(s), dl = l != 'dataLayer' ? '&l=' + l : ''; j.async = true; j.src =
              'https://www.googletagmanager.com/gtm.js?id=' + i + dl + '&gtm_auth=' + a + '&gtm_preview=' + p + '&gtm_cookies_win=x'; f.parentNode.insertBefore(j, f);
  })(window, document, 'script', 'dataLayer', 'GTM-WD5V8S3', gtm_auth, gtm_preview);

export default class GoogleTagManagerWebPartWebPart extends BaseClientSideWebPart<IGoogleTagManagerWebPartProps> {

  public render(): void {
    // display nothing
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
