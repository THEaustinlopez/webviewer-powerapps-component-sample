import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import WebViewerControl from "./WebViewerControl";

export class ReactWebViewerControl
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private notifyOutputChanged: () => void;
  private pdfURI: string = "";
  private context: ComponentFramework.Context<IInputs>;
  private container: HTMLDivElement;
  private interceptedUrl: string;
  private fileData: {content: string, contentType: string, contentLength: string};
  private fileContent: string;
  private fileContentLength: number;
  private fileContentType: string;
  private fetchingContent: boolean = false;
  private statusMessage: string = "Initializing...";

  constructor() {
    console.log("Constructor: Initializing request intercpetion...");
    this.initializeRequestInterception();
    this.interceptScriptTags();
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary
  ): void {
    console.log("Init: Component Initialized");
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;
    this.container = document.createElement("div");
  }

  public updateView(
    context: ComponentFramework.Context<IInputs>
  ): React.ReactElement {
    console.log("updateView: updating view");
    this.context = context;

    const docUrl = context.parameters.doc.raw!;
    const viewerHeight = context.parameters.viewerheight.raw!;
    const viewerWidth = context.parameters.viewerwidth.raw!;

    if (context.parameters.fileData.raw! && this.fileData != context.parameters.fileData.raw) {
      this.fileData = JSON.parse(context.parameters.fileData.raw);
      console.log('fileData.content', this.fileData.content);
      console.log('fileData.contentType', this.fileData.contentType);
      console.log('fileData.contentLength', this.fileData.contentLength);

      this.fetchingContent = false;
      // this.processNextUrl();
    }

    return React.createElement(WebViewerControl, {
      docUrl,
      viewerHeight,
      viewerWidth,
      onDocSave: this.handleDocSave.bind(this),
    });
  }

  private handleDocSave(docUrl: string): void {
    this.pdfURI = docUrl;
    this.notifyOutputChanged();
  }

  public getOutputs(): IOutputs {
    console.log("interceptedUrl", this.interceptedUrl);
    return {
      pdfdoc: this.pdfURI,
      interceptedUrl: this.interceptedUrl,
    };
  }

  public destroy(): void {
    console.log("destroying component");
    ReactDOM.unmountComponentAtNode(this.container);
  }

  public initializeRequestInterception() {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;

    // Intercepting XMLHttpRequest
    const originalXHROpen = XMLHttpRequest.prototype.open;
    const originalXHRSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function (
      this: XMLHttpRequest,
      method: string,
      url: string | URL,
      async: boolean = true,
      username?: string | null | undefined,
      password?: string | null | undefined
    ): void {
      console.log("XMLHttpRequest.open called with URL: ", url);
      // this._url = url;
      return originalXHROpen.call(this, method, url, async, username, password);
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    XMLHttpRequest.prototype.send = function (
      this: XMLHttpRequest,
      ...args: [body?: Document | XMLHttpRequestBodyInit | null | undefined]
    ) {
      // eslint-disable-next-line @typescript-eslint/no-this-alias
      const xhrInstance = this; // TODO: this or self?
      const url = xhrInstance.responseURL;
      console.log("XMLHttpRequest.send called with URL: ", url);
      if (self.shouldIntercept(url)) {
        console.log(`Intercepted XHR request to: ${url}`);
        xhrInstance.abort();
        self.interceptedUrl = url;
        self.notifyOutputChanged();
        self
          .waitForFileContent()
          .then((responseText) => {
            console.log("responseText", responseText);
            // setTimeout(() => {
            //   console.log("Generating response for intercepted url: ", url);
            //   xhrInstance.readyState = 4;
            //   xhrInstance.status = 200;
            //   xhrInstance.responseText = responseText;
            //   xhrInstance.onreadystatechange &&
            //     xhrInstance.onreadystatechange();
            // }, 0);
          })
          .catch((error) => {
            console.error(
              "[ANALYTICS ERROR]: Error fetching fileContent for because ",
              error
            );
          });
      } else {
        console.log("Not intercepting request to: ", url);
        return originalXHRSend.apply(xhrInstance, args);
      }
    };

    // Intercepting fetch
    const originalFetch = window.fetch;
    window.fetch = async function (
      input: RequestInfo | URL,
      init?: RequestInit
    ): Promise<Response> {
      const url = typeof input === "string" ? input : "";
      console.log("fetch called with URL: ", url);
      if (self.shouldIntercept(url)) {
        console.log(`Intercepted fetch request to: ${url}`);
        self.interceptedUrl = url;
        self.fetchingContent = true;
        self.notifyOutputChanged();
        // version 2 //
        return new Promise((resolve) => {
          self.waitForFileContent().then(response => {
            console.log('response in waitForFileContent.then()', response)
            // const body = self.decodeBase64(response.content, response.contentType);
            const newResponse = new Response(response.content /*body*/, {
              status: 200,
              headers: {
                "Content-Length": response.contentLength.toString(),
                "Content-Type": response.contentType + '; charset=UTF-8',
                // "Accept-Ranges": "bytes",
                "Content-Encoding": "br"
              }
            });
            console.log('newResponse', newResponse);
            resolve(newResponse);

          }).catch(error => {
            console.error(`Error fetching file for ${url}:`, error);
            resolve(new Response('<html><body>Error fetching file content</body></html>', {
              status: 500,
              headers: { 'Content-Type': 'text/html' }
            }));
          });
        });
       
      } else {
        console.log("Not intercepting fetch request to: ", url);
        return originalFetch.apply(self, [input, init]);
      }
    }; /*.bind(this);*/
  }

  private interceptScriptTags() {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;

    const originalCreateElement = document.createElement;
    document.createElement = function(tagName: string, options?: ElementCreationOptions): HTMLElement {
      const element = originalCreateElement.call(document, tagName, options);
      if (tagName.toLowerCase() === 'script') {
        const originalSetAttribute = element.setAttribute;
        element.setAttribute = function(name: string, value: string) {
          if (name === 'src') {
            console.log(`Intercepted script tag with src: ${value}`);
            if (self.shouldIntercept(value)) {
              self.interceptedUrl = value;
              self.notifyOutputChanged();
            }
          }
          return originalSetAttribute.call(this, name, value);
        };
      }
      return element;
    };
  }

  private shouldIntercept(url: string): boolean {
    // Implement logic to determine if the URL should be intercepted
    // For example, intercept all requests to a specific path
    const intercept = url.includes('/public/ui/') || url.includes('/core/webviewer-core.min.js') || url.includes('/ui/webviewer-ui.min.js');
    console.log(`Should intercept ${url}: ${intercept}`);
    return intercept;
  }

  private processNextUrl(): void {
    console.log('processNextUrl(): ', { fileContent: this.fileContent, })
    if (this.fileData) {
      this.clearFileContent();

      // Simulate the response for XMLHttpRequest
      const mockXhr = new MockXMLHttpRequest();
      mockXhr.open("GET", this.interceptedUrl);
      mockXhr.send();
      setTimeout(() => {
        mockXhr.readyState = 4;
        mockXhr.status = 200;
        mockXhr.response = this.fileContent/*decodedContent*/;
        mockXhr.responseText = this.fileContent; // Set base64 content as responseText
        if (mockXhr.onreadystatechange) {
          mockXhr.onreadystatechange();
        }
        mockXhr.setResponseHeaders({
          "Content-Length": this.fileContentLength.toString(),
          "Content-Type": this.fileContentType,
        });
      }, 0);
    }
  }

  private clearFileContent(): void {
    this.fileData = { content: "", contentType: "", contentLength: 0 }    console.log('fileData cleared')
  }

  private waitForFileContent(): Promise<{
    content: string;
    contentLength: number;
    contentType: string;
  }> {
    return new Promise((resolve) => {
      const interval = setInterval(() => {
        if (this.fileData) {
          console.log('waitingForFileContent promisefulfilled with: ', this.fileData)
          clearInterval(interval);
          resolve({
            content: this.fileData.content,
            contentLength: this.fileData.contentLength,
            contentType: this.fileData.contentType,
          });
        }
      }, 100);
    });
  }
}

class MockXMLHttpRequest {
  public readyState: number = 0;
  public status: number = 0;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public response: any = null;
  public responseText: string = "";
  public onreadystatechange: (() => void) | null = null;

  open(method: string, url: string): void {
    console.log(`MockXMLHttpRequest.open called with URL: ${url}`);
  }

  send(): void {
    console.log(`MockXMLHttpRequest.send called`);
  }

  setRequestHeader(name: string, value: string): void {
    console.log(
      `MockXMLHttpRequest.setRequestHeader called with ${name}: ${value}`
    );
  }

  setResponseHeaders(headers: { [key: string]: string }): void {
    for (const [key, value] of Object.entries(headers)) {
      console.log(
        `MockXMLHttpRequest.setResponseHeader called with ${key}: ${value}`
      );
    }
  }
}
