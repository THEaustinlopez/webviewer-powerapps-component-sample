import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import WebViewerControl from "./WebViewerControl";

interface FileData {
  content: string;
  contentType: string;
  contentLength: string;
}

export class ReactWebViewerControl implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private notifyOutputChanged: () => void;
  private context: ComponentFramework.Context<IInputs>;
  private container: HTMLDivElement;
  private interceptedUrl: string | null = null;
  private fileData: FileData | null = null;
  private urlQueue: string[] = [];
  private isProcessingQueue: boolean = false;
  private statusMessage: string = "Initializing...";

  constructor() {
    this.container = document.createElement("div");
    this.initializeRequestInterception();
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;
    console.log("Component Initialized");
  }

  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    this.context = context;

    const docUrl = context.parameters.doc.raw!;
    const viewerHeight = context.parameters.viewerheight.raw!;
    const viewerWidth = context.parameters.viewerwidth.raw!;

    // Check if fileData is already a JavaScript object
    if (context.parameters.fileData.raw) {
        console.log('Raw fileData:', context.parameters.fileData.raw);

        if (typeof context.parameters.fileData.raw === 'string') {
            try {
                const parsedFileData = JSON.parse(context.parameters.fileData.raw);
                if (this.fileData != parsedFileData) {
                    this.fileData = parsedFileData;
                    console.log('Parsed fileData:', this.fileData);
                    this.fetchingContent = false;
                    this.processNextUrl();
                }
            } catch (error) {
                console.error('Error parsing fileData:', error);
                console.error('fileData.raw content:', context.parameters.fileData.raw);
            }
        } else if (typeof context.parameters.fileData.raw === 'object') {
            this.fileData = context.parameters.fileData.raw as FileData;
            console.log('Assigned fileData directly:', this.fileData);
            this.fetchingContent = false;
            this.processNextUrl();
        }
    }

    return React.createElement(WebViewerControl, {
        docUrl,
        viewerHeight,
        viewerWidth,
        onDocSave: this.handleDocSave.bind(this),
    });
}


  private handleDocSave(docUrl: string): void {
    this.notifyOutputChanged();
    console.log("Document saved:", docUrl);
  }

  public getOutputs(): IOutputs {
    return {
      interceptedUrl: this.interceptedUrl,
    };
  }

  public destroy(): void {
    ReactDOM.unmountComponentAtNode(this.container);
    console.log("Component destroyed.");
  }

  private initializeRequestInterception() {
    const self = this;

    // Intercepting XMLHttpRequest
    const originalXHROpen = XMLHttpRequest.prototype.open;
    const originalXHRSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function (this: XMLHttpRequest, method: string, url: string | URL, async: boolean = true, username?: string | null, password?: string | null): void {
      this._url = url.toString();
      return originalXHROpen.call(this, method, url, async, username, password);
    };

    XMLHttpRequest.prototype.send = function (this: XMLHttpRequest, body?: Document | XMLHttpRequestBodyInit | null) {
      const url = this._url;
      if (self.shouldIntercept(url)) {
        console.log(`Intercepted XHR request to: ${url}`);
        self.enqueueUrl(url);
      } else {
        return originalXHRSend.apply(this, arguments);
      }
    };

    // Intercepting fetch
    const originalFetch = window.fetch;
    window.fetch = function (input: RequestInfo | URL, init?: RequestInit): Promise<Response> {
      const url = typeof input === "string" ? input : input.url;
      if (self.shouldIntercept(url)) {
        console.log(`Intercepted fetch request to: ${url}`);
        self.enqueueUrl(url);
        return new Promise((resolve) => {
          self.waitForFileContent().then(response => {
            resolve(new Response(response.content, {
              status: 200,
              headers: {
                "Content-Length": response.contentLength,
                "Content-Type": response.contentType,
              }
            }));
          }).catch(error => {
            console.error(`Error fetching file for ${url}:`, error);
            resolve(new Response('<html><body>Error fetching file content</body></html>', {
              status: 500,
              headers: { 'Content-Type': 'text/html' }
            }));
          });
        });
      } else {
        return originalFetch.apply(this, arguments);
      }
    };

    this.interceptScriptTags();
  }

  private interceptScriptTags() {
    const self = this;
    const originalCreateElement = document.createElement;
    document.createElement = function (tagName: string, options?: ElementCreationOptions): HTMLElement {
      const element = originalCreateElement.call(document, tagName, options);
      if (tagName.toLowerCase() === 'script') {
        const originalSetAttribute = element.setAttribute;
        element.setAttribute = function (name: string, value: string) {
          if (name === 'src') {
            if (self.shouldIntercept(value)) {
              self.enqueueUrl(value);
            }
          }
          return originalSetAttribute.call(this, name, value);
        };
      }
      return element;
    };
  }

  private shouldIntercept(url: string): boolean {
    return url.includes('/public/ui/') || url.includes('/core/webviewer-core.min.js') || url.includes('/ui/webviewer-ui.min.js');
  }

  private enqueueUrl(url: string): void {
    this.urlQueue.push(url);
    if (!this.isProcessingQueue) {
      this.processNextUrl();
    }
  }

  private processNextUrl(): void {
    if (this.urlQueue.length === 0) {
      this.isProcessingQueue = false;
      return;
    }

    this.isProcessingQueue = true;
    this.interceptedUrl = this.urlQueue.shift()!;
    this.notifyOutputChanged();

    // wait for file content and process
    this.waitForFileContent().then(response => {
      const mockXhr = new MockXMLHttpRequest();
      mockXhr.open("GET", this.interceptedUrl);
      setTimeout(() => {
        mockXhr.readyState = 4;
        mockXhr.status = 200;
        mockXhr.response = response.content;
        mockXhr.responseText = response.content;
        if (mockXhr.onreadystatechange) {
          mockXhr.onreadystatechange();
        }
        mockXhr.setResponseHeaders({
          "Content-Length": response.contentLength,
          "Content-Type": response.contentType,
        });
        this.processNextUrl(); // Move to the next URL
      }, 0);
    }).catch(error => {
      console.error(`Error processing URL ${this.interceptedUrl}:`, error);
      this.processNextUrl(); // Continue with the next URL even if there's an error
    });
  }

  private waitForFileContent(): Promise<FileData> {
    return new Promise((resolve, reject) => {
      const interval = setInterval(() => {
        if (this.fileData) {
          clearInterval(interval);
          resolve(this.fileData);
        }
      }, 100);
    });
  }

  private clearFileContent(): void {
    this.fileData = null;
  }
}

class MockXMLHttpRequest {
  public readyState: number = 0;
  public status: number = 0;
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
    console.log(`MockXMLHttpRequest.setRequestHeader called with ${name}: ${value}`);
  }

  setResponseHeaders(headers: { [key: string]: string }): void {
    for (const [key, value] of Object.entries(headers)) {
      console.log(`MockXMLHttpRequest.setResponseHeader called with ${key}: ${value}`);
    }
  }
}
