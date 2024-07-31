import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import WebViewerControl from "./WebViewerControl";

interface PowerAutomateResponse {
  content: string; // base64 encoded content
  contentLength: number;
  contentType: string;
}

export class ReactWebViewerControl implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private notifyOutputChanged: () => void;
  private pdfURI: string = "";
  private context: ComponentFramework.Context<IInputs>;
  private container: HTMLDivElement;
  private interceptedUrl: string = "";
  private fileContent: string = ""; // base64 encoded content
  private contentLength: number = 0;
  private contentType: string = "";
  private fetchingContent: boolean = false;
  private statusMessage: string = "Initializing...";

  constructor() {
    this.container = document.createElement('div');
  }

  public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context; // Store context for later use
    this.initializeRequestInterception();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    this.context = context; // Update the context

    const docUrl = context.parameters.doc.raw!;
    const viewerHeight = context.parameters.viewerheight.raw!;
    const viewerWidth = context.parameters.viewerwidth.raw!;

    if (context.parameters.fileContent.raw) {
      this.fileContent = context.parameters.fileContent.raw;
      this.fetchingContent = false;
      this.processNextUrl();
    }

    if (context.parameters.contentLength.raw) {
      this.contentLength = context.parameters.contentLength.raw;
    }

    if (context.parameters.contentType.raw) {
      this.contentType = context.parameters.contentType.raw;
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
    return {
      interceptedUrl: this.interceptedUrl,
    };
  }

  public destroy(): void {
    ReactDOM.unmountComponentAtNode(this.container);
  }

  private initializeRequestInterception() {
    const self = this;

    // Intercepting XMLHttpRequest
    const originalXHROpen = XMLHttpRequest.prototype.open;
    const originalXHRSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function(this: XMLHttpRequest, method: string, url: string, async?: boolean, username?: string | null, password?: string | null) {
      console.log(`XMLHttpRequest.open called with URL: ${url}`);
      this._url = url; // Store the URL for later use
      return originalXHROpen.apply(this, [method, url, async, username, password]);
    };

    XMLHttpRequest.prototype.send = function(this: XMLHttpRequest, body?: Document | XMLHttpRequestBodyInit | null) {
      const xhrInstance = this;
      const url = xhrInstance._url!;
      console.log(`XMLHttpRequest.send called with URL: ${url}`);
      if (self.shouldIntercept(url)) {
        console.log(`Intercepted XHR request to: ${url}`);
        xhrInstance.abort(); // Optionally abort the original request
        self.interceptedUrl = url;
        self.notifyOutputChanged();
      } else {
        console.log(`Not intercepting request to: ${url}`);
        return originalXHRSend.apply(xhrInstance, [body]);
      }
    };

    // Intercepting fetch
    const originalFetch = window.fetch;
    window.fetch = function(input: RequestInfo, init?: RequestInit): Promise<Response> {
      const url = typeof input === 'string' ? input : input.url;
      console.log(`fetch called with URL: ${url}`);
      if (self.shouldIntercept(url)) {
        console.log(`Intercepted fetch request to: ${url}`);
        self.interceptedUrl = url;
        self.notifyOutputChanged();
        return new Promise((resolve) => {
          self.waitForFileContent().then(response => {
            const body = self.decodeBase64(response.content, response.contentType);
            resolve(new Response(body, {
              status: 200,
              headers: {
                "Content-Length": response.contentLength.toString(),
                "Content-Type": response.contentType,
                "Accept-Ranges": "bytes",
                "Content-Encoding": "br"
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
        console.log(`Not intercepting fetch request to: ${url}`);
        return originalFetch.apply(window, [input, init]);
      }
    };
  }

  private shouldIntercept(url: string): boolean {
    // Implement logic to determine if the URL should be intercepted
    // For example, intercept all requests to a specific path
    return url.includes('/public/ui/');
  }

  private processNextUrl(): void {
    if (this.fileContent && this.contentLength && this.contentType) {
      const decodedContent = this.decodeBase64(this.fileContent, this.contentType);
      this.clearFileContent();

      // Simulate the response for XMLHttpRequest
      const mockXhr = new MockXMLHttpRequest();
      mockXhr.open("GET", this.interceptedUrl);
      mockXhr.send();
      setTimeout(() => {
        mockXhr.readyState = 4;
        mockXhr.status = 200;
        mockXhr.response = decodedContent;
        mockXhr.responseText = this.fileContent; // Set base64 content as responseText
        if (mockXhr.onreadystatechange) {
          mockXhr.onreadystatechange();
        }
        mockXhr.setResponseHeaders({
          "Content-Length": this.contentLength.toString(),
          "Content-Type": this.contentType
        });
      }, 0);
    }
  }

  private clearFileContent(): void {
    this.fileContent = "";
    this.contentLength = 0;
    this.contentType = "";
  }

  private waitForFileContent(): Promise<{ content: string; contentLength: number; contentType: string }> {
    return new Promise((resolve) => {
      const interval = setInterval(() => {
        if (this.fileContent && this.contentLength && this.contentType) {
          clearInterval(interval);
          resolve({
            content: this.fileContent,
            contentLength: this.contentLength,
            contentType: this.contentType
          });
        }
      }, 100);
    });
  }

  private decodeBase64(base64: string, contentType: string): Blob | ArrayBuffer {
    const binaryString = atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }

    if (contentType.startsWith('application/pdf') || contentType.startsWith('application/octet-stream') || contentType.startsWith('application/wasm')) {
      return bytes.buffer;
    } else {
      return new Blob([bytes], { type: contentType });
    }
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
      console.log(`Mock
