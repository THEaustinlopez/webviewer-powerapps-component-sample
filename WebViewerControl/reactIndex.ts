import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import WebViewerControl from "./WebViewerControl";

export class ReactWebViewerControl implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private notifyOutputChanged: () => void;
  private pdfURI: string = "";
  private context: ComponentFramework.Context<IInputs>;
  private container: HTMLDivElement;

  constructor() {
    this.initializeRequestInterception();
  }

  public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context; // Store context for later use
    this.container = document.createElement('div');
  }

  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    this.context = context; // Update the context

    const docUrl = context.parameters.doc.raw!;
    const viewerHeight = context.parameters.viewerheight.raw!;
    const viewerWidth = context.parameters.viewerwidth.raw!;

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
      pdfdoc: this.pdfURI,
    };
  }

  public destroy(): void {
    // Add code to cleanup control if necessary
    ReactDOM.unmountComponentAtNode(this.container);
  }

  private initializeRequestInterception() {
    // Intercepting XMLHttpRequest
    const originalXHROpen = XMLHttpRequest.prototype.open;
    const originalXHRSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function(this: XMLHttpRequest, method: string, url: string, ...rest: any[]) {
      (this as any)._url = url; // Store the URL for later use
      return originalXHROpen.apply(this, [method, url, ...rest]);
    };

    XMLHttpRequest.prototype.send = function(this: XMLHttpRequest, ...args: any[]) {
      const xhrInstance = this;
      const url = (xhrInstance as any)._url;
      if (shouldIntercept(url)) {
        console.log(`Intercepted XHR request to: ${url}`);
        xhrInstance.abort(); // Optionally abort the original request
        setTimeout(() => {
          xhrInstance.readyState = 4;
          xhrInstance.status = 200;
          xhrInstance.responseText = '<html><body>Intercepted Content</body></html>';
          xhrInstance.onreadystatechange && xhrInstance.onreadystatechange();
        }, 0);
      } else {
        return originalXHRSend.apply(this, args);
      }
    };

    // Intercepting fetch
    const originalFetch = window.fetch;
    window.fetch = function(...args: [RequestInfo, RequestInit?]): Promise<Response> {
      const [url, options] = args;
      if (shouldIntercept(url.toString())) {
        console.log(`Intercepted fetch request to: ${url}`);
        return new Promise((resolve) => {
          resolve(new Response('<html><body>Intercepted Content</body></html>', {
            status: 200,
            headers: { 'Content-Type': 'text/html' }
          }));
        });
      } else {
        return originalFetch.apply(this, args);
      }
    }.bind(this);

    // Helper functions
    function shouldIntercept(url: string): boolean {
      // Add your logic to determine if the request should be intercepted
      return url.includes('public/ui/index.html');
    }

    function getCachedResponse(url: string): string | null {
      const cachedData = localStorage.getItem(url);
      return cachedData ? JSON.parse(cachedData).response : null;
    }

    function cacheResponse(url: string, response: string): void {
      const cachedData = {
        response: response,
        timestamp: Date.now()
      };
      localStorage.setItem(url, JSON.stringify(cachedData));
    }
  }
}

