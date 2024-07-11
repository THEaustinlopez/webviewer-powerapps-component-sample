import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import WebViewerControl from "./WebViewerControl";

export class ReactWebViewerControl implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private notifyOutputChanged: () => void;
  private pdfURI: string = "";
  private context: ComponentFramework.Context<IInputs>;
  private container: HTMLDivElement;

  constructor() {}

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
}
