import * as React from 'react';
import WebViewer from '@pdftron/webviewer';

interface IWebViewerControlProps {
  docUrl: string;
  viewerHeight: number;
  viewerWidth: number;
  onDocSave: (docUrl: string) => void;
}

interface IWebViewerControlState {
  currentDocUrl: string;
}

class WebViewerControl extends React.Component<IWebViewerControlProps, IWebViewerControlState> {
  private viewerElementRef: React.RefObject<HTMLDivElement>;

  constructor(props: IWebViewerControlProps) {
    super(props);
    this.viewerElementRef = React.createRef();
    this.state = {
      currentDocUrl: props.docUrl,
    };
  }

  componentDidMount() {
    const viewerElement = this.viewerElementRef.current;

    if (viewerElement) {
      WebViewer(
        {
          path: 'http://localhost:3000/lib',
          config: 'http://localhost:3000/config.js',
          initialDoc: this.props.docUrl,
        },
        viewerElement
      ).then(instance => {
        const iframeWindow = viewerElement.querySelector('iframe')!.contentWindow!;

        window.addEventListener('message', this.receiveMessage, false);
      });
    }
  }

  componentWillUnmount() {
    window.removeEventListener('message', this.receiveMessage);
  }

  componentDidUpdate(prevProps: IWebViewerControlProps) {
    if (prevProps.docUrl !== this.props.docUrl) {
      this.setState({ currentDocUrl: this.props.docUrl });
      this.handleDocOpen(this.props.docUrl);
    }
  }

  handleDocOpen = (docUrl: string): void => {
    const payload = {
      file: docUrl
    };
    this.viewerElementRef.current?.querySelector('iframe')!.contentWindow!.postMessage({ type: 'OPEN_DOCUMENT', payload }, '*');
  };

  handleDocSave = (docUrl: string): void => {
    this.props.onDocSave(docUrl);
  };

  receiveMessage = (event: any) => {
    if (event.isTrusted && typeof event.data === 'object') {
      switch (event.data.type) {
        case 'SAVE_DOCUMENT':
          this.handleDocSave(event.data.payload.file);
          break;
        default:
          break;
      }
    }
  };

  render(): React.ReactNode {
    return (
      <div ref={this.viewerElementRef} style={{ height: this.props.viewerHeight + 'px', width: this.props.viewerWidth + 'px' }}></div>
    );
  }
}

export default WebViewerControl;

