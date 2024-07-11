import * as React from 'react';
import WebViewer from '@pdftron/webviewer';

interface IWebViewerControlProps {
  docUrl: string;
  viewerHeight: number;
  viewerWidth: number;
  onDocSave: (docUrl: string) => void;
}

const WebViewerControl: React.FC<IWebViewerControlProps> = ({ docUrl, viewerHeight, viewerWidth, onDocSave }) => {
  const viewerElementRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const viewerElement = viewerElementRef.current;

    if (viewerElement) {
      WebViewer(
        {
          path: 'http://localhost:3000/lib',
          config: 'http://localhost:3000/config.js',
          initialDoc: docUrl,
        },
        viewerElement
      ).then(instance => {
        const iframeWindow = viewerElement.querySelector('iframe')!.contentWindow!;

        window.addEventListener('message', receiveMessage, false);

        function receiveMessage(event: any) {
          if (event.isTrusted && typeof event.data === 'object') {
            switch (event.data.type) {
              case 'SAVE_DOCUMENT':
                onDocSave(event.data.payload.file);
                break;
              default:
                break;
            }
          }
        }
      });
    }

    return () => {
      window.removeEventListener('message', receiveMessage);
    };
  }, [docUrl]);

  return <div ref={viewerElementRef} style={{ height: viewerHeight + 'px', width: viewerWidth + 'px' }}></div>;
};

export default WebViewerControl;
