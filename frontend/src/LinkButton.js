import React from 'react'

export default function LinkButton({ downloadUrl }) {
    const download = url => {
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = true;
        anchor.target = "_blank";
        anchor.rel = "noreferrer"
    
        anchor.click();
      }
    
      return (
        <button onClick={() => download(downloadUrl)}>Download</button>
      );
}
