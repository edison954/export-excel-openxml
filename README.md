# export-excel-openxml

Test en net.core 2.1
pruebas del servicio en angular:

Export() {

    const requestUrl = 'https://localhost:44341/api/Excel/Export';
    const fileName = 'report-openXML.xlsx';
    fetch(requestUrl, {
      method: 'POST'
    })
      .then(function(response) {
        const blob = response.blob();
        return blob;
      })
      .then(blob => {
        DownloadFile(
          blob,
          fileName,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
      });
}
  
public DownloadFile(data: Blob, filename: string, mime: string): void {

    const blob = new Blob([data], { type: mime || 'application/octet-stream' });
    if (typeof window.navigator.msSaveBlob !== 'undefined') {
      window.navigator.msSaveBlob(blob, filename);
    } else {
      const blobURL = window.URL.createObjectURL(blob);
      const tempLink = document.createElement('a');
      tempLink.href = blobURL;
      tempLink.setAttribute('download', filename);
      tempLink.setAttribute('target', '_blank');
      document.body.appendChild(tempLink);
      tempLink.click();
      document.body.removeChild(tempLink);
    }
		
}
