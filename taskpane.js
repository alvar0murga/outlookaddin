// taskpane.js
function downloadEmailAsPDF() {
  alert('El botón funciona');
Office.onReady(() => {
  // Initialization if needed
});

function downloadEmailAsPDF() {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // Use jsPDF from window.jspdf.jsPDF
      const jsPDF = window.jspdf.jsPDF;
      const doc = new jsPDF();
      doc.setFontSize(12);
      doc.text("Email content:", 10, 10);
      doc.text(result.value, 10, 20);
      doc.save('email.pdf');
    } else {
      alert('No se pudo obtener el contenido del email.');
    }
  });
}

window.downloadEmailAsPDF = downloadEmailAsPDF;
