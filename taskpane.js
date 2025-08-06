// taskpane.js

// Espera a que Office esté listo
Office.onReady(() => {
  // Puedes poner aquí código de inicialización si lo necesitas
});

// Función para descargar el email como PDF
function downloadEmailAsPDF() {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // Crea el PDF
      const { jsPDF } = window.jspdf;
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

// Haz que la función esté disponible globalmente
window.downloadEmailAsPDF = downloadEmailAsPDF;
