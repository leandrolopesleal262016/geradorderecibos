document.getElementById('receiptForm').addEventListener('submit', function(event) {
    event.preventDefault();
    const name = document.getElementById('name').value;
    const amount = document.getElementById('amount').value;

    axios.post('/generate_receipt', {
        name: name,
        amount: amount
    }, {
        responseType: 'blob'
    })
    .then(response => {
        const pdfBlob = new Blob([response.data], { type: 'application/pdf' });
        const pdfUrl = URL.createObjectURL(pdfBlob);
        document.getElementById('pdfEmbed').setAttribute('src', pdfUrl);
        document.getElementById('pdfViewer').classList.remove('hidden');
    })
    .catch(error => {
        console.error('Erro ao gerar recibo:', error);
    });
});

document.getElementById('csvFile').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const formData = new FormData();
    formData.append('file', file);

    axios.post('/upload_csv', formData, {
        headers: {
            'Content-Type': 'multipart/form-data'
        }
    })
    .then(response => {
        console.log('Nomes carregados:', response.data);
        // Aqui você pode adicionar a lógica para gerar recibos para cada nome
    })
    .catch(error => {
        console.error('Erro ao carregar CSV:', error);
    });
});