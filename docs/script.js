document.addEventListener('DOMContentLoaded', () => {

    const excelInput = document.getElementById('excel_file');
    const docxInput = document.getElementById('docx_file');
    const sendButton = document.getElementById('form_button');
    const logDiv = document.getElementById('form_log');

    function log(message) {
        logDiv.innerHTML += '\n' + message;
        logDiv.scrollTop = logDiv.scrollHeight;
    }

    sendButton.addEventListener('click', async () => {
        if (!excelInput.files.length || !docxInput.files.length) {
            log("Por favor, selecione os dois arquivos.");
            return;
        }

        sendButton.disabled = true;
        sendButton.textContent = "Processando...";
        logDiv.innerHTML = '';

        try {
            log("Lendo modelo...");
            const templateFile = docxInput.files[0];
            const templateBuffer = await templateFile.arrayBuffer();
            log("Modelo lido com sucesso!");

            log("Lendo planilha...");
            const excelFile = excelInput.files[0];
            const excelBuffer = await excelFile.arrayBuffer();
            const workbook = XLSX.read(excelBuffer, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            log(`Encontrados ${data.length} registros!`);

            log("Iniciando geração de documentos...");
            const zip = new PizZip();
            
            for (const row of data) {
                const doc = new docxtemplater(new PizZip(templateBuffer), {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                doc.setData(row);

                try {
                    doc.render();
                    log(`- Documento gerado para: ${row.VAR_NOME || 'Registro sem nome'}`);
                    const outputBlob = doc.getZip().generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                    
                    const fileName = `${(row.VAR_NOME || 'documento').toString().replace(/ /g, '_')}.docx`;
                    zip.file(fileName, outputBlob);

                } catch (error) {
                    log(`ERRO ao renderizar o documento para ${row.VAR_NOME}: ${error.message}`);
                }
            }

            log("Compactando todos os arquivos...");
            const zipBlob = zip.generate({ type: 'blob' });
            saveAs(zipBlob, 'documentos_gerados.zip');
            log("Download iniciado!");

        } catch (error) {
            log(`ERRO GERAL: ${error.message}`);
            console.error(error);

        } finally {
            sendButton.disabled = false;
            sendButton.textContent = "Gerar Documentos";
        }
    });
});