document.addEventListener('DOMContentLoaded', async () => {
    const excelInput = document.getElementById('excel_file');
    const docxInput = document.getElementById('docx_file');
    const sendButton = document.getElementById('form_button');
    const logDiv = document.getElementById('form_log');

    function log(message) {
        logDiv.innerHTML += message + '\n';
        logDiv.scrollTop = logDiv.scrollHeight;
    }

    async function setupPyodide() {
        log("\nCarregando ambiente...");
        sendButton.disabled = true;

        const pyodide = await loadPyodide();
        await pyodide.loadPackage(["pandas", "micropip"]);
        const micropip = pyodide.pyimport("micropip");
        await micropip.install("openpyxl");
        
        const extractorCode = await (await fetch('./backend/extractor.py')).text();
        const mainPythonCode = await (await fetch('./bakcend/main.py')).text();
        pyodide.runPython(extractorCode);
        pyodide.runPython(mainPythonCode);

        log("Ambiente pronto!");
        sendButton.disabled = false;
        return pyodide;
    }
    const pyodideReadyPromise = setupPyodide();

    sendButton.addEventListener('click', async () => {
        if (!excelInput.files.length || !docxInput.files.length) {
            log("\nPor favor, selecione os dois arquivos.");
            return;
        }

        sendButton.disabled = true;
        logDiv.innerHTML = '';

        try {
            log("Lendo modelo...");
            const templateFile = docxInput.files[0];
            const templateBuffer = await templateFile.arrayBuffer();

            log("Lendo planilha...");
            const excelFile = excelInput.files[0];
            const excelBuffer = await excelFile.arrayBuffer();
            
            const pyodide = await pyodideReadyPromise;
            const extractFiles = pyodide.globals.get('extractFiles');
            const pythonReturn = await extractFiles(excelBuffer);
            const data = pythonReturn.toJs({ dict_converter: Object.fromEntries });
            pythonReturn.destroy();

            log(`Encontrados ${data.length} registros válidos!`);
            if (data.length === 0) throw new Error("Nenhum dado válido encontrado na planilha.");

            log("Gerando documentos...");
            const zip = new PizZip();
            
            for (const row of data) {
                const doc = new docxtemplater(new PizZip(templateBuffer), {
                    paragraphLoop: true,
                    linebreaks: true,
                });
                doc.setData(row);
                
                try {
                    doc.render();
                    const outputBlob = doc.getZip().generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                    const fileName = `${(row.VAR_NOME || 'documento').toString().replace(/ /g, '_')}.docx`;
                    zip.file(fileName, outputBlob);
                    log(`- Documento gerado para: ${row.VAR_NOME}`);
                } catch (error) {
                    log(`ERRO ao renderizar documento para ${row.VAR_NOME}: ${error.message}`);
                    console.error("Erro do Docxtemplater:", error);
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
        }
    });
});