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
        
        const extractorCode = await (await fetch('./extractor.py')).text();
        const mainPythonCode = await (await fetch('./main.py')).text();
        pyodide.FS.writeFile("extractor.py", extractorCode, { encoding: "utf8" });
        pyodide.FS.writeFile("main.py", mainPythonCode, { encoding: "utf8" });

        log("Ambiente pronto!");
        sendButton.disabled = false;
        return pyodide;
    }
    const pyodideReadyPromise = setupPyodide();

    sendButton.addEventListener('click', async () => {
        if (!excelInput.files.length || !docxInput.files.length) {
            log("Por favor, selecione os dois arquivos.");
            return;
        }

        sendButton.disabled = true;

        try {
            log("Lendo modelo...");
            const templateFile = docxInput.files[0];
            const templateBuffer = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => resolve(e.target.result);
                reader.onerror = (e) => reject(e.target.error);
                reader.readAsArrayBuffer(templateFile);
            });

            log("Lendo planilha...");
            const excelFile = excelInput.files[0];
            const excelBuffer = await excelFile.arrayBuffer();
            
            const pyodide = await pyodideReadyPromise;

            const mainModule = pyodide.pyimport("main"); 
            const extractFiles = mainModule.extractFiles;
            const pythonReturn = await extractFiles(excelBuffer);
            const data = pythonReturn[0].toJs({ dict_converter: Object.fromEntries });
            const errors = pythonReturn[1].toJs({ dict_converter: Object.fromEntries });
            pythonReturn.destroy();

            log(`Encontrados ${data.length} registros válidos!`);
            if (data.length === 0) throw new Error("Nenhum dado válido encontrado na planilha.");

            log("Gerando documentos...");
            const finalZip = new PizZip();
            const templateZip = new PizZip(templateBuffer);
            console.log(data);
            console.log(errors);
            for (const row of data) {
                const doc = new docxtemplater(templateZip, {
                    paragraphLoop: true,
                    linebreaks: true,
                    delimiters: {
                        start: '{',
                        end: '}'
                    }
                });
                
                doc.setData(row);

                try {
                    doc.render();
                    const outputBlob = doc.getZip().generate({ type: 'blob' });
                    const fileName = `${(row.VAR_NOME || 'documento').toString().replace(/ /g, '_')}.docx`;
                    finalZip.file(fileName, outputBlob);
                } catch (error) {
                    log(`ERRO ao renderizar documento para ${row.VAR_NOME}: ${error.message}`);
                    console.error("Erro do Docxtemplater:", error);
                }
            }

            log("Compactando todos os arquivos...");
            const zipBlob = finalZip.generate({ type: 'blob' });
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