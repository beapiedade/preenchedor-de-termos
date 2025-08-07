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
        sendButton.textContent = "Carregando";

        const pyodide = await loadPyodide();
        await pyodide.loadPackage(["pandas", "micropip", "lxml"]);

        const micropip = pyodide.pyimport("micropip");
        await micropip.install(["openpyxl", "python-docx", "docxtpl"]);

        const extractorCode = await (await fetch('./extractor.py')).text();
        const factoryCode = await (await fetch('./factory.py')).text();
        const mainPythonCode = await (await fetch('./main_python.py')).text();
        pyodide.runPython(extractorCode);
        pyodide.runPython(factoryCode);
        pyodide.runPython(mainPythonCode);

        log("Ambiente pronto!");
        sendButton.disabled = false;
        sendButton.textContent = "Gerar Documentos";
            
        return pyodide;
    }

    const pyodideReadyPromise = setupPyodide();

    sendButton.addEventListener('click', async () => {
        if (!excelInput.files.length || !docxInput.files.length) {
            log("\nPor favor, selecione os dois arquivos.");
            return;
        }

        sendButton.disabled = true;
        log("Iniciando processamento...");

        try {
            const pyodide = await pyodideReadyPromise;

            const excelFile = excelInput.files[0];
            const excelBuffer = await excelFile.arrayBuffer();
            const templateFile = docxInput.files[0];
            const templateBuffer = await templateFile.arrayBuffer();
            
            log("Enviando arquivos para o ambiente Python...");
            
            const generate = pyodide.globals.get("generate");
            const generatedData = await generate(excelBuffer, templateBuffer);

            const finalData = generatedData.get('arquivos').toJs();
            const nomesComErro = generatedData.get('erros').toJs();
            generatedData.destroy(); // Limpa a memória

            log("Processamento finalizado!");

            if (finalData.size > 0) {
                log("Compactando arquivos gerados...");
                const zip = new PizZip();
                for (const [fileName, fileContent] of finalData) {
                    log(`- Adicionando: ${fileName}`);
                    zip.file(fileName, fileContent);
                }
                const zipBlob = zip.generate({ type: 'blob' });
                saveAs(zipBlob, 'documentos_gerados.zip');
                log("Download iniciado!");
            }

            if (nomesComErro.length > 0) {
                log("\nRegistros com erro encontrados:");
                nomesComErro.forEach(nome => log(`- ${nome}`));
            }

        } catch (error) {
            log(`ERRO GERAL: ${error.message}`);
            console.error(error);
            log("Processamento interrompido.");
        } finally {
            sendButton.disabled = false;
            sendButton.textContent = "Gerar Documentos";
        }
    });
});