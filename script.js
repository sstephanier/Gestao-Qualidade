document.addEventListener("DOMContentLoaded", () => {
    const nextButton = document.getElementById("nextButton");
    const inspetorForm = document.getElementById("inspetorForm");
    const testesSection = document.getElementById("testes");
    const medicoesContainer = document.getElementById("medicoesContainer");
    const exportButton = document.getElementById("exportButton");

    const testes = [
        "Medição da bobina mãe",
        "Conferência do substrato",
        "Limpeza da primeira volta da bobina",
        "Uso do alinhador enroladeira",
        "Uso do alinhador desenroladeira",
        "Tensão na bobina",
        "Temperatura do rolo refrigerado",
        "Tratamento Corona potência",
        "Tratamento de Função",
        "Contragrafismo sujo",
        "Faca Raspadora do verniz",
        "Vida útil da lâmpada UV",
        "Vazamento do Grupo Impressor",
        "Alinhamento do material após o corte",
        "Aplicação do verniz (registro/res)",
        "Preparação do tubete",
        "Aspereza amostra"
    ];

    nextButton.addEventListener("click", () => {
        if (inspetorForm.checkValidity()) {
            inspetorForm.style.display = "none";
            testesSection.style.display = "block";
            generateInputs(testes, 100); // Gerar 100 campos para cada teste
        } else {
            alert("Por favor, preencha todos os campos antes de continuar.");
        }
    });

    function generateInputs(tests, count) {
        medicoesContainer.innerHTML = ""; // Limpa o contêiner antes de adicionar novos campos
        tests.forEach(test => {
            const section = document.createElement("div");
            const header = document.createElement("h3");
            header.textContent = test;
            section.appendChild(header);

            for (let i = 1; i <= count; i++) {
                const input = document.createElement("input");
                input.type = "text"; // Permite qualquer tipo de entrada
                input.name = `${test.replace(/[^a-zA-Z]/g, "_").toLowerCase()}_${i}`;
                input.placeholder = `${test}${i}`;
                section.appendChild(input);
            }
            medicoesContainer.appendChild(section);
        });
    }

    exportButton.addEventListener("click", () => {
        const inspetorData = {
            of: document.getElementById("of").value,
            codProd: document.getElementById("cod_prod").value,
            data: document.getElementById("data").value,
            turno: document.getElementById("turno").value,
            re: document.getElementById("re").value,
            inspetor: document.getElementById("inspetor").value
        };

        const sheetData = [];

        for (let i = 1; i <= 100; i++) { // Agora gera 100 linhas para exportação
            const rowData = {
                "OF": inspetorData.of,
                "CÓD. PROD": inspetorData.codProd,
                "DATA": inspetorData.data,
                "TURNO": inspetorData.turno,
                "RE": inspetorData.re,
                "NOME DO INSPETOR": inspetorData.inspetor,
            };

            testes.forEach(test => {
                rowData[test] = document.querySelector(`input[name='${test.replace(/[^a-zA-Z]/g, "_").toLowerCase()}_${i}']`).value || '';
            });

            sheetData.push(rowData);
        }

        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Testes");
        XLSX.writeFile(workbook, "testes_inspetor.xlsx");
    });
});