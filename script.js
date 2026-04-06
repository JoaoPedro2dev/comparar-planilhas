const arcDesc1 = document.querySelector('#archive-desc-1');
const arcDesc2 = document.querySelector('#archive-desc-2');

const inputPlan1 = document.querySelector('#plan1');
const inputPlan2 = document.querySelector('#plan2');

const modalFileArray = document.querySelectorAll('.modal-file');

const messageError = document.querySelector('#error-msg');
const resultTable = document.querySelector('#modal-help');

const totalAmount = document.querySelector('#total-amount');

const exportBtn = document.querySelector('#exportExcelBtn');
const totalAmountBox = document.querySelector('#total-amount-box')

let exportData = [];

inputPlan1.addEventListener('change', () => {
    if (inputPlan1.files[0]) arcDesc1.textContent = inputPlan1.files[0].name
})
inputPlan2.addEventListener('change', () => {
    if (inputPlan2.files[0]) arcDesc2.textContent = inputPlan2.files[0].name
})

function showError(menssagem) {
    modalFileArray.forEach(modal => {
        modal.classList.add('border-text-red')
    })
    messageError.classList.remove('display-none');
    messageError.textContent = menssagem;
}

function hiddenError() {
    modalFileArray.forEach(modal => {
        modal.classList.remove('border-text-red')
    })
    messageError.classList.add('display-none');
}

function lerExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: "array" });

                const primeiraAba = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[primeiraAba];

                const json = XLSX.utils.sheet_to_json(worksheet, { header: "A" });

                resolve(json);
            } catch (erro_onload) {
                reject(erro_onload);
            }
        }

        reader.onerror = function (erro) {
            reject(erro);
        };

        reader.readAsArrayBuffer(file);
    });
}

function padronizarValor(valor) {
    return String(valor ?? "").trim().replace(/\s+/g, "").replace(/\.0$/, "");
}

function compararPlanilhas(plan1, plan2) {
    const resultado = [];
    const mapaPlan2 = new Map();

    //indexando planilha 2 pela coluna "A"
    plan2.forEach((linha2, index) => {
        const valorA = padronizarValor(linha2.A)

        if (valorA) {
            mapaPlan2.set(valorA, { ...linha2, __linhaPlan2: index + 1 });
        }
    });

    //Percorrendo planilha 1 para dar matchs com coluna A
    plan1.forEach(linha1 => {
        let valorA = padronizarValor(linha1.A);

        if (valorA.includes("-")) {
            valorA = padronizarValor(valorA.split("-")[1].trim());
        } else {
            valorA = padronizarValor(valorA.trim());
        }

        if (mapaPlan2.has(valorA)) {
            const linha2 = mapaPlan2.get(valorA);
            resultado.push({ valorComparado: valorA, dadosPlan2: linha2 });
        }
    });

    return resultado;
}

const startBtn = document.querySelector('#action-btn')
startBtn.addEventListener('click', async () => {
    const plan1 = document.querySelector('#plan1').files[0];
    const plan2 = document.querySelector('#plan2').files[0];

    if (!plan1 || !plan2) {
        showError('Selecione duas planilhas');
        return;
    }

    hiddenError();

    try {
        const dadosPlan1 = await lerExcel(plan1);
        const dadosPlan2 = await lerExcel(plan2);


        const resultado = compararPlanilhas(dadosPlan1, dadosPlan2);

        const contagem = resultado.reduce((acc, item) => {
            const nome = item.dadosPlan2.B;
            acc[nome] = (acc[nome] || 0) + 1;
            return acc;
        }, {})

        renderTable(contagem);

        resultTable.classList.remove('display-none')


    } catch (erro_planilhas) {
        showError('Algo deu errado ao analisar planilhas');
    }
})

function renderTable(contagem) {
    const tabela = document.querySelector('#dinamicTable');
    const thead = tabela.querySelector('thead');
    const tbody = tabela.querySelector('tbody');

    thead.innerHTML = "";
    tbody.innerHTML = "";
    totalAmount.textContent = 0;

    const dadosTabela = Object.entries(contagem).map(([nome, quantidade]) => ({
        nome,
        quantidade
    }));

    if (dadosTabela.length === 0) {
        tbody.innerHTML = `<tr><td colspan="2">Nenhum dado compatível</td></tr>`;

        totalAmountBox.classList.add('display-none')
        exportBtn.disabled = true;
        exportBtn.classList.add('display-none');

        return
    }

    totalAmountBox.classList.remove('display-none')
    exportBtn.disabled = false;
    exportBtn.classList.remove('display-none');

    exportData = dadosTabela;

    const colunas = Object.keys(dadosTabela[0]);

    const trHead = document.createElement("tr");

    colunas.forEach(coluna => {
        const th = document.createElement('th');
        th.textContent = coluna;
        trHead.appendChild(th);
    })

    thead.appendChild(trHead);

    dadosTabela.forEach(item => {
        const tr = document.createElement('tr');

        colunas.forEach(coluna => {
            const td = document.createElement('td');
            td.textContent = item[coluna];
            tr.appendChild(td);
        })

        tbody.appendChild(tr);
    })

    totalAmount.textContent = Object.values(contagem).reduce((acc, valor) => acc + valor, 0);
}

function exportExcel() {
    if (!exportData.length) {
        alert("Nenhum dado para exportar");
        return;
    }

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workBook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workBook, worksheet, "Respostas SSI por vendedor");
    XLSX.writeFile(workBook, "respostas_ssi_por_vendedor.xlsx");
}

exportBtn.addEventListener('click', exportExcel);