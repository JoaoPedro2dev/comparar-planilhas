
const arcDesc1 = document.querySelector('#archive-desc-1');
const arcDesc2 = document.querySelector('#archive-desc-2');

const inputPlan1 = document.querySelector('#plan1');
const inputPlan2 = document.querySelector('#plan2');

const selectBox = document.querySelectorAll(".select-box");
const selectPlan1 = document.querySelector("#col_planilha_1")
const selectPlan2 = document.querySelector("#col_planilha_2")
const planilhaRetornar = document.querySelector("#planilha_retornar");
const colunaRetornar = document.querySelector("#coluna_retornar");

const modalFileArray = document.querySelectorAll('.modal-file');

const messageError = document.querySelector('#error-msg');
const resultTable = document.querySelector('#modal-help');

const totalAmount = document.querySelector('#total-amount');

const exportBtn = document.querySelector('#exportExcelBtn');
const totalAmountBox = document.querySelector('#total-amount-box')

const startBtn = document.querySelector('#action-btn')

let exportData = [];

let colunasPlan1 = [];
let colunasPlan2 = [];

inputPlan1.addEventListener('change', async () => {
    if (!inputPlan1.files[0]) return;

    arcDesc1.textContent = formatFileName(inputPlan1.files[0].name);

    colunasPlan1 = await obterCabecalho(inputPlan1.files[0]);

    construirSelect(selectPlan1, colunasPlan1);

    selectBox[0].style.display = "flex";

    visibleRetornar();
})
inputPlan2.addEventListener('change', async () => {
    if (!inputPlan2.files[0]) return;

    arcDesc2.textContent = formatFileName(inputPlan2.files[0].name)

    colunasPlan2 = await obterCabecalho(inputPlan2.files[0]);

    construirSelect(selectPlan2, colunasPlan2);

    selectBox[1].style.display = "flex";

    visibleRetornar();
})

planilhaRetornar.addEventListener('change', () => {
    colunaRetornar.innerHTML = "";

    if (planilhaRetornar.value == 1) {

        colunasPlan1.forEach(coluna => {
            const option = document.createElement("option");

            option.value = coluna;
            option.textContent = coluna;

            colunaRetornar.appendChild(option);
        })
    }

    if (planilhaRetornar.value == 2) {

        colunasPlan2.forEach(coluna => {
            const option = document.createElement("option");

            option.value = coluna;
            option.textContent = coluna;

            colunaRetornar.appendChild(option);
        })
    }
})

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

        const colunaPlan1Selecionada = selectPlan1.value;
        const colunaPlan2Selecionada = selectPlan2.value;

        const colunaRetornoSelecionada = colunaRetornar.value;

        const origemRetornoSelecionada =
            planilhaRetornar.value;

        console.log('Planilha1', dadosPlan1);
        console.log('Planilha2', dadosPlan2);

        const resultado = compararPlanilhas(dadosPlan1, dadosPlan2, colunaPlan1Selecionada, colunaPlan2Selecionada, colunaRetornoSelecionada, origemRetornoSelecionada);

        const contagem = resultado.reduce((acc, item) => {
            const nome = item.valorRetornado;

            acc[nome] = (acc[nome] || 0) + 1;

            return acc;
        }, {})

        renderTable(contagem);

        resultTable.classList.remove('display-none')


    } catch (erro_planilhas) {
        showError('Algo deu errado ao analisar planilhas');
    }
})

exportBtn.addEventListener('click', exportExcel);

async function obterCabecalho(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: "array" });

                const primeiraAba = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[primeiraAba];

                const linhas = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1
                });

                resolve(linhas[0] || []);
            } catch (erro) {
                reject(erro);
            }
        };

        reader.readAsArrayBuffer(file);
    });
}

function formatFileName(fileName) {
    if (fileName.length > 13) {
        return fileName.slice(0, 13).trim() + '...'
    }

    return fileName.trim();
}

function visibleRetornar() {
    if (selectPlan1?.value != '' && selectPlan2?.value != '') {
        selectBox[2].style.display = "flex";
        selectBox[3].style.display = "flex";
    }
}

function construirSelect(select, colunas) {
    select.innerHTML = "";

    colunas.forEach(coluna => {
        const option = document.createElement("option");

        option.value = coluna;
        option.textContent = coluna;

        select.appendChild(option);
    })
}

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

                const json = XLSX.utils.sheet_to_json(worksheet);

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

function compararPlanilhas(plan1, plan2, colPlan1, colPlan2, colunaRetorno, origemRetorno) {
    const resultado = [];
    const mapaPlan2 = new Map();

    //indexando planilha 2 pela coluna "A"
    plan2.forEach((linha2, index) => {
        const valorComparacao = padronizarValor(linha2[colPlan2])

        if (valorComparacao) {
            mapaPlan2.set(valorComparacao, { ...linha2, __linhaPlan2: index + 1 });
        }
    });

    //Percorrendo planilha 1 para dar matchs com coluna A
    plan1.forEach(linha1 => {
        const valorComparacao = padronizarValor(linha1[colPlan1]);

        if (mapaPlan2.has(valorComparacao)) {
            const linha2 = mapaPlan2.get(valorComparacao);

            let valorRetornado;

            if (origemRetorno === "1") {
                valorRetornado = linha1[colunaRetorno];
            } else {
                valorRetornado = linha2[colunaRetorno];
            }

            resultado.push({
                valorComparado: valorComparacao,
                valorRetornado
            });
        }
    });

    return resultado;
}

